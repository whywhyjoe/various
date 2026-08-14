/*
 * intake-list-to-markdown.js
 *
 * Exposes window.intakeListToMarkdown(config) for use as an admin.autos plugin.
 * Reads a SharePoint list via PnPjs v2 (global `pnp` from pnp2.bundle.js),
 * renders the items as a grouped markdown document, and saves it to a document
 * library folder, overwriting any existing file. No user interaction.
 *
 * config:
 *   siteUrl       : URL of the site containing the list       (required)
 *   listTitle     : title of the list                         (required)
 *   outputFolder  : full or server-relative URL of the folder
 *                   the markdown file is saved into           (required)
 *   outputSiteUrl : site of the output folder                 (default siteUrl)
 *   fileName      : name of the markdown file                 (required)
 *   title         : heading at the top of the markdown        (required)
 *   top           : number of records to pull                 (default 50)
 *   orderBy       : internal name of the sort field           (default "ID")
 *   ascending     : sort direction                            (default false)
 *   groupBy       : internal name of the group-by field       (default "AssignedTo")
 *   fields        : internal names of the fields to output; the first one is
 *                   used as the "### Item:" heading           (required)
 *
 * Returns a Promise resolving to { fileUrl, itemCount, groupCount }.
 *
 * Example:
 *   intakeListToMarkdown({
 *     siteUrl: "https://nervedotnet.sharepoint.com/sites/NewNerve",
 *     listTitle: "Intake Test",
 *     outputFolder: "https://nervedotnet.sharepoint.com/sites/NewNerve/FCUPortal/Dev",
 *     fileName: "intake-test.md",
 *     title: "Intake Test Report",
 *     fields: ["Title", "Status", "TaskType", "DueDate", "Submitter", "Description", "RefLink"]
 *   }).then(function (r) { console.log(r.fileUrl); });
 */
(function () {
  "use strict";

  var EASTERN = new Intl.DateTimeFormat("en-US", {
    timeZone: "America/New_York",
    month: "2-digit", day: "2-digit", year: "numeric",
    hour: "2-digit", minute: "2-digit", hour12: true
  });

  var DATE_ONLY = new Intl.DateTimeFormat("en-US", {
    timeZone: "UTC",
    month: "2-digit", day: "2-digit", year: "numeric"
  });

  function fmtDate(value, fieldInfo) {
    // Date-only fields (DisplayFormat 0) keep their calendar date — converting
    // midnight to Eastern would shift them to the previous day. Date+time
    // fields convert to Eastern. -> MM/DD/YYYY [HH:MM AM/PM]
    var dateOnly = fieldInfo && fieldInfo.DisplayFormat === 0;
    return (dateOnly ? DATE_ONLY : EASTERN).format(new Date(value)).replace(",", "");
  }

  function asArray(value) {
    // multi-value fields arrive as [..] or { results: [..] } depending on odata mode
    if (!value) { return []; }
    return Array.isArray(value) ? value : (value.results || []);
  }

  function person(u) {
    return u.Title + " | " + (u.EMail || "");
  }

  function stripHtml(html) {
    var div = document.createElement("div");
    div.innerHTML = html;
    // textContent drops <br> and block-element boundaries; turn them into
    // newlines first so rich-text line breaks survive.
    Array.prototype.forEach.call(div.querySelectorAll("br"), function (br) {
      br.parentNode.replaceChild(document.createTextNode("\n"), br);
    });
    Array.prototype.forEach.call(div.querySelectorAll("p,div,li"), function (block) {
      block.appendChild(document.createTextNode("\n"));
    });
    return (div.textContent || "").replace(/\n{3,}/g, "\n\n").trim();
  }

  // Normalize any field value to markdown-safe text based on its field type.
  function normalize(value, fieldInfo) {
    if (value === null || value === undefined || value === "") { return ""; }
    switch (fieldInfo.TypeAsString) {
      case "User":
        return person(value);
      case "UserMulti":
        return asArray(value).map(person).join("; ");
      case "Lookup":
        return String(value[fieldInfo.LookupField || "Title"] || "");
      case "LookupMulti":
        return asArray(value).map(function (v) { return v[fieldInfo.LookupField || "Title"]; }).join("; ");
      case "MultiChoice":
        return asArray(value).join("; ");
      case "DateTime":
        return fmtDate(value, fieldInfo);
      case "URL":
        return "[" + (value.Description || value.Url) + "](" + value.Url + ")";
      case "Boolean":
        return value ? "Yes" : "No";
      case "Note":
        return stripHtml(String(value));
      default:
        return String(value);
    }
  }

  // Multiline values keep their line breaks but stay inside the bullet by
  // indenting continuation lines.
  function bulletText(text) {
    return text.replace(/\r\n|\r|\n/g, "\n    ");
  }

  // Headings must stay on one line; collapse any internal whitespace.
  function headingText(text) {
    return String(text).replace(/\s+/g, " ").trim();
  }

  function serverRelative(url) {
    return url.replace(/^https?:\/\/[^/]+/, "").replace(/\/$/, "");
  }

  window.intakeListToMarkdown = function (config) {
    if (!config) { throw new Error("config is required"); }
    ["siteUrl", "listTitle", "outputFolder", "fileName", "title"].forEach(function (name) {
      if (!config[name]) { throw new Error("config." + name + " is required"); }
    });
    if (!Array.isArray(config.fields) || config.fields.length === 0) {
      throw new Error("config.fields must be a non-empty array");
    }

    var top = config.top || 50;
    var orderBy = config.orderBy || "ID";
    var ascending = config.ascending === true;
    var groupBy = config.groupBy || "AssignedTo";
    var fields = config.fields;

    var fieldInfoByName = {};
    var listUrl = "";
    var list, items;

    return pnp.sp.createIsolated({ baseUrl: config.siteUrl })
      .then(function (sp) {
        list = sp.web.lists.getByTitle(config.listTitle);
        return list.fields.select("InternalName", "Title", "TypeAsString", "LookupField", "DisplayFormat")();
      })
      .then(function (allFields) {
        allFields.forEach(function (f) { fieldInfoByName[f.InternalName] = f; });
        return list.rootFolder.select("ServerRelativeUrl")();
      })
      .then(function (rootFolder) {
        listUrl = new URL(rootFolder.ServerRelativeUrl, config.siteUrl).href;

        // Build select/expand: person and lookup fields need projections.
        var select = ["ID"];
        var expand = [];
        fields.concat([groupBy]).forEach(function (name) {
          var info = fieldInfoByName[name];
          if (!info) { throw new Error("Field not found in list: " + name); }
          var type = info.TypeAsString;
          if (type === "User" || type === "UserMulti") {
            select.push(name + "/Title", name + "/EMail");
            expand.push(name);
          } else if (type === "Lookup" || type === "LookupMulti") {
            select.push(name + "/" + (info.LookupField || "Title"));
            expand.push(name);
          } else {
            select.push(name);
          }
        });

        var query = list.items.select.apply(list.items, select);
        if (expand.length) { query = query.expand.apply(query, expand); }
        return query.orderBy(orderBy, ascending).top(top)();
      })
      .then(function (fetched) {
        items = fetched;

        // Group items by the normalized group-by value, preserving sort order
        // within each group. Groups are ordered alphabetically. Null prototype:
        // a plain {} breaks on values like "constructor".
        var groups = Object.create(null);
        items.forEach(function (item) {
          var key = headingText(normalize(item[groupBy], fieldInfoByName[groupBy])) || "(blank)";
          if (!groups[key]) { groups[key] = []; }
          groups[key].push(item);
        });
        var groupKeys = Object.keys(groups).sort(function (a, b) { return a.localeCompare(b); });

        var md = [];
        md.push("# " + config.title + " (as of " + fmtDate(new Date()) + ")");
        md.push("");
        md.push("* URL: " + listUrl);
        md.push("");
        md.push("---");

        groupKeys.forEach(function (key) {
          md.push("");
          md.push("## " + fieldInfoByName[groupBy].Title + ": " + key);
          groups[key].forEach(function (item) {
            var heading = headingText(normalize(item[fields[0]], fieldInfoByName[fields[0]]));
            md.push("");
            md.push("### Item: " + heading);
            md.push("");
            fields.slice(1).forEach(function (name) {
              var info = fieldInfoByName[name];
              md.push("* " + info.Title + ": " + bulletText(normalize(item[name], info)));
            });
          });
        });

        md.push("");
        md.push("---");
        md.push("");

        // Save to the output folder, overwriting any existing file.
        var outSite = config.outputSiteUrl || config.siteUrl;
        return pnp.sp.createIsolated({ baseUrl: outSite })
          .then(function (outSp) {
            return outSp.web.getFolderByServerRelativeUrl(serverRelative(config.outputFolder))
              .files.add(config.fileName, md.join("\n"), true);
          })
          .then(function (result) {
            return {
              fileUrl: new URL(result.data.ServerRelativeUrl, outSite).href,
              itemCount: items.length,
              groupCount: groupKeys.length
            };
          });
      });
  };
})();
