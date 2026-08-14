# Code Review: `intake-list-to-markdown`

## Scope and verification

Reviewed `intake-list-to-markdown.js` and `README.md`, after reading the repository-level `CLAUDE.md`. Per the requested scope, this review deliberately excludes security analysis.

The script passes `node --check`. I also exercised it through a mocked PnPjs chain to validate grouping and configuration edge cases. There is no automated test suite in the project; the README records only a manual integration test against one SharePoint list.

**Summary:** 2 critical correctness issues, 5 suggestions, and several solid implementation practices.

## **🔴 Critical Issues** - Must fix before merge

### 1. Valid group values can crash the entire export

**Lines:** `intake-list-to-markdown.js:153-158`

`groups` is a normal object, so inherited property names are already present. If a configurable text/choice/lookup group field normalizes to `constructor`, `toString`, or a similar key, `groups[key] || []` returns the inherited value and `.push(item)` throws. I reproduced `constructor` as `TypeError: groups[key].push is not a function`. `__proto__` is another problematic key.

Use a null-prototype dictionary (or `Map`) and initialize buckets explicitly:

```js
var groups = Object.create(null);

items.forEach(function (item) {
  var key = normalize(item[groupBy], fieldInfoByName[groupBy]) || "(blank)";
  if (!groups[key]) { groups[key] = []; }
  groups[key].push(item);
});

var groupKeys = Object.keys(groups).sort(function (a, b) {
  return a.localeCompare(b);
});
```

This makes every SharePoint value a valid data key and also makes the stated alphabetical ordering locale-aware.

### 2. Date-only SharePoint fields can silently report the previous day

**Lines:** `intake-list-to-markdown.js:39-48, 80-81, 117`

All `DateTime` fields are parsed as instants and converted to Eastern time, but SharePoint distinguishes date-only fields (`DisplayFormat = 0`) from date-and-time fields (`DisplayFormat = 1`). A date-only-like REST value of `2026-09-01T00:00:00Z` currently renders as `08/31/2026 08:00 PM` in Eastern time. That silently changes the business date.

Fetch `DisplayFormat`, preserve the logical calendar date for date-only fields, and use Eastern conversion only for actual date-time fields:

```js
var DATE_ONLY = new Intl.DateTimeFormat("en-US", {
  timeZone: "UTC",
  month: "2-digit", day: "2-digit", year: "numeric"
});

function fmtDate(value, fieldInfo) {
  var date = new Date(value);
  if (isNaN(date.getTime())) {
    throw new Error("Invalid date value: " + value);
  }

  var formatter = fieldInfo && fieldInfo.DisplayFormat === 0
    ? DATE_ONLY
    : EASTERN;
  return formatter.format(date).replace(",", "");
}

// Include DisplayFormat when reading field metadata.
list.fields.select(
  "InternalName", "Title", "TypeAsString", "LookupField", "DisplayFormat"
)();

// In normalize:
case "DateTime":
  return fmtDate(value, fieldInfo);
```

Add an integration assertion against the tenant's actual OData mode, especially around DST boundaries. Correct date handling matters more than showing a uniform timestamp for a field that has no time component.

## **🟡 Suggestions** - Improvements to consider

### 1. Validate the public configuration before making network calls

**Lines:** `intake-list-to-markdown.js:103-108, 129, 146, 171`

The documented required fields are not enforced. Missing `fields` fails later at `fields.concat(...)`; an empty array produces blank item headings; `top: 0` unexpectedly becomes `50` because of `config.top || 50`; and a bad `orderBy` is not checked with the other field names. These failures are late and harder for an automated runner to diagnose.

Validate once at the API boundary and default only when a value is absent:

```js
function validateConfig(config) {
  if (!config || typeof config !== "object") {
    throw new TypeError("config must be an object");
  }

  ["siteUrl", "listTitle", "outputFolder", "fileName", "title"].forEach(function (name) {
    if (typeof config[name] !== "string" || !config[name].trim()) {
      throw new TypeError(name + " must be a non-empty string");
    }
  });

  if (!Array.isArray(config.fields) || config.fields.length === 0) {
    throw new TypeError("fields must contain at least the item-heading field");
  }

  var top = config.top === undefined ? 50 : config.top;
  if (typeof top !== "number" || top % 1 !== 0 || top < 1) {
    throw new TypeError("top must be a positive integer");
  }
  return top;
}
```

Also include `orderBy` in schema validation, and either reject unsupported complex field types or document exactly which types are supported. Failing fast gives callers actionable errors without spending SharePoint requests first.

### 2. Separate text normalization from Markdown serialization

**Lines:** `intake-list-to-markdown.js:66-97, 155, 161-177`

The comment says normalized values are "markdown-safe," but values are inserted without escaping. A title, group, display name, field label, URL description, or item heading containing Markdown punctuation can change formatting. Newlines in the first field or group field can break the heading structure entirely. This is a rendering-correctness issue even when every input is legitimate SharePoint content.

Keep normalized plain text separate, then escape according to its output context:

```js
function escapeInline(text) {
  return String(text).replace(/([\\`*_[\]<>#~])/g, "\\$1");
}

function headingText(text) {
  return escapeInline(text).replace(/\s+/g, " ").trim();
}

md.push("# " + headingText(config.title) + " (as of " + fmtDate(new Date()) + ")");
md.push("## " + headingText(fieldInfoByName[groupBy].Title) + ": " + headingText(key));
md.push("### Item: " + headingText(heading));
```

For URL fields, escape the link label separately and normalize/encode the URL before composing the Markdown link. For bullet bodies, escape each line before adding the four-space continuation indentation. Context-specific serialization makes the output stable and keeps `normalize` focused on SharePoint types.

### 3. Preserve structural line breaks when stripping rich text

**Lines:** `intake-list-to-markdown.js:60-64, 86-87`; `README.md:78`

`textContent` does not turn `<br>` or adjacent block elements into newline characters. For example, rich text such as `First<br>Second` becomes `FirstSecond`, contrary to the README's promise that multiline line breaks are preserved.

Convert structural elements to text newlines before reading `textContent`:

```js
function stripHtml(html) {
  var div = document.createElement("div");
  div.innerHTML = html;

  Array.prototype.forEach.call(div.querySelectorAll("br"), function (br) {
    br.parentNode.replaceChild(document.createTextNode("\n"), br);
  });
  Array.prototype.forEach.call(div.querySelectorAll("p,div,li"), function (block) {
    block.appendChild(document.createTextNode("\n"));
  });

  return (div.textContent || "")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
```

Exercise this with `<br>`, paragraphs, lists, nested blocks, and already-plain multiline text. Preserving boundaries prevents distinct sentences and list entries from being merged.

### 4. Define the retrieval limit and paging behavior explicitly

**Lines:** `intake-list-to-markdown.js:104, 144-146`; `README.md:27`

The README describes `top` as the number of records to pull, while the implementation performs one page request and accepts any value. For the default `50` this is efficient, but large values interact with SharePoint page/list-view limits and may fail or return fewer rows than the caller expects. A non-indexed custom `orderBy` can also cause threshold failures on large lists.

Either enforce and document a supported maximum, or page until the requested count is reached:

```js
function takeItems(query, limit) {
  var rows = [];

  return query.top(Math.min(limit, 5000)).getPaged().then(function visit(page) {
    rows = rows.concat(page.results);
    if (rows.length >= limit || !page.hasNext) {
      return rows.slice(0, limit);
    }
    return page.getNext().then(visit);
  });
}
```

Keep the default `ID` ordering, and document that custom sort fields should be indexed for large lists. This makes `itemCount` reliably mean what the API contract implies.

### 5. Add repeatable unit/contract tests and document failure behavior

**Lines:** `README.md:86-90`; project-wide (no test files or test command)

The one manual SharePoint run covers the happy path but cannot prevent regressions in pure normalization/rendering logic or the fluent PnPjs query contract. The two critical issues above are both small edge cases that automated tests would catch without a live tenant.

Add dependency-light Node tests using `node:test`, `vm`, a minimal DOM stub (or a browser DOM test for rich text), and a mocked PnPjs chain. At minimum cover:

```js
test("groups prototype-named values", async function () {
  var run = await runPlugin({ groupValue: "constructor" });
  assert.match(run.markdown, /## Group: constructor/);
});

test("keeps a date-only value on its calendar day", function () {
  assert.equal(normalizeDate("2026-09-01T00:00:00Z", { DisplayFormat: 0 }), "09/01/2026");
});

test("rejects an empty fields array before calling PnPjs", async function () {
  await assert.rejects(callPlugin({ fields: [] }), /fields must contain/);
  assert.equal(mockPnp.callCount, 0);
});
```

Also test every documented field type, blank groups, duplicate `groupBy`/`fields` projections, DST dates, empty result sets, upload overwrite arguments, cross-site output, upload rejection, and exact returned counts/URL. Document that the returned promise rejects on list/schema/query/upload errors, and wrap low-level failures with list/folder/stage context while preserving the original error as `cause`.

## **✅ Good Practices** - What's done well

- **Clear, narrow public API (`intake-list-to-markdown.js:103-200`):** The script exposes one promise-returning function and returns useful operational metadata (`fileUrl`, `itemCount`, and `groupCount`). That is a good fit for a no-UI automation runner.
- **Source and destination concerns are separated (`intake-list-to-markdown.js:114-124, 186-195`):** Independent PnPjs contexts support cross-site output without relying on whichever site hosts the script.
- **The item query is economical for the normal case (`intake-list-to-markdown.js:126-146`):** It projects only required item properties, expands person/lookup fields deliberately, orders server-side, and applies a small default limit rather than downloading the entire list.
- **Normalization is centralized (`intake-list-to-markdown.js:50-91`):** Person, lookup, multi-choice, Boolean, URL, note, and date handling live behind one function instead of being duplicated throughout rendering.
- **Grouping preserves item order (`intake-list-to-markdown.js:151-158`):** Items retain the server query order within each group, matching the README's documented behavior.
- **Documentation is unusually helpful for a single-file plugin (`README.md:3-84`):** The configuration table, example, output sample, return shape, field labels, timezone, and normalization rules give integrators a strong starting contract.
