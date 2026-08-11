# intake-list-to-markdown

An admin.autos plugin: reads a SharePoint intake list and writes its contents to a
single markdown file in a document library, overwriting the previous file. No user
interaction — meant to be called by the admin.autos runner.

Loads as a single script that puts one function on the window object:

```js
window.intakeListToMarkdown(config) // -> Promise<{ fileUrl, itemCount, groupCount }>
```

Requires PnPjs v2 (`pnp2.bundle.js`) to be loaded first so the global `pnp` exists.
Uses `pnp.sp.createIsolated({ baseUrl })`, so it works no matter which site the
hosting page is on.

## Config

| key | required | default | notes |
| --- | --- | --- | --- |
| `siteUrl` | yes | | site containing the list |
| `listTitle` | yes | | list title |
| `outputFolder` | yes | | full or server-relative URL of the destination folder |
| `outputSiteUrl` | no | `siteUrl` | set if the output library is on a different site |
| `fileName` | yes | | markdown file name, overwritten each run |
| `title` | yes | | H1 of the markdown (run timestamp is appended) |
| `top` | no | 50 | number of records to pull |
| `orderBy` | no | `"ID"` | internal name of the sort field |
| `ascending` | no | `false` | sort direction |
| `groupBy` | no | `"AssignedTo"` | internal name of the group-by field |
| `fields` | yes | | internal names of fields to output; the **first** one becomes the `### Item:` heading |

## Example

```js
intakeListToMarkdown({
  siteUrl: "https://nervedotnet.sharepoint.com/sites/NewNerve",
  listTitle: "Intake Test",
  outputFolder: "https://nervedotnet.sharepoint.com/sites/NewNerve/FCUPortal/Dev",
  fileName: "intake-test.md",
  title: "Intake Test Report",
  fields: ["Title", "Status", "TaskType", "DueDate", "Submitter", "Description", "RefLink"]
});
```

## Output shape

```markdown
# Title (as of 08/11/2026 10:36 AM)

* URL: https://.../Lists/Intake%20Test

---

## Assigned To: Benay Yocum | benay.yocum@nerve.digital

### Item: New vendor intake form

* Status: New
* Task Type: Feature; Admin
* Due Date: 09/01/2026 12:30 PM
* Reference Link: [Spec doc](https://example.com/spec/vendor-form)

---
```

Groups are sorted alphabetically; items keep the pull sort order (default ID desc).
Items with an empty group-by value land in a `(blank)` group.

## Field normalization

Everything is rendered as text:

- **Person** → `Display Name | email` (multi-person joined with `; `)
- **Choice** → as-is; **multi-choice** → `choice; choice; choice`
- **Date** → `MM/DD/YYYY HH:MM AM/PM` Eastern time
- **URL** → `[description](url)`
- **Multiline** → line breaks preserved, continuation lines indented to stay inside the bullet (rich text is stripped to plain text)
- **Lookup** → the looked-up display value (multi joined with `; `)
- **Yes/No** → `Yes` / `No`
- Empty values → blank

Field labels in the output are the SharePoint display names; `fields`, `groupBy`,
and `orderBy` in the config use internal names.

## Testing

Tested against the `Intake Test` list on the NewNerve dev site (created for this
purpose), outputting to `FCUPortal/Dev/intake-test.md`. Both are throwaway and can
be deleted.
