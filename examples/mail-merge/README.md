# Mail merge sample

This sample builds a mail merge on `@docx-editor.dev/core` and `@docx-editor.dev/react` only. Both packages use the Apache-2.0 license. One merge engine runs on a server through `createServerAutomationHost` and in the browser through `createBrowserAutomationHost`.

## Run the example

From the repository root, install dependencies, and then start the browser client:

```bash
bun install
bun run dev:mail-merge
```

Open `http://localhost:5181/`. Use the side panel to list variables, insert placeholders, create rules from a selection, and preview or download merged documents.

To verify the merge on a server without a browser, run the verification script:

```bash
bun run --filter './examples/mail-merge' verify
```

The script edits the template through the API, merges two records, reopens each result, and checks its text and XML. It writes the files to `examples/mail-merge/out/`.

## Template conventions

The merge engine in `src/merge/merge-engine.ts` recognizes these constructs:

| Construct | Template form | Merge result |
| --- | --- | --- |
| Text variable | `{{Name}}` in the body, a header, or a footer | Replaced with `values.Name`. Run formatting stays. |
| Variable control | Plain text content control tagged `var:Name` | Filled with `values.Name`, then unwrapped. |
| Rule | Content control tagged `if:flag` or `if:!flag` | Content kept or removed by `flags.flag`. |
| Repeating rows | Table row with `{{list.column}}` cells | One row for each item in `lists.list`. |
| Repeating section | Repeating section tagged `repeat:list`, item text with `{{list.column}}` | One item for each entry in `lists.list`, then unwrapped. |
| Word field | `MERGEFIELD Name` | Replaced with plain text `values.Name`. |
