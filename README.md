# NX 2306 Exterior Export — website source

Private interactive configuration website for the single-body assembly exterior workflow.

## Run locally

1. Extract this ZIP to a folder on your computer.
2. Open a terminal in the extracted `nx2306-exterior-export` folder.
3. If Python is installed, run:

   ```sh
   python -m http.server 8000 --bind 127.0.0.1 --directory dist
   ```

4. Open http://localhost:8000 in your browser.
5. Press Ctrl+C in the terminal to stop the server.

On systems where the Python command is `python3`, use `python3` instead of `python`.
Serve the directory over HTTP because the website uses JavaScript modules. Opening `index.html` directly from the file system may prevent those modules from loading.

## Files

| File | Purpose |
| --- | --- |
| `dist/index.html` | Configuration form, review report and help dialog |
| `dist/style.css` | Responsive layout and visual styles |
| `dist/app.js` | UI interactions, local draft storage and downloads |
| `dist/config.js` | Defaults, validation, configuration schema and report generation |
| `tests/config.test.mjs` | Configuration rule and source checks |
| `tests/interactions.test.mjs` | DOM-harness interaction and download checks |
| `package.json` | Project metadata and test command |

The website has no third-party runtime dependencies and no build step. Python is used only to serve the files locally. Node.js is needed only to run the tests.

## Run tests

With Node.js installed, run:

```sh
node --test tests/config.test.mjs tests/interactions.test.mjs
```

The supplied version passed 17 automated checks. The interaction harness does not render CSS; browser visual QA was unavailable during the original build.

## What this version does

- Configure assembly scope, visible-component filtering and internal-component exclusion rules.
- Choose exterior surfaces or a solid target, extraction quality, opening closure, sewing and conditional solid creation.
- Configure STEP AP214/AP242, Parasolid text/binary output, recipient version targets and export-both.
- Review a live summary, planned filenames and an expected result report.
- Download the configuration as JSON and the expected report as text.
- Keep the draft in the browser's local storage.

## NX integration boundary

This is the website source, not an NXOpen journal. It has no connection to Siemens NX and does not read assembly geometry, create bodies, write STEP/Parasolid files or modify a CAD session. A compatible NX 2306 journal or utility must map the downloaded JSON plan to supported NX operations and produce actual exports.

The workflow targets one unassociated body in the existing part. Closure, sewing and solid creation depend on the source geometry; a single body is not guaranteed. Actual component counts, body counts and export statuses remain pending until NX runs. Recipient-specified Parasolid versions must be verified against the installed translator.

## Hosting and privacy

The original hosted website uses Sites access controls to restrict access to its owner. These static source files do not implement that hosting authentication. A local server bound to 127.0.0.1 is intended for use on the same computer. If you host a separate copy, configure its access controls separately.

The archive intentionally excludes hosting identity, credentials and Git history. No hosted site or external CAD system was changed to create this source package.

Source commit: `d5191b1e1709103f27d25fa8e01dadb96101ea2f`
