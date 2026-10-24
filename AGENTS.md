# AGENTS.md

## Project and communication
- GC AutoLoad is an Excel VBA add-in for normalizing order input and exporting XML.
- Explain work and results to the user in Thai. Keep existing code identifiers and naming conventions.
- Treat this file's directory as the repository root. Use git rev-parse --show-toplevel to confirm the active repository, especially when folders are nested.
- Inspect current files and git status before making changes. Preserve unrelated user changes.
- Work independently for routine authorized edits. Follow the user's current scope and environment approval requirements.

## Current repository
- src/ currently contains modBarcode.bas, modClipboard.bas, modConfig.bas, modLoader.bas, modMain.bas, modParser.bas, modVersion.bas and modXML.bas.
- Read implementations before assuming module responsibilities from their names.
- README.md describes the project. CHANGELOG.md at the root is the canonical changelog.
- The empty src/CHANGELOG.md is a leftover placeholder, not a second release history.
- modVersion.bas currently declares version 3.3 and release date 2026-10-07. Recheck it rather than assuming these values remain current.
- Do not assume a complete automated export/import/build pipeline exists. Document missing inputs and tools honestly.

## VBA changes and file encoding
- Keep one standard VBA Module per .bas file and preserve its existing name.
- Keep class modules as .cls and UserForms as paired .frm/.frx files when present. Do not discard .frx resources.
- Preserve exported Attribute metadata, public procedure names, Ribbon callback names, signatures and workbook integration points.
- Some source files may be pasted code rather than native VBE exports. Check file headers and module names before promising automatic import.
- Detect each existing source file's encoding and line endings before editing. Preserve Thai and other Unicode text; do not convert all VBA files to UTF-8 blindly.
- Validate any required encoding conversion with an Excel import/export round trip using synthetic text.
- Preserve existing Option Explicit declarations and VBA compatibility. Do not add external dependencies without a demonstrated need.
- Preserve conditional compilation and PtrSafe/LongPtr declarations when editing Windows API calls.
- Make focused changes. Do not split or rewrite the parser as part of repository setup.

## Data and configuration
- Never commit actual order data, customer names, addresses, phone numbers, raw production inputs, passwords, access tokens, private keys or secrets.
- Use synthetic fixtures only. Inspect staged content; .gitignore alone does not detect sensitive text inside source code or documents.
- Keep machine-specific server paths, deployment destinations and credentials in ignored local configuration. Use examples without actual internal values in tracked documentation.
- Respect .gitignore. Do not force-add ignored data or add-in binaries as a shortcut.
- Preserve original workbooks and installed add-ins. Do not overwrite network releases or the installed production add-in during development.
- Treat external files as references unless the user has authorized editing those files.

## Ribbon and workbook objects
- Ribbon XML and its images belong in ribbon/ when they have been extracted and verified. modXML.bas is VBA code, not the Ribbon UI definition.
- Preserve Ribbon package relationships, callback names and referenced image IDs.
- ThisWorkbook and Worksheet code must be restored to the matching existing document modules. Do not import these exports as ordinary class modules.
- A collection of .bas files alone may not reproduce the original add-in. Check workbook sheets, CodeNames, named ranges, references, forms, Ribbon parts and required configuration.
- Do not fabricate missing Ribbon XML or workbook event handlers.

## Export, import and build workflow
1. Keep an untouched backup of the authoritative add-in and identify the exact version and path.
2. Export or reconcile VBA components, workbook object code, forms, references and Ribbon resources with the tracked source.
3. Edit tracked source in VS Code.
4. Import into a separate development copy of the add-in, preserving document modules and supporting workbook structure.
5. Compile the VBA project and test it in Excel.
6. Build a separate release artifact only after successful checks. Keep development and release add-ins distinct.
- These steps are the intended workflow, not a claim that scripts implementing them already exist.
- Do not run macros or open unreviewed workbooks merely to inspect source.
- If automation needs Trust access to the VBA project object model, explain the requirement. Do not silently weaken Excel security settings.
- Editing VBA invalidates an existing VBA signature; any required signing belongs after the final build.

## Validation
- For documentation-only changes, verify paths, statements and formatting; no Excel run is required.
- For VBA changes, compile and exercise the affected behavior in desktop Excel when available.
- For quantity validation changes, cover matching and mismatching explicit quantities, input with no explicit quantity, duplicate Order ID + SKU rows and different SKUs in the same order.
- Test Thai and representative Unicode text and XML escaping when changing text or XML handling.
- Check Ribbon callbacks and version labels when changing Ribbon or version integration.
- Use an isolated test Excel session/add-in; avoid loading development and production copies together.
- Report checks actually performed and any checks still pending. Do not present source inspection as successful Excel testing.

## Changelog and releases
- Add new work under root CHANGELOG.md -> Unreleased.
- Describe user-visible outcomes and development workflow changes accurately.
- Do not invent release dates, version history, completed builds or test results.
- Do not increment version constants solely for repository setup or documentation changes.
- On an authorized release, keep modVersion.bas, changelog, release artifact metadata and any visible version labels consistent.
- Store reviewed distributable binaries in release assets rather than the source history, subject to the user's release instructions.

## Git and deployment
- Review git diff and staged files before committing. Stage explicit paths.
- Do not use force push, destructive reset, deletion, history rewriting or moves to fix a repository layout without explicit authorization.
- Do not modify an outer or sibling repository when working in this repository.
- Do not commit, push, publish, tag, install or deploy unless the user's request authorizes that action.
- Never run an updater as part of routine source editing or documentation work.

