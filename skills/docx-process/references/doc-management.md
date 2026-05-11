# Document Management — Parameter Details

## create_document
Create a new Word document and save it to the specified path.
```json
{"file_path": "path/to/document.docx"}
```
- `file_path` (required) — path where the new document will be saved

## open_document
Open an existing Word document for editing. Detects if the file is locked by Microsoft Word.
```json
{"file_path": "path/to/document.docx"}
```
- `file_path` (required) — path to the .docx file to open

## save_document
Save the current document to its original file path. Includes retry logic (3 attempts, 2s interval) if the file is locked.
```json
{}
```

## save_as_document
Save the current document as a new file. The current file path reference is updated to the new path.
```json
{"new_file_path": "path/to/new.docx"}
```
- `new_file_path` (required) — the new file path to save to

## create_document_copy
Create a copy of the current document with a suffix appended to the filename. The original document remains open.
```json
{"suffix": "-副本"}
```
- `suffix` (optional, default: `"-副本"`) — suffix added before the file extension

## close_document
Close the current document without saving. Cleans up internal references.
```json
{}
```

## reload_document
Reload the current document from disk, discarding all unsaved changes.
```json
{}
```
