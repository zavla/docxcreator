# docxcreator is an http service that creates and fills docx template documents with your json data and serves a document back

May be started as a Windows service.

Fills document templates with your data.
Fills header and a table in a document.
Data is recieved through http listener, POST method, expects a json.

May be run in command line mode.

Uses "baliance.com/gooxml/document".
Writes into Windows event log and dedicated log file.
