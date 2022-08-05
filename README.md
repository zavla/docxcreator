# Docxcreator - creates docx documents with supplied data from your docx templates. #

Merges user supplied data into the document using Word MergeFields.

And also adds rows to any table in a document.

The 'template' directory holds your docx files used as your templates for new documents.

Tables in a template is searched by the presence of text "1" "2" in its first two cells. The row that contains this text will be replaced with the user's JSON input data. Rows are added for every input JSON array element.

Uses "baliance.com/gooxml/document".


Example of POST request:
```
curl.exe  --data '@./testdata/пример_Input_JSON.json' --output ./testdata/output.docx  http://127.0.0.1:8080/docxcreator
```

Example of input JSON:
```
	{
		"DocxTemplateName": "yourTemplateFile.docx",
		"ShowFields": false,
		"Header": {
			"Номер": "ЮХ000000084",
			"Дата": "19.08.2021 11:31:20",
			"КлиентПолноеНаименование": "",
			"МестоСоставления": "",
			"НомерДоговора": "",
			"ДатаДоговора": "",
			"ИтогоСумма": "645,26"
		},
		"Table1": [
			{
				"1": "1",
				"2": "Ремкомплект MM для редуктора СО2 Premium",
				"3": "шт",
				"4": "2",
				"5": "322,63",
				"6": "645,26"
			}
		]
	}
```

May be started as a Windows service. Data is recieved with POST method as a JSON.

Logs into Windows event log and as an option to a dedicated log file.

Also it may be run as standalone executable in CLI mode.
