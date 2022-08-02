# Docxcreator is an http service that creates docx documents from your docx templates. #

May be started as a Windows service.
Has 'template' directory with docx templates.

Merges user supplied data using MergeFields.
Adds rows to a table data in a document.
Data is recieved with POST method, expects a json.

Uses "baliance.com/gooxml/document".
Writes into Windows event log and as on option to a dedicated log file.


```
Help message: a service expects a JSON input in the following form
	{
		"DocxTemplateName": "youtemplatenamehere.docx",         //nonempty - a document template name
		"ShowFields": false, 									//help if you need a list of available MergeFields in document template
		"Header": {
			"Номер": "ЮХ000000084",								// these are MergeFields in the document template
			"Дата": "19.08.2021 11:31:20",                      // they are shown as example
			"КонтрагентПолноеНаименование": "?",
			"МестоСоставления": "?",
			"НомерДоговора": "?",
			"ДатаДоговора": "?",
			"ИтогоСумма": "645,26"
		},
		"Table1": [
			{
				// text 1, 2 ... in columns in you template document are just text, not MergeFields, in any row of any table in the document template
				"1": "1",
				"2": "Ремкомплект MM для редуктора СО2 Premium",
				"3": "шт",
				"4": "2",
				"5": "322,63",
				"6": "645,26"
			}
            // you may specify more rows as an input
		]
	}
	You must specify DocxTemplateName.
	You may use the following DocxTemplateName values:
	```

May be run as standalone executable in command line mode.
