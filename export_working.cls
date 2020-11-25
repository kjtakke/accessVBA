'EXPORT CLASS
'This class is designed to create web dashboards in Microsoft Access

'REQUIRED PACKAGES
'	Visual Basic For Applications
'	Microsoft Access 16.0 Object Library
'	OLE Automation
'	System
'	Microsoft ActiveX Data Objects 6.1 Library
'	Microsoft ActiveX Data Objects Recordset 6.0 Library
'	Microsoft ADO Ext. 6.0 for DDL and Security
'	Microsoft Data Access Components Installed Version
'	Microsoft ADO 3.6 Object Library
'	Microsoft Outlook 16.0 Object Library
'	Microsoft Forms 2.0 Object Library (Browse for FO20.DLL)

'NAMING COVENTIONS
'Private Functions and Subs: _finctionName
'Public Functions and Subs: finctionName

'ELEMENTS
'Properties:
'	init_HTML
'	add_Element
'	export_HTML_File
'	preview
'	email
'	csv

'Public Functions:
'	table() as Variant
'	metric() as Variant
'	chart() as Variant
'	heading() as Variant
'	div() as Variant
'	styleLink() as Variant
'	style() as Variant
'	scriptTopLink() as Variant
'	scriptTop() as Variant
'	scriptBottom() as Variant
'	scriptBottomLink() as Variant
'	export() as Variant

'Public Subs:
'	to_Clipboard()
'	SQL_to_JS_array()
'	SQL_to_json()

'Private Functions:
'	_SQL_to_array() as Variant
'	_dimentionCount() as Integer
'	_chartTemplate() as String
'	_HTMLTemplate() as String
'	_styleTagTemplate() as String
'	_scriptTagTemplate() as String
'	_styleTemplate() as String
'	_scriptTemplate() as String

'Private Subs:
'	_barChart()
'	_hBarChart()
'	_pieChart()
'	_lineChart()
'	_table()
'	_div()
'	_metric()
'	_heading()

'Enums
'	fonts
'	tableClasses
'	chartType
'	headings
'	metricClasses
'	true_false
'	colors