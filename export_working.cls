'EXPORT CLASS
'This class is designed to create web dashboards in Microsoft Access
'
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
'
'RESOURCES
'	Charts.JS
'	Bootstrap 4
'	HTML 5
'	fontsAwsome
'
'REFRENCES
'	
'
'
'NAMING COVENTIONS
'	Private: _finctionName
'	Public: finctionName
'
'	Private Variables:
'		variant: a_VariableName
'		string: s_VariableName
'		integer: i_VariableName
'		boolean: b_VariableName
'		double: d_VariableName
'		collection: c_VariableName
'		dictionary: dict_VariableName
'		object: o_VariableName:
'			o_file
'			o_fileInstance
'			o_mail
'		Counters:
'			_i_, j_, _k_, _h_ as Single
'
'ELEMENTS
'	Properties:
'		init_HTML
'		add_Element
'		export_HTML_File
'		preview
'		email
'		csv
'		compile
'
'	Public Functions:
'		table() as Variant
'		metric() as Variant
'		chart() as Variant
'		heading() as Variant
'		div() as Variant
'		styleLink() as Variant
'		style() as Variant
'		scriptTopLink() as Variant
'		scriptTop() as Variant
'		scriptBottom() as Variant
'		scriptBottomLink() as Variant
'		export() as Variant
'		compile_to_string() as String
'		compile_to_array() as Variant
'		compile_and_store() as Variant
'		compile_and_export() as Variant
'		compile_and_export_all() as Variant
'		export_all() as Variant
'		export_key() as Variant
'		export_variable() as Variant
'
'	Public Subs:
'		to_Clipboard()
'		SQL_to_JS_array()
'		SQL_to_json()
'		
'	Private Functions:
'		_SQL_to_array() as Variant
'		_dimentionCount() as Integer
'		_chartTemplate() as String
'		_HTMLTemplate() as String
'		_styleTagTemplate() as String
'		_scriptTagTemplate() as String
'		_styleTemplate() as String
'		_scriptTemplate() as String
'
'	Private Subs:
'		_barChart()
'		_hBarChart()
'		_pieChart()
'		_lineChart()
'		_table()
'		_div()
'		_metric()
'		_heading()
'
'	Enums:
'		fonts
'		tableClasses
'		chartType
'		headings
'		metricClasses
'		true_false
'		colors
'		icons
'
'	Public Variables:
'		HTML_Array as Variant
'		HTML_Column_Count as Integer
'		HTML_Row_Count as Integer
'		HTML_Settings as Variant
'		HTML_Dictionary as Dictionary
'		HTML_Script as String
'		HTML_Style as String
'
'	Private Constants:
'		_bootstrapCSS
'		_bootstrapJS
'		_chartsCSS
'		_chartsJS
'		_jQuery
'		_fontsAwsomeCSS
'		_fontsAwsomeJS
'
'COLORS
'	black: #000000	rgb(0,0,0)
'	
'COLOR SETS
'	default: [red, green, ...]
'	...: [..., ..., ...]
'
'FONTS
'
'
'ICONS
'
'