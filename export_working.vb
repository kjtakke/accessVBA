'EXPORT CLASS
'This class is designed to create web dashboards in Microsoft Access
'
'IMPORT CLASS
'
'
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
'	Private: pv_finctionName
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
'			i, j, k, h as Single
'
'ELEMENTS
'	Properties:
'
'
'	Public Functions
'		SQL_to_array() as Variant
'		SQL_to_JS_array() as Variant
'		SQL_to_json() as String
'
'
'	Public Subs:
'		HTML_dimentions()
'		add_table()
'		add_metric()
'		add_chart()
'		add_heading()
'		add_div()
'		add_styleLink()
'		add_style()
'		add_scriptTopLink()
'		add_scriptTop()
'		add_scriptBottom()
'		add_scriptBottomLink()
'		export()
'		compile_to_string()
'		compile_to_array()
'		compile_and_store()
'		compile_and_export()
'		compile_and_export_all()
'		export_all()
'		export_key()
'		export_variable()
'		to_Clipboard()
'		
'	Private Functions:
'		pv_dimentionCount() as Integer
'		pv_chartTemplate() as String
'		pv_HTMLTemplate() as String
'		pv_styleTagTemplate() as String
'		pv_scriptTagTemplate() as String
'		pv_styleTemplate() as String
'		pv_scriptTemplate() as String
'		pv_icon_Template() as String
'		pv_metric_Template as String
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
'		HTML_File_Path as String
'		HTML_File_Name as String
'		Colors as Variant
'		Fonts as Variant
'		Icons as Variant
'
'	Private Constants:
'		pv_bootstrapCSS
'		pv_bootstrapJS
'		pv_chartsCSS
'		pv_chartsJS
'		pv_jQuery
'		pv_fontsAwsomeCSS
'		pv_fontsAwsomeJS
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
'
'EXAMPLES
'
'Sub btn_click()
'
'	Const SQL_Sales = "SELECT ..."
'	Const SQL_Metrics = "SELECT ..."
'	Const MH = "My Page Heading"
'
'	Dim d as export
'	Set d = New export
'
'	'Set HTML Document's rows and columns
'	d.HTML_dimentions(5,5)
'
'	'Add HTML Elements
'	d.add_heading(2,1,1,MH)
'	d.add_table(2,1,SQL_Sales)
'	d.add_chart(2,2,"pie",SQL_Metrics)
'
'	'Compile and export HTML Document
'	d.compile_and_export()
'
'End Sub
