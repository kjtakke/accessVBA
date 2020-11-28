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
'	Public Constants:
'		bootstrapCSS
'		bootstrapJS
'		chartsCSS
'		chartsJS
'		jQuery
'		fontsAwsomeCSS
'		fontsAwsomeJS
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


'Public Constants:

public const bootstrapCSS as string = "<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css'>"
public const bootstrapJS  as string = "<script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js'></script>"
public const chartsJS  as string = "<script src='https://cdn.jsdelivr.net/npm/chart.js@2.8.0'></script>"
public const jQuery  as string = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'>"
public const fontsAwsomeCSS  as string = "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css'>"
public const googleapis as string = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'></script>"
public const cloudflare as String = "<script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js'></script>"

'Public Variables:

Public HTML_Array as Variant
Public HTML_Column_Count as Integer
Public HTML_Row_Count as Integer
Public HTML_Settings as Variant
Public HTML_Dictionary as Dictionary
Public HTML_Script as String
Public HTML_Style as String
public HTML_File_Name as String
public HTML_File_Path as Strng
public HTML_Elements_Count as Integer

Public Current_Colors as Variant
Public Current_Icon as Variant
public Current_SQL as string
public Current_Array as Variant
public Current_Row as Integer
public Current_Column as Integer

Public h as single
Public i as single
Public j as single
Public k as single


'Enums:

Public Enum tableClasses
	borderless
	hover
	striped
	border
End Enum

Public Enum chartType
	line
	area
	bar
	hBar
	pie
	donut
End Enum

Public Enum headings
	H1
	H2
	H3
	H4
	H5
	H6
End Enum

Public Enum headings
	Success
	Info
	Warning
	Danger
	Primary
	Secondary
	Dark
	Light
End Enum

Public Enum headings
	true
	false
End Enum

Public Enum colors

End Enum

Public Enum icons

End Enum


'Public Subs:

	Sub HTML_dimentions(rows as integer, columns as integer, optional fileName as string, optional filepath as string)
		If Len(fileName) = 0 then fileName = "Report"
		If Len(fileName) = 0 then filepath = "C:\Users\" & Environ("Username") & "\Desktop\"
		fileName = fileName & ".html"
		filepath = filepath & fileName
		HTML_Column_Count = Columns
		HTML_Row_Count = rows
		HTML_File_Name = fileName
		HTML_File_Path = filepath
		Redim HTML_Array(1 to rows, 1 to columns)
	End Sub

	'Add Elements
			Sub add_table()

			End Sub

			Sub add_metric()

			End Sub

			Sub add_chart()

			End Sub

			Sub add_heading()

			End Sub

			Sub add_div()

			End Sub

			Sub add_styleLink()

			End Sub

			Sub add_style()

			End Sub

			Sub add_scriptTopLink()

			End Sub

			Sub add_scriptTop()

			End Sub

			Sub add_scriptBottom()

			End Sub

			Sub add_scriptBottomLink()

			End Sub

		'Compile
				Sub export()

				End Sub

				Sub compile_and_export()

				End Sub

		'Other
				Sub to_Clipboard()

				End Sub


'Public Functions

	Public Function SQL_to_array(SQL as string) as Variant
		Dim o_rst As DAO.Recordset
		Dim a_SQL as variant
		Dim a_varField As Variant
		Dim i_DimCount as Integer

		Set o_rst = CurrentDb.OpenRecordset(SQL)

		'Set Array Dimentions
		o_rst.MoveLast
		ReDim a_SQL(0 to o_rst.RecordCount, 0 to o_rst.Fields.count)
		o_rst.MoveFirst

		'Add Filed Headers To Array
		For i = 0 To o_rst.Fields.count
				a_SQL(0,i) = o_rst.Fields(i)
		Next i

		'SQL Body to VBA Array
		Do While Not o_rst.EOF
		    For Each a_varField In o_rst.Fields
		    a_SQL(o_rst.AbsolutePosition + 1, a_varField.OrdinalPosition) = a_varField
		    Next a_varField
		    o_rst.MoveNext
		Loop

		'Get/Confirm Array Dimentions
		i_DimCount = arrayDimentionCounter(a_SQL)

		'Set Public Variables
		Current_SQL = SQL
		Current_Array = a_SQL

		'Return Array
		SQL_to_array = a_SQL

	End Function

	Public Function SQL_to_JS_array(SQL as string) as Variant

	End Function

	Public Function SQL_to_json(SQL as string) as String

	End Function

	Public Function compile_to_string() as String

	End Function

	Public Function compile_to_array() as Variant

	End Function




'Private Functions:

	Private Function pv_dimentionCount(index as Variant) as Integer

	End Function

	Private Function pv_chartTemplate(index as Variant) as String

	End Function

	Private Function pv_HTMLTemplate(index as Variant) as String

	End Function

	Private Function pv_styleTagTemplate(index as Variant) as String

	End Function

	Private Function pv_scriptTagTemplate(index as Variant) as String

	End Function

	Private Function pv_styleTemplate(index as Variant) as String

	End Function

	Private Function pv_scriptTemplate(index as Variant) as String

	End Function

	Private Function pv_scriptTemplate(index as Variant) as String

	End Function

	Private Function pv_icon_Template(index as Variant) as String

	End Function

	Private Function pv_metric_Template(index as Variant) as String

	End Function
