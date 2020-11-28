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
' jQuery
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
'		String: s_VariableName
'		Integer: i_VariableName
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
'COLORS
'	black: #000000	rgb(0,0,0)
'
'COLOR SETS
'	default: [red, green, ...]
'	...: [..., ..., ...]
'
'ICONS
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

public const bootstrapCSS as String = "<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css'>"
public const bootstrapJS  as String = "<script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js'></script>"
public const chartsJS  as String = "<script src='https://cdn.jsdelivr.net/npm/chart.js@2.8.0'></script>"
public const jQuery  as String = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'>"
public const fontsAwsomeCSS  as String = "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css'>"
public const googleapis as String = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'></script>"
public const cloudflare as String = "<script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js'></script>"


'Public Variables:

Public HTML_Array as Variant
Public HTML_Column_Count as Integer
Public HTML_Row_Count as Integer
Public HTML_Script as String
Public HTML_Style as String
Public HTML_Script_Top_Links as String
Public HTML_Script_Bottom_Links as String
Public HTML_Style_Links as String
public HTML_File_Name as String
public HTML_File_Path as Strng
public HTML_Elements_Count as Integer
public HTML_Title as String
public HTML_Heading as string
Public HTML_Heaader as boolean

Public Current_Colors as Variant
Public Current_Icon as Variant
public Current_SQL as String
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

	Sub HTML_Setup(rows as Integer, columns as Integer, optional fileName as String, optional filepath as String, optional heading as string, optional title as string)
		'This Sub sets the dimentsions for the HTML Document

		'Optional Arguments
		If Len(fileName) = 0 then fileName = "Report"
		If Len(fileName) = 0 then filepath = "C:\Users\" & Environ("Username") & "\Desktop\"

		'Load Public Variables
		If Len(heading) = 0 then
			HTML_Heaader = False
		Else
			HTML_Heaader = true
		End if
		If Len(title) = 0 then HTML_Title = "Report"
		fileName = fileName & ".html"
		filepath = filepath & fileName
		HTML_Column_Count = columns
		HTML_Row_Count = rows
		HTML_File_Name = fileName
		HTML_File_Path = filepath
		HTML_Script = ""
		HTML_Style = ""
		HTML_Elements_Count = 0
		HTML_Heading = heading
		HTML_Style = ""
		HTML_Script = ""
		HTML_Style_Links = ""
		HTML_Script_Top_Links = ""
		HTML_Script_Bottom_Links = ""
		'Set HTML_Array Dimentsions (BASE 1)
		Redim HTML_Array(1 to rows, 1 to columns)

		'Set each array element to an empty string
		For i = 1 to rows
			For j = 1 to columns
				HTML_Array(i,j) = ""
			Next j
		Next i

	End Sub

	'Add Elements
			Sub add_table()

			End Sub

			Sub add_metric()

			End Sub

			Sub add_chart()

			End Sub

			Sub add_heading(row as integer, column as integer, heading_tag as Headings, heading_Text as String, Optional heading_style as string, Optional heading_class as string optional heading_id as string)
				'This Sub creates a <H1-6> Tag
				Dim s_heading_text as String
				Dim s_open_Tag as String
				Dim s_close_tag as String

				'Tag
				Select Case true
					Case heading_tag = Headings.h1
						s_open_Tag = "<h1 "
						s_close_tag = "</h1>"
					Case heading_tag = Headings.h2
						s_open_Tag = "<h2 "
						s_close_tag = "</h2>"
					Case heading_tag = Headings.h3
						s_open_Tag = "<h3 "
						close_tag = "</h3>"
					Case heading_tag = Headings.h4
						s_open_Tag = "<h4 "
						s_close_tag = "</h4>"
					Case heading_tag = Headings.h5
						s_open_Tag = "<h5 "
						s_close_tag = "</h5>"
					Case heading_tag = Headings.h6
						s_open_Tag = "<h6 "
						s_close_tag = "</h6>"
					Case Else
						s_open_Tag = "<h1 "
						s_close_tag = "</h1>"
				End Select

					s_heading_text = s_open_Tag

				'Optional Arguments
				If Len(heading_style) > 0 Then
					s_heading_text = s_heading_text & "Style='" & heading_style & "' "
				End if

				If Len(heading_class) > 0 Then
					s_heading_text = s_heading_text & "class='" & heading_class & "' "
				End if

				If Len(heading_id) > 0 Then
					s_heading_text = s_heading_text & "id='" & heading_id & "' "
				End if

				'Full Heading text in HTML
				s_heading_text = s_heading_text & ">" & heading_Text & s_close_tag

				'Heading HTML Text added to HTML_Array
				HTML_Array(row, column) = HTML_Array(row, column) & s_heading_text & vbNewLine

			End Sub


			Sub add_div(row as integer, column as integer, div_text as string)
				'This Sub is used to add custom elements to HTML_Array

				HTML_Array(row, column) = HTML_Array(row, column) & div_text

			End Sub


			'Script and Style links and code
			Sub add_styleLink(style_link as string)
				HTML_Style_Links = HTML_Style_Links & "<link rel='stylesheet' href='"& style_text & "'>" & vbNewLine
			End Sub


			Sub add_style(style_text as string)
				HTML_Style = HTML_Style & "<Style>" & vbNewLine & style_text & vbNewLine & "</style>" & vbNewLine
			End Sub


			Sub add_script_top_link(script_Link as string)
				HTML_Script_Top_Links = HTML_Script_Top_Links & "<script src='"& script_Link & "'>" & vbNewLine
			End Sub


			Sub add_scriptBottomLink(script_Link as string)
				HTML_Script_Bottom_Links = HTML_Script_Bottom_Links & "<script src='"& script_Link & "'>" & vbNewLine
			End Sub


			Sub  add_scriptBottom(script_Text as string)
				HTML_Script = HTML_Script & "<script>" & vbNewLine & script_Text & vbNewLine & "</script>" & vbNewLine
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

	Public Function SQL_to_array(SQL as String) as Variant
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

	Public Function SQL_to_JS_array(SQL as String) as Variant

	End Function

	Public Function SQL_to_json(SQL as String) as String

	End Function

	Public Function compile_to_String() as String

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
