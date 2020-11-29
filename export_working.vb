Option Compare Database

'EXPORT CLASS
'This class is designed to create web dashboards in Microsoft Access
'
'REQUIRED PACKAGES
'   Visual Basic For Applications
'   Microsoft Access 16.0 Object Library
'   OLE Automation
'   System
'   Microsoft ActiveX Data Objects 6.1 Library
'   Microsoft ActiveX Data Objects Recordset 6.0 Library
'   Microsoft ADO Ext. 6.0 for DDL and Security
'   Microsoft Data Access Components Installed Version
'   Microsoft ADO 3.6 Object Library
'   Microsoft Outlook 16.0 Object Library
'   Microsoft Forms 2.0 Object Library (Browse for FO20.DLL)
'
'RESOURCES
'   Charts.JS
'   Bootstrap 4
'   HTML 5
'   fontsAwsome
' jQuery
'
'REFRENCES
'
'
'
'NAMING COVENTIONS
'   Private: pv_finctionName
'   Public: finctionName
'
'   Private Variables:
'       variant: a_VariableName
'       String: s_VariableName
'       Integer: i_VariableName
'       boolean: b_VariableName
'       double: d_VariableName
'       collection: c_VariableName
'       dictionary: dict_VariableName
'       object: o_VariableName:
'           o_file
'           o_fileInstance
'           o_mail
'       Counters:
'           i, j, k, h as Single
'
'COLORS
'   black: #000000  rgb(0,0,0)
'
'COLOR SETS
'   default: [red, green, ...]
'   ...: [..., ..., ...]
'
'ICONS
'
'
'EXAMPLES
'
'Sub btn_click()
'
'   Const SQL_Sales = "SELECT ..."
'   Const SQL_Metrics = "SELECT ..."
'   Const MH = "My Page Heading"
'
'   Dim d as export
'   Set d = New export
'
'   'Set HTML Document's rows and columns
'   d.HTML_dimentions(5,5)
'
'   'Add HTML Elements
'   d.add_heading(2,1,1,MH)
'   d.add_table(2,1,SQL_Sales)
'   d.add_chart(2,2,"pie",SQL_Metrics)
'
'   'Compile and export HTML Document
'   d.compile_and_export()
'
'End Sub


'Public Constants:

Const bootstrapCSS As String = "<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css'>"
Const bootstrapJS As String = "<script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js'></script>"
Const chartsJS As String = "<script src='https://cdn.jsdelivr.net/npm/chart.js@2.8.0'></script>"
Const jQuery As String = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'>"
Const fontsAwsomeCSS As String = "<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css'>"
Const googleapis As String = "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'></script>"
Const cloudflare As String = "<script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js'></script>"
Const dashboardCSS As String = "<link rel='stylesheet' href='https://allform-tech-200815-customer.github.io/page-templates/styles.css'>"


'Public Variables:

Public HTML_Array As Variant
Public HTML_Column_Count As Integer
Public HTML_Row_Count As Integer
Public HTML_Script As String
Public HTML_Style As String
Public HTML_Script_Top_Links As String
Public HTML_Script_Bottom_Links As String
Public HTML_Style_Links As String
Public HTML_File_Name As String
Public HTML_File_Path As String
Public HTML_Elements_Count As Integer
Public HTML_Title As String
Public HTML_Heading As String
Public HTML_Heaader As Boolean

Public Current_Colors As Variant
Public Current_Icon As Variant
Public Current_SQL As String
Public Current_Array As Variant
Public Current_Row As Integer
Public Current_Column As Integer
Public Current_Dim_Count As Integer

Public h As Single
Public i As Single
Public j As Single
Public k As Single


'Enumerations:

Public Enum tableClasses
    'table
    Table
    'table table-striped
    table_striped
    'table table-bordered
    table_bordered
    'table table-hover
    table_hover
    'table table-dark
    table_dark
    'table table-dark table-striped
    table_dark_striped
    'table table-dark table-hover
    table_dark_hover
    'table table-borderless
    table_borderless
End Enum


Public Enum chartType
    line_chart
    area_chart
    bar_chart
    hBar_chart
    pie_chart
    donut_chart
End Enum


Public Enum headings
    h1
    h2
    h3
    h4
    h5
    h6
End Enum

Public Enum metrics
    'tile.overdue background-color: #f21313;
    overdue

    'tile.due background-color: #f08f11;
    due

    'tile.completed background-color: #000;
    completed

    'tile.open background-color: #598bff;
    outstanding

End Enum






'Public Subs:

    Sub HTML_Setup(rows As Integer, columns As Integer, Optional fileName As String = "", Optional filepath As String = "", Optional heading As String = "", Optional Title As String = "")
        'This Sub sets the dimentsions for the HTML Document

        'Optional Arguments
        If Len(fileName) = 0 Then fileName = "Report"
        If Len(fileName) = 0 Then filepath = "C:\Users\" & Environ("Username") & "\Desktop\"

        'Load Public Variables
        If Len(heading) = 0 Then
            HTML_Heaader = False
        Else
            HTML_Heaader = True
        End If
        If Len(Title) = 0 Then HTML_Title = "Report"
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
        ReDim HTML_Array(1 To rows, 1 To columns)

        'Set each array element to an empty string
        For i = 1 To rows
            For j = 1 To columns
                HTML_Array(i, j) = ""
            Next j
        Next i

    End Sub


    'Add Elements

            Sub add_table(row As Integer, column As Integer, sql As String, Optional table_style As String = "", Optional table_class As tableClasses = 0, Optional table_id As String = "")
                'This sub creats a HTML Table from SQL->Aray

                Dim s_table_text As String
                Dim index As Variant
                Dim i_dimCount As Integer
                
                i_dimCount = pv_dimentionCount(sql)
                
                'SQL to array
                index = SQL_to_array(sql)

                'Table Tag Optional Arguments
                s_table_text = "<table "
                If Len(table_style) > 0 Then s_table_text = s_table_text & "style='" & table_style & "' "

                If Len(table_id) > 0 Then s_table_text = s_table_text & "style='" & table_id & "' "
                s_table_text = s_table_text & vbNewLine

                If table_class = 0 Then
                    s_table_text = s_table_text & "class='table' "
                Else
                    Select Case True
                        'table
                        Case table_class = tableClasses.Table
                            s_table_text = s_table_text & "class='table' "

                        'table table-striped
                        Case table_class = tableClasses.table_striped
                            s_table_text = s_table_text & "class='table table-striped' "

                        'table table-bordered
                        Case table_class = tableClasses.table_bordered
                            s_table_text = s_table_text & "class='table table-bordered' "

                        'table table-hover
                        Case table_class = tableClasses.table_hover
                            s_table_text = s_table_text & "class='table table-hover' "

                        'table table-dark
                        Case table_class = tableClasses.table_dark
                            s_table_text = s_table_text & "class='table table-dark' "

                        'table table-dark table-striped
                        Case table_class = tableClasses.table_dark_striped
                            s_table_text = s_table_text & "class='table table-dark table-striped' "

                        'table table-dark table-hover
                        Case table_class = tableClasses.table_dark_hover
                            s_table_text = s_table_text & "class='table table-dark table-hover' "

                        'table table-borderless
                        Case table_class = tableClasses.table_borderless
                            s_table_text = s_table_text & "class='table table-borderless' "

                        Case Else
                            s_table_text = s_table_text & "class='table' "
                    End Select
                End If

                    s_table_text = s_table_text & ">" & vbNewLine

                    'Table Headers
                    s_table_text = s_table_text & "<thead>" & vbNewLine
                    s_table_text = s_table_text & "<tr>" & vbNewLine
                    For i = 0 To Current_Dim_Count
                        s_table_text = s_table_text & "<th>" & vbNewLine
                            s_table_text = s_table_text & index(0, i) & vbNewLine
                        s_table_text = s_table_text & "</th>" & vbNewLine
                    Next i
                    s_table_text = s_table_text & "</tr>" & vbNewLine
                    s_table_text = s_table_text & "</thead>" & vbNewLine

                    'Table Body
                    s_table_text = s_table_text & "<tbody>" & vbNewLine
                        For i = 1 To UBound(index)
                            s_table_text = s_table_text & "<tr>" & vbNewLine
                                For j = 0 To Current_Dim_Count
                                    s_table_text = s_table_text & "<td>" & vbNewLine
                                        s_table_text = s_table_text & index(i, j) & vbNewLine
                    'Debug.Print (index(i, j))
                                    s_table_text = s_table_text & "</td>" & vbNewLine
                                Next j
                            s_table_text = s_table_text & "</tr>" & vbNewLine
                        Next i
                    s_table_text = s_table_text & "</tbody>" & vbNewLine

                'Table close
                s_table_text = s_table_text & "</table>"

                'Load s_table_text (HTML) to HTML_Array
                HTML_Array(row, column) = HTML_Array(row, column) & s_table_text

            End Sub


            Sub add_metric(row As Integer, column As Integer, sql As String, Optional metric_prefix As String = "", Optional metric_sufix As String, Optional metric_style As String = "", Optional metric_class As metrics = 3, Optional metric_id As String = "")
                'This Sub adds a button Metric examples at: https://allform-tech-200815-customer.github.io/page-templates/index.html

                Dim s_metric As String
                Dim s_metric_heading As String
                Dim s_metric_number As String
                Dim s_metric_class As String
                Dim index As Variant

                Select Case True

                    'tile.overdue background-color: #f21313;
                    Case metric_class = metrics.overdue
                        s_metric_class = "tile overdue"

                    'tile.due background-color: #f08f11;
                    Case metric_class = metrics.due
                        s_metric_class = "tile due"

                    'tile.completed background-color: #000;
                    Case metric_class = metrics.completed
                        s_metric_class = "tile completed"

                    'tile.open background-color: #598bff;
                    Case metric_class = metrics.outstanding
                        s_metric_class = "tile open"

                    Case Else
                        s_metric_class = "tile open"

                End Select

                'SQL to array
                index = SQL_to_array(sql)

                'Assign metric elements
                s_metric_heading = index(0, 0)
                s_metric_number = index(1, 0)

                'Optional Arguments added to metric
                If Len(metric_prefix) > 0 Then s_metric_number = metric_prefix & s_metric_number
                If Len(metric_sufix) > 0 Then s_metric_number = s_metric_number & metric_sufix

                s_metric = "<div align='center'>" & vbNewLine & _
                                     "<button type='button' name='button' class='" & s_metric_class & " style='" & metric_style & "' " & "id='" & metric_id & "'>" & _
                                     "<div class='tile-measure'>" & s_metric_number & "</div><br>" & vbNewLine & _
                                     "<span class='tile-comment'>" & s_metric_heading & "</span>" & vbNewLine & _
                                     "</button>" & vbNewLine & _
                                     "</div>"

                HTML_Array(row, column) = HTML_Array(row, column) & s_metric

            End Sub


            Sub add_chart(row As Integer, column As Integer, sql As String, chart_type As chartType, chart_id As String, Optional chart_prefix As String = "", Optional chart_sufix As String = "", Optional chart_style = "")

                    Dim index As Variant
                    Dim s_chart_data As String
                    Dim s_chart_lables As String
                    index = SQL_to_array(sql)

                    Select Case True
                        Case chart_type = chartType.line_chart


                        Case chart_type = chartType.pie_chart

                            'Pie Chart Data
                            s_chart_data = "["
                            For i = 1 To UBound(index)
                                If i = UBound(index) Then
                                    s_chart_data = s_chart_data & index(i, 1)
                                Else
                                    s_chart_data = s_chart_data & index(i, 1) & ", "
                                End If
                            Next i
                            s_chart_data = s_chart_data & "]"

                            'Pie Chart Lables
                            s_chart_lables = "["
                            For i = 1 To UBound(index)
                                If i = UBound(index) Then
                                    s_chart_lables = s_chart_lables & index(i, 0)
                                Else
                                    s_chart_lables = s_chart_lables & index(i, 0) & ", "
                                End If
                            Next i
                            s_chart_lables = s_chart_lables & "]"

                            HTML_Script = HTML_Script & pv_pieChartScript(chart_id, s_chart_data, s_chart_lables, chart_prefix, chart_sufix)

                            HTML_Array(row, column) = HTML_Array(row, column) & "<div>"
                            HTML_Array(row, column) = HTML_Array(row, column) & "<table class='charts-table'>"
                            HTML_Array(row, column) = HTML_Array(row, column) & "<td class='charts-td-50' class='charts-canvas' style='" & chart_style & "'>"
                            HTML_Array(row, column) = HTML_Array(row, column) & "<canvas id='" & chart_id & "'></canvas>"
                            HTML_Array(row, column) = HTML_Array(row, column) & "</td>"
                            HTML_Array(row, column) = HTML_Array(row, column) & "</table>"
                            HTML_Array(row, column) = HTML_Array(row, column) & "</div>" & vbNewLine

                        Case chart_type = chartType.area_chart


                        Case chart_type = chartType.hBar_chart


                        Case chart_type = chartType.bar_chart


                        Case chart_type = chartType.donut_chart


                        Case Else


                    End Select
            End Sub


            Sub add_heading(row As Integer, column As Integer, heading_tag As headings, heading_Text As String, Optional heading_style As String = "", Optional heading_class As String = "", Optional heading_id As String = "")
                'This Sub creates a <H1-6> Tag
                Dim s_heading_text As String
                Dim s_open_Tag As String
                Dim s_close_tag As String

                'Tag
                Select Case True
                    Case heading_tag = headings.h1
                        s_open_Tag = "<h1 "
                        s_close_tag = "</h1>"
                    Case heading_tag = headings.h2
                        s_open_Tag = "<h2 "
                        s_close_tag = "</h2>"
                    Case heading_tag = headings.h3
                        s_open_Tag = "<h3 "
                        close_tag = "</h3>"
                    Case heading_tag = headings.h4
                        s_open_Tag = "<h4 "
                        s_close_tag = "</h4>"
                    Case heading_tag = headings.h5
                        s_open_Tag = "<h5 "
                        s_close_tag = "</h5>"
                    Case heading_tag = headings.h6
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
                End If

                If Len(heading_class) > 0 Then
                    s_heading_text = s_heading_text & "class='" & heading_class & "' "
                End If

                If Len(heading_id) > 0 Then
                    s_heading_text = s_heading_text & "id='" & heading_id & "' "
                End If

                'Full Heading text in HTML
                s_heading_text = s_heading_text & ">" & heading_Text & s_close_tag

                'Heading HTML Text added to HTML_Array
                HTML_Array(row, column) = HTML_Array(row, column) & s_heading_text & vbNewLine

            End Sub


            Sub add_div(row As Integer, column As Integer, div_text As String)
                'This Sub is used to add custom elements to HTML_Array

                HTML_Array(row, column) = HTML_Array(row, column) & div_text

            End Sub


            'Script and Style links and code
            Sub add_styleLink(style_link As String)
                HTML_Style_Links = HTML_Style_Links & "<link rel='stylesheet' href='" & style_text & "'>" & vbNewLine
            End Sub


            Sub add_style(style_text As String)
                HTML_Style = HTML_Style & "<Style>" & vbNewLine & style_text & vbNewLine & "</style>" & vbNewLine
            End Sub


            Sub add_script_top_link(script_Link As String)
                HTML_Script_Top_Links = HTML_Script_Top_Links & "<script src='" & script_Link & "'>" & vbNewLine
            End Sub


            Sub add_scriptBottomLink(script_Link As String)
                HTML_Script_Bottom_Links = HTML_Script_Bottom_Links & "<script src='" & script_Link & "'>" & vbNewLine
            End Sub


            Sub add_scriptBottom(script_Text As String)
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

    Public Function SQL_to_array(sql As String) As Variant
        Dim o_rst As DAO.Recordset
        Dim a_SQL As Variant
        Dim a_varField As Variant
        Dim i_dimCount As Integer
        Set o_rst = CurrentDb.OpenRecordset(sql)

        'Set Array Dimentions
        o_rst.MoveLast
        ReDim a_SQL(0 To o_rst.RecordCount, 0 To o_rst.Fields.count - 1)
        o_rst.MoveFirst

        'Add Filed Headers To Array
        For i = 0 To o_rst.Fields.count - 1
                a_SQL(0, i) = o_rst.Fields(i).Name
        Next i

        'SQL Body to VBA Array
        Do While Not o_rst.EOF
            For Each a_varField In o_rst.Fields
            a_SQL(o_rst.AbsolutePosition + 1, a_varField.OrdinalPosition) = a_varField
            Next a_varField
            o_rst.MoveNext
        Loop

        'Get/Confirm Array Dimentions
        Curent_Dim_Count = pv_dimentionCount(a_SQL)

        'Set Public Variables
        Current_SQL = sql
        Current_Array = a_SQL

        'Return Array
        SQL_to_array = a_SQL

    End Function

    Public Function SQL_to_JS_array(sql As String) As Variant

    End Function

    Public Function SQL_to_json(sql As String) As String

    End Function

    Public Function compile_to_String() As String

    End Function


'Private Functions:

    Private Function pv_dimentionCount(index As Variant) As Integer
    'This Function Counts the Columns/Dimentions in an Array
    'index is the input array

        On Error GoTo LC:
        For i = 1 To 100
                TempVar = index(1, i)
        Next i
LC:
        i = i - 1
        On Error GoTo 0
        pv_dimentionCount = i
        Current_Dim_Count = i
    End Function


    Private Function pv_chartTemplate(index As Variant) As String

    End Function


    Private Function pv_HTMLTemplate(index As Variant) As String

    End Function


    Private Function pv_styleTagTemplate(index As Variant) As String

    End Function


    Private Function pv_scriptTagTemplate(index As Variant) As String

    End Function


    Private Function pv_styleTemplate(index As Variant) As String

    End Function


    Private Function pv_scriptTemplate(index As Variant) As String

    End Function

    Private Function pv_icon_Template(index As Variant) As String

    End Function


    Private Function pv_metric_Template(index As Variant) As String

    End Function

    Private Function pv_pieChartScript(pie_id, pie_data As String, pie_labels As String, Optional pie_prefix As String = "", Optional pie_sufix As String = "", Optional pie_colors As String = "['#9c7272', '#9c8d72', '#729c7d', '#729c8e', '#727a9c', '#80729c', '#94729c', '#9c7280']") As String
        'This Function REturns the <SCRIPT> for a pie chart

        pv_pieChartScript = "var ctx = document.getElementById('" & pie_id & "').getContext('2d');" & vbNewLine
            pv_pieChartScript = pv_pieChartScript & "var myPie = new Chart(ctx, {" & vbNewLine
                pv_pieChartScript = pv_pieChartScript & "type: 'pie'," & vbNewLine
                pv_pieChartScript = pv_pieChartScript & "data: {" & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "labels: " & pie_labels & "," & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "datasets: [{" & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "backgroundColor: " & pie_colors & "," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "borderColor: '#000'," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "borderWidth: '0px'," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "data: " & pie_data & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "}]," & vbNewLine
                pv_pieChartScript = pv_pieChartScript & "}," & vbNewLine
                pv_pieChartScript = pv_pieChartScript & "options: {" & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "title: {" & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "display: true," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "text: 'By State'," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "fontStyle: 'bold'," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "fontSize: 20," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "fontColor: 'black'," & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "}," & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "legend: {" & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "display: true," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "labels: {" & vbNewLine
                            pv_pieChartScript = pv_pieChartScript & "fontColor: 'black'," & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "}," & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "tooltips: {" & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "callbacks: {" & vbNewLine
                            ' this callback is used to create the tooltip label
                            pv_pieChartScript = pv_pieChartScript & "label: function(tooltipItem, data) {" & vbNewLine
                                ' get the data label and data value to display
                            ' convert the data value to local string so it uses a comma seperated number
                                pv_pieChartScript = pv_pieChartScript & "var dataLabel = data.labels[tooltipItem.index];" & vbNewLine
                                pv_pieChartScript = pv_pieChartScript & "var value = '" & pie_prefix & "' + data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index].toLocaleString() + '" & pie_sufix & "';" & vbNewLine

                                ' make this isn't a multi-line label (e.g. [["label 1 - line 1, "line 2, ], [etc...]])
                                pv_pieChartScript = pv_pieChartScript & "if (Chart.helpers.isArray(dataLabel)) {" & vbNewLine
                                    ' show value on first line of multiline label
                                    ' need to clone because we are changing the value
                                    pv_pieChartScript = pv_pieChartScript & "dataLabel = dataLabel.slice();" & vbNewLine
                                    pv_pieChartScript = pv_pieChartScript & "dataLabel[0] += value;" & vbNewLine
                                pv_pieChartScript = pv_pieChartScript & "} else {" & vbNewLine
                                    pv_pieChartScript = pv_pieChartScript & "dataLabel += value;" & vbNewLine
                                pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine

                                ' return the text to display on the tooltip
                                pv_pieChartScript = pv_pieChartScript & "return dataLabel;" & vbNewLine
                            pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
                        pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
                    pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
                pv_pieChartScript = pv_pieChartScript & "}" & vbNewLine
            pv_pieChartScript = pv_pieChartScript & "});" & vbNewLine

    End Function


    Private Function pv_barChartScript(bar_id, bar_data As String, bar_labels As String, Optional bar_prefix As String = "", Optional bar_sufix As String = "", Optional bar_colors As String = "['#9c7280']") As String

    End Function

