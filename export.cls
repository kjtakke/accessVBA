Option Compare Database
'This class is designe to export SQL data to HTML (preview), CSV and Email
'Dim ex as New export
'ex.csv = "SELECT ..."
'ex.prieview = "SELECT ..."
'ex.email = "SELECT ..."


Public Property Let preview(SQL As String)
    Dim rst As DAO.Recordset
    Dim varMyArray As Variant
    Dim varField As Variant
    
    On Error GoTo En:
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
    
    rst.MoveLast
    ReDim varMyArray(rst.RecordCount - 1, rst.Fields.count - 1)
    rst.MoveFirst
    
    Do While Not rst.EOF
        For Each varField In rst.Fields
        varMyArray(rst.AbsolutePosition, varField.OrdinalPosition) = varField
        Next varField
        rst.MoveNext
    Loop
    
    'Get Array Dimentions
    Dim DimCount As Integer
    DimCount = arrayDimentionCounter(varMyArray)
            
    Dim i As Single, j As Single, k As Single
    Dim html As String
    Dim field
    
    html = "<!DOCTYPE html><html lang='en'><head><title>Access Table View</title><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>"
    html = html & "<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css'>"
    html = html & "<script src='https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js'></script>"
    html = html & "<script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js'></script>"
    html = html & "<script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js'></script>"
    html = html & "</head>"
    html = html & "<body>"
    html = html & "<table class='table table-hover'>"

    html = html & "<thead>"
    
        'Cycle through Filed Headers
        For i = 1 To rst.Fields.count
            html = html & "<th>"
            html = html & rst.Fields(i - 1).Name
            html = html & "</th>"
        Next i
    
    html = html & "</thead>"
    html = html & "<tbody>"
        
        'Cycle through Filed Body
        For i = 0 To UBound(varMyArray)
            html = html & "<tr>"
            For j = 0 To DimCount
                html = html & "<td>"
                html = html & varMyArray(i, j)
                html = html & "</td>"
            Next j
            html = html & "</tr>"
        Next i
    html = html & "</tbody>"
    
    html = html & "</table>"
    html = html & "</body>"
    html = html & "</html>"
    
    FilePath = "C:\Users\" & Environ("Username") & "\Desktop\table.html"

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(FilePath, True)
    a.WriteLine html
    a.Close
    saved = True
    
    Application.FollowHyperlink (FilePath)
    
    GoTo Fn:
En:
MsgBox ("Invalid Query or query refering to an attachment field")

Fn:


End Property

Public Property Let csv(SQL As String)
    Dim rst As DAO.Recordset
    Dim varMyArray As Variant
    Dim varField As Variant
    
    On Error GoTo En:
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
    rst.MoveLast
    ReDim varMyArray(rst.RecordCount - 1, rst.Fields.count - 1)
    rst.MoveFirst
    
    Do While Not rst.EOF
        For Each varField In rst.Fields
        varMyArray(rst.AbsolutePosition, varField.OrdinalPosition) = varField
        Next varField
        rst.MoveNext
    Loop
    
    'Get Array Dimentions
    Dim DimCount As Integer
    DimCount = arrayDimentionCounter(varMyArray)

    Dim i As Single, j As Single, k As Single
    Dim csv As String
    Dim field
    Dim fld As String
    
    csv = ""
        'Cycle through Filed Headers
        For i = 1 To rst.Fields.count

            csv = csv & rst.Fields(i - 1).Name & ","
        Next i
        csv = csv & vbNewLine
        
        'Cycle through Filed Body
        For i = 0 To UBound(varMyArray)
            For j = 0 To DimCount
                
                fld = Replace(varMyArray(i, j), ",", " ")
                If j = DimCount Then
                csv = csv & fld
                Else
                csv = csv & fld & ","
                End If
            Next j
            csv = csv & vbNewLine
        Next i

    FilePath = "C:\Users\" & Environ("Username") & "\Desktop\table.csv"

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(FilePath, True)
    a.WriteLine csv
    a.Close
    saved = True
    
    
    GoTo Fn:
En:
MsgBox ("Invalid Query or query refering to an attachment field")

Fn:

End Property

Public Property Let mail(SQL As String)
    Dim OlApp As Outlook.Application
    Dim olEmail As Outlook.MailItem
    Dim i, ii As Long
    Dim data As Variant
    Dim Str As String

    Dim rst As DAO.Recordset
    Dim varMyArray As Variant
    Dim varField As Variant

    
    Set OlApp = New Outlook.Application
    Set olEmail = OlApp.CreateItem(olMailItem)
    On Error GoTo En:
    
    Set rst = CurrentDb.OpenRecordset(SQL)
    
    rst.MoveLast
    ReDim varMyArray(rst.RecordCount - 1, rst.Fields.count - 1)
    rst.MoveFirst
    
    Do While Not rst.EOF
        For Each varField In rst.Fields
        varMyArray(rst.AbsolutePosition, varField.OrdinalPosition) = varField
        Next varField
        rst.MoveNext
    Loop
    
    'Get Array Dimentions
    Dim DimCount As Integer
    DimCount = arrayDimentionCounter(varMyArray)

    Dim field
    
    Str = "<br><br><br><!DOCTYPE html><html lang='en'><head><title>Access Table View</title><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>"
    Str = Str & "</head>"
    Str = Str & "<body>"
    Str = Str & "<table  cellspacing='0' border='1'>"
    Str = Str & "<thead>"
        Str = Str & "<tr>"
        
        'Cycle through Filed Headers
        For i = 1 To rst.Fields.count
            Str = Str & "<th>"
            Str = Str & rst.Fields(i - 1).Name
            Str = Str & "</th>"
        Next i
        Str = Str & "</tr>"
        
    Str = Str & "</thead>"
    Str = Str & "<tbody>"
    
        'Cycle through Filed Body
        For i = 0 To UBound(varMyArray)
            Str = Str & "<tr>"
            For j = 0 To DimCount
                Str = Str & "<td>"
                Str = Str & varMyArray(i, j)
                Str = Str & "</td>"
            Next j
            Str = Str & "</tr>"
        Next i
    Str = Str & "</tbody>"
    
    Str = Str & "</table>"

    Str = Str & "</body>"
    Str = Str & "</html>"
    
With olEmail
    .BodyFormat = olFormatHTML
    .Display
    .HTMLBody = Str
End With
    
    GoTo Fn:
En:
MsgBox ("Invalid Query or query refering to an attachment field")
Fn:
End Property



'Array Dementions
Function arrayDimentionCounter(index As Variant) As Integer
'This Function Counts the Columns/Dimentions in an Array
'index is the input array

    On Error GoTo LC:
    For L = 1 To 100
        TempVar = index(1, L)
    Next L
LC:
    L = L - 1
    On Error GoTo 0
    arrayDimentionCounter = L
End Function
