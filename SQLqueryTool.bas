Attribute VB_Name = "SQLqueryTool"
Private row2$
Private rs As New ADODB.Recordset
Private loadList$, parameter$, sqlText$, operatorINNOTIN$
'------------------------------------------ START ------------------------------------
Sub Initialization()
LoginForm.ComboBox1.AddItem "NORTHWND"
LoginForm.ComboBox1.AddItem "AdventureWorks2017"
LoginForm.TextBox2.value = "adam"
LoginForm.TextBox3.value = "123"
LoginForm.Show '---------------------------------- LOGIN TO DB SCREEN ----------------
End Sub

'#####################################################################################

Sub Data_query(pwd$, dbAddress$, serverName$, userID$)
Dim connectionObject As ADODB.connection
Dim X As Long
Dim arr() As Variant
Dim myPAth$
Dim conObj As ConnectionClass
Dim len_Str As Integer
Dim hnd As Integer
' ------------------------ PREPARING NEW SHEET FOR DATA
myPAth = ""

Unload LoginForm  '---------------- DISABLE LOGIN FORM

'Open myPAth For Input As hnd
sqlText = ""
hnd = FreeFile
myPAth = File_Loader
If myPAth = "" Then
    MsgBox "No file selected"
GoTo end_sub
End If

Set conObj = New ConnectionClass
conObj.LETpassword = pwd
pwd = ";)"
conObj.LETuserID = userID
conObj.LETserverName = serverName
conObj.LETdbAddress = dbAddress
conObj.openConnection

 '------------------------------------- QUERY FILE READING ---------------------------

Dim row$
Open myPAth For Input As FreeFile
Do Until EOF(hnd)
    Line Input #hnd, row
    row = rowDoubleSpaceTABRemover(row)
    If InStr(1, row, "--\\--") Then '----------------- CONDITION FOR PARAMETER
        row = filtering(row, conObj)
    End If
    sqlText = sqlText & row & " " & vbNewLine '--------------------- ADDING NEW LINE FROM TEXT/SQL FILE
Loop
Close FreeFile

Set connectionObject = conObj.GetConnectionObject
On Error GoTo sql_error_info:
rs.Open sqlText, connectionObject, adOpenDynamic, adLockReadOnly '------------------- SENDING BUILT SQL QUERY TO SERVER VIA MS ACTIVE X LIBRARY 6.1

Set Data = ActiveWorkbook.Sheets.Add
Data.Select

For X = 1 To rs.Fields.Count
    Data.Cells(1, X) = rs.Fields(X - 1).Name
Next
    
If rs.RecordCount < Rows.Count Then
    Data.Range("A2").CopyFromRecordset rs
Else
    Do While Not rs.EOF
        row = row + 1
        For Findex = 0 To rs.Fields.Count - 1
            If row >= Rows.Count - 50 Then
                Exit For
            End If
            Data.Cells(row + 1, Findex + 1) = rs.Fields(Findex).value
        Next Findex
        rs.MoveNext
    Loop
End If

Cells.EntireColumn.AutoFit
Set connectionObject = Nothing
Set conObj = Nothing
Set conn = Nothing
Exit Sub

sql_error_info:
Set connectionObject = Nothing
Set conObj = Nothing
Set conn = Nothing
MsgBox "File selected may contain compile error or some objects in query may not exist"
end_sub:
End Sub
Private Function filtering(row As String, ByRef conObj)
Dim len_Str As Integer
Dim row1 As String
    
If InStr(1, row, "= '") > 0 Then
    len_Str = InStr(1, row, "= '")
    row1 = Mid(row, 1, len_Str)
    ParamInsert.Label1.Caption = "Insert: " & Mid(row1, InStr(1, row1, " "))
    ParamInsert.Label2.Caption = "Insert exact phrase"
    parameter = ""
    ParamInsert.Show
    If parameter = "" Then
        filtering = "--\\--" & row
    Else
        filtering = (row1 & " '" & parameter & "'")
    End If
ElseIf InStr(1, UCase(row), "LIKE '") > 0 Or InStr(1, UCase(row), "LIKE'") > 0 Then
    len_Str = InStr(1, UCase(row), "LIKE '")
    If len_Str = 0 Then len_Str = InStr(1, row, "like '")
    row1 = Mid(row, 1, len_Str - 1)
    ParamInsert.Label1.Caption = "Insert: " & Mid(row1, InStr(1, row1, " "))
    ParamInsert.Label2.Caption = "Insert whole or part of phrase"
    parameter = ""
    ParamInsert.Show
    If parameter = "" Then
        filtering = "--\\--" & row
    Else
        filtering = (row1 & "LIKE '%" & parameter & "%'")
    End If
ElseIf InStr(1, UCase(row), "NOT IN ('") > 0 Or InStr(1, UCase(row), "NOT IN('") > 0 Then
    ' Dim list_sting As String
    operatorINNOTIN = "NOT IN"
    len_Str = InStr(1, Replace(row, "NOT IN ('", "NOT IN('"), "NOT IN('")
    If len_Str = 0 Then len_Str = InStr(1, Replace(row, "not in ('", "NOT IN('"), "NOT IN('")
    row1 = Trim(Left(row, len_Str - 1))
    row2 = Mid(row, len_Str)
    dbField = Mid(row1, InStr(1, row1, ".") + 1)
    listing = get_list("SELECT DISTINCT " & dbField & " FROM " & GetTableFromAlias(row1, dbField) & " ORDER BY 1 ASC", conObj)
    If listing <> "" Then
        filtering = row1 & " " & operatorINNOTIN & " (" & listing & ")"
    Else
        filtering = " "
    End If
ElseIf InStr(1, UCase(row), "IN ('") > 0 Or InStr(1, UCase(row), "IN('") > 0 Then
    ' Dim list_sting As String
    operatorINNOTIN = "IN"
    len_Str = InStr(1, Replace(row, "IN ('", "IN('"), "IN('")
    If len_Str = 0 Then len_Str = InStr(1, Replace(row, "in ('", "IN('"), "IN('")
    row1 = Trim(Left(row, len_Str - 1))
    row2 = Mid(row, len_Str)
    dbField = Mid(row1, InStr(1, row1, ".") + 1)
    listing = get_list("SELECT DISTINCT " & dbField & " FROM " & GetTableFromAlias(row1, dbField) & " ORDER BY 1 ASC", conObj)
    If listing <> "" Then
        filtering = row1 & " IN(" & listing & ")"
    Else
        filtering = " "
    End If
End If
End Function
Private Function GetTableFromAlias(ByVal row1$, ByVal dbField$) As String
tableAlias = Mid(Left(row1, InStr(1, row1, ".") - 1), InStrRev(Left(row1, InStr(1, row1, ".") - 1), " ") + 1)
aliasLocation = InStr(1, sqlText, " " & tableAlias & " ")
If aliasLocation > 0 Then
    GetTableFromAlias = Mid(Left(sqlText, aliasLocation - 1), InStrRev(Left(sqlText, aliasLocation - 1), " ") + 1)
Else
    GetTableFromAlias = "Alias not found"
    MsgBox "Reconstruct the query"
    End
End If
End Function
Private Function get_list(list_string As String, ByRef conObj)
Dim arrList() As Variant
Dim connectionObject As ADODB.connection
    
Set connectionObject = conObj.GetConnectionObject
rs.Open list_string, connectionObject, adOpenDynamic, adLockReadOnly
If rs.RecordCount < Rows.Count Then
    'ActiveSheet.Range("A2").CopyFromRecordset RS
    rs.MoveFirst
    arrList = rs.GetRows
Else
    Do While Not rs.EOF
        row = row + 1
        For Findex = 0 To rs.Fields.Count - 1
            If row >= Rows.Count - 50 Then
                Exit For
            End If
            Data.Cells(row + 1, Findex + 1) = rs.Fields(Findex).value
        Next Findex
        rs.MoveNext
    Loop
End If
rs.Close

Set rs = Nothing
For I = LBound(arrList, 2) To UBound(arrList, 2)
        ParamSelectList.ListBox1.AddItem arrList(0, I)
Next I
ParamSelectList.Label1.Caption = "rs.Fields.Item(0).Name"
loadList = ""
ParamSelectList.ComboBox1.AddItem "IN"
ParamSelectList.ComboBox1.AddItem "NOT IN"
If operatorINNOTIN = "IN" Then
    ParamSelectList.ComboBox1.ListIndex = 0
Else
    ParamSelectList.ComboBox1.ListIndex = 1
End If
ParamSelectList.Show
get_list = Mid(loadList, 3)
End Function
Private Function File_Loader()
Dim slashPos As Long
Dim fileNameLen As Byte

Set FilePicker = Application.FileDialog(msoFileDialogOpen)
With FilePicker
    .Title = "Select A Target File"
    .AllowMultiSelect = False
    .Filters.Clear
    .Filters.Add "All Files", "*.txt;*.sql"
    On Error Resume Next
    If .Show <> -1 Then GoTo NextCode
    myPAth = .SelectedItems(1) & "\"
End With

slashPos = InStrRev(myPAth, "\")
If InStrRev(myPAth, "\") = Len(myPAth) Then
    myPAth = Mid(myPAth, 1, slashPos - 1)
End If
slashPos = InStrRev(myPAth, "\")
fileNameLen = Len(myPAth) - slashPos
File_Loader = myPAth
NextCode:
End Function
Private Function rowDoubleSpaceTABRemover(row$) As String
Do While InStr(1, row, vbTab)
    row = Replace(row, vbTab, " ")
Loop
Do While InStr(1, row, "  ")
    row = Replace(row, "  ", " ")
Loop
rowDoubleSpaceTABRemover = Trim(row)
End Function
Sub loadStringList()
For I = 0 To ParamSelectList.ListBox2.ListCount - 1
    loadList = loadList & ", " & "'" & ParamSelectList.ListBox2.List(I) & "'"
Next I
End Sub
Sub loadListINNOTIN()
    operatorINNOTIN = ParamSelectList.ComboBox1.value
End Sub
Sub LETparameter(value As String)
    parameter = value
End Sub
