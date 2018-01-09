VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private connection As ADODB.connection
Private cmd As ADODB.Command
Private password$, dbAddress$, serverName$, userID$
'------------------ CREATION OF CONNECTION OBJECT --------------------
Private Sub class_initialize()
Set connection = New ADODB.connection
Set cmd = New ADODB.Command
End Sub
Private Sub Class_Terminate() 'Clean up after yourself
connection.Close
Set connection = Nothing
Set cmd = Nothing
End Sub
Property Let LETdbAddress(value As String)
dbAddress = value
End Property
Property Let LETserverName(value As String)
serverName = value
End Property
Property Let LETuserID(value As String)
userID = value
End Property
Property Let LETpassword(value As String)
password = value
End Property
Public Sub openConnection()
connection.Open "Driver={SQL Server};Server=" & serverName & ";Database=" & dbAddress & _
    ";Uid=" & "adam" & ";Pwd=" & password & ";"
End Sub
Public Function GetConnectionObject() As ADODB.connection
Set GetConnectionObject = connection
End Function
