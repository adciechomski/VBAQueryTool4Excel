VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ParamSelectList 
   Caption         =   "ParamSelectList"
   ClientHeight    =   3648
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6012
   OleObjectBlob   =   "ParamSelectList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ParamSelectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Remove_param_Click()
I = 0
Do Until I > ListBox2.ListCount - 1
    If ListBox2.Selected(I) Then
        ListBox1.AddItem ListBox2.List(I)
        ListBox2.RemoveItem (I)
        I = I - 1
    End If
I = I + 1
Loop
End Sub
Private Sub CommandButton1_Click()
Call load_string_list
Unload Me
End Sub

'Private Sub Add_Click()
 '   Call Deal_list_Input
'End Sub

Private Sub Add_param_Click()
Do Until I > ListBox1.ListCount - 1
    If ListBox1.Selected(I) = True Then
        ListBox2.AddItem ListBox1.List(I)
        ListBox1.RemoveItem (I)
        I = I - 1
    End If
I = I + 1
Loop

End Sub


Private Sub userform_terminate() '----------- X terminate and after display terminate
    'Application.Visible = True
    'Call Terminate_form
End Sub

