VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Log In:"
   ClientHeight    =   4212
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4608
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Call Data_query(Me.TextBox3, Me.ComboBox1, Me.TextBox4, Me.TextBox2)
End Sub
