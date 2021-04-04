VERSION 5.00
Begin VB.Form UserForm 
   Caption         =   "Userform"
   ClientHeight    =   1890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   Icon            =   "UserForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1890
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "UserForm.frx":1CFA
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Text1.Width = Me.Width - 125
Text1.Height = Me.Height - 200
End Sub
