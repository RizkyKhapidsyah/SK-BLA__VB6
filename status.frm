VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form status 
   Caption         =   "Spy Status"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "status.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox spycommand 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
   End
   Begin RichTextLib.RichTextBox spytxt 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      TextRTF         =   $"status.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If main.Winsock1.State = 7 Then Cancel = -1
End Sub

Private Sub Form_Resize()
On Error Resume Next
status.spytxt.Width = Me.Width - 125
status.spycommand.Width = Me.Width - 125
status.spytxt.Height = Me.Height - status.spycommand.Height - 500
End Sub

Private Sub spycommand_KeyPress(KeyAscii As Integer)
If main.Winsock1.State <> 7 Then
MsgBox "Non c'è nessuna connessione in atto.", vbExclamation
Exit Sub
End If
If KeyAscii = 13 Then
main.script.Run "SpyCommandline", spycommand.Text
End If

End Sub

