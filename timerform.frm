VERSION 5.00
Begin VB.Form timerform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timer"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Timer 3"
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Caption         =   "Intervallo (msec)"
         Height          =   615
         Left            =   1200
         TabIndex        =   15
         Top             =   120
         Width           =   1455
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Disabilita"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Abilita"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Timer 2"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Intervallo (msec)"
         Height          =   615
         Left            =   1200
         TabIndex        =   10
         Top             =   120
         Width           =   1455
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Disabilita"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Abilita"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Timer 1"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Intervallo (msec)"
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   120
         Width           =   1455
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Disabilita"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Abilita"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "timerform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer01.Enabled = True" & vbCrLf
End If
If Option2.Value = True Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer01.Enabled = False" & vbCrLf
End If
If IsNumeric(Text1.Text) Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer01.Interval = " & Text1.Text & vbCrLf
End If

End Sub

Private Sub Command2_Click()
If Option3.Value = True Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer02.Enabled = True" & vbCrLf
End If
If Option4.Value = True Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer02.Enabled = False" & vbCrLf
End If
If IsNumeric(Text2.Text) Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer02.Interval = " & Text2.Text & vbCrLf
End If

End Sub

Private Sub Command3_Click()
If Option5.Value = True Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer03.Enabled = True" & vbCrLf
End If
If Option6.Value = True Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer03.Enabled = False" & vbCrLf
End If
If IsNumeric(Text3.Text) Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Timer03.Interval = " & Text3.Text & vbCrLf
End If

End Sub

Private Sub Form_Load()
Me.Icon = ScriptEditor.ImageList1.ListImages(9).Picture
End Sub

