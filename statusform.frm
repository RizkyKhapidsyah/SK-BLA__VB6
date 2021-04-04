VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form statusform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spy Form"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Anteprima"
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   1200
      Width           =   3735
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Info Window"
         Height          =   375
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Command Line"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Info Window"
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "Inserisci"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "ForeColor"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "BackColor"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Command Line"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "Inserisci"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "ForeColor"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "BackColor"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "statusform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "SpyStatus.spycommand.BackColor = " & Label2.BackColor & vbCrLf
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "SpyStatus.spycommand.Forecolor = " & Label3.BackColor & vbCrLf
End Sub

Private Sub Form_Load()
Me.Icon = ScriptEditor.ImageList1.ListImages(1).Picture
Label8.BackColor = Label2.BackColor
Label8.ForeColor = Label3.BackColor
Label9.BackColor = Label5.BackColor
Label9.ForeColor = Label6.BackColor

End Sub

Private Sub Label2_Click()
CommonDialog1.ShowColor
Label2.BackColor = CommonDialog1.Color
Label8.BackColor = CommonDialog1.Color
End Sub

Private Sub Label3_Click()
CommonDialog1.ShowColor
Label3.BackColor = CommonDialog1.Color
Label8.ForeColor = CommonDialog1.Color
End Sub


Private Sub Label5_Click()
CommonDialog1.ShowColor
Label5.BackColor = CommonDialog1.Color
Label9.BackColor = CommonDialog1.Color
End Sub

Private Sub Label6_Click()
CommonDialog1.ShowColor
Label6.BackColor = CommonDialog1.Color
Label9.ForeColor = CommonDialog1.Color
End Sub

