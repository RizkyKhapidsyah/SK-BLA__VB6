VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form iconform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Center"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   2955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Icona da Aggiungere"
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         ScaleHeight     =   735
         ScaleWidth      =   975
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Archivio Icone"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "iconform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.Filter = "ICO |*.ico"
CommonDialog1.DialogTitle = "Icone disponibili"
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.Flags = cdlOFNNoChangeDir
CommonDialog1.FileName = ""
CommonDialog1.InitDir = App.Path + "\icone"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
End If

End Sub

Private Sub Command2_Click()
If CommonDialog1.FileName = "" Then Exit Sub
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Iconcenter.ListImages.Add , , loadpicture(" & Chr$(34) & CommonDialog1.FileName & Chr$(34) & ")" & vbCrLf
ScriptEditor.ImageList2.ListImages.Add , , LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Form_Load()
Me.Icon = ScriptEditor.ImageList1.ListImages(3).Picture
End Sub
