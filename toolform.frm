VERSION 5.00
Begin VB.Form toolform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tool Bar"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   2760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Inserisci"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   3480
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Proprietà Bottone"
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   2400
      Width           =   2655
      Begin VB.Frame Frame4 
         Caption         =   "Icona"
         Height          =   615
         Left            =   1800
         TabIndex        =   5
         Top             =   240
         Width           =   735
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Descrizione"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Elenco Icone Inserite"
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1215
         Left            =   1200
         ScaleHeight     =   1155
         ScaleWidth      =   1275
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "toolform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Toolbar.Buttons.Add , , " & Chr$(34) & Text1.Text & Chr$(34) & ",," & Text2.Text & vbCrLf

End Sub

Private Sub Form_Load()
Me.Icon = ScriptEditor.ImageList1.ListImages(8).Picture
For t = 0 To ScriptEditor.ImageList2.ListImages.Count - 1
List1.AddItem t + 1
Next t
End Sub

Private Sub List1_Click()
Picture1.Picture = ScriptEditor.ImageList2.ListImages(List1.ListIndex + 1).Picture
Text2.Text = List1.ListIndex + 1
End Sub
