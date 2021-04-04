VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form SubForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Routine"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Descrizione"
      Height          =   3015
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   5295
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4683
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"SubForm.frx":0000
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
   Begin VB.Frame Frame1 
      Caption         =   "Indice"
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   2520
         Width           =   855
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2160
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "SubForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ScriptEditor.editor.SelBold = False
ScriptEditor.editor.SelColor = vbBlue

If List1.ListIndex = 0 Then
    ScriptEditor.editor.SelText = "Sub MAIN()" & vbCrLf & vbCrLf & "End Sub" & vbCrLf
End If

If List1.ListIndex = 1 Then
    ScriptEditor.editor.SelText = "Sub SOCKCONNECT()" & vbCrLf & vbCrLf & "End Sub" & vbCrLf
End If

If List1.ListIndex = 2 Then
ScriptEditor.editor.SelText = "Sub TBUTTON"
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlack
ScriptEditor.editor.SelText = "?"
ScriptEditor.editor.SelBold = False
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = " ()" & vbCrLf & vbCrLf
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "End Sub" & vbCrLf
End If



End Sub

Private Sub Form_Load()
Me.Icon = ScriptEditor.ImageList1.ListImages(2).Picture
Dim a As String
Open App.Path + "\vbscript\indsub.txt" For Input As #1
Do While Not EOF(1)
Line Input #1, a
List1.AddItem a
Loop
Close #1
End Sub

Private Sub List1_Click()
RichTextBox1.LoadFile App.Path + "\vbscript\" + Right$(Str(List1.ListIndex), Len(Str(List1.ListIndex)) - 1) + ".txt", rtfText
End Sub
