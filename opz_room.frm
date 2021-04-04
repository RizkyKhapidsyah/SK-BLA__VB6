VERSION 5.00
Begin VB.Form opz_room 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rooms"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ClipControls    =   0   'False
   Icon            =   "opz_room.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Room"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4335
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "psw +q/+o"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Room preferite"
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4335
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   2160
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "aggiungi"
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "elimina"
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Entra"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Esci"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   3240
      Width           =   975
   End
End
Attribute VB_Name = "opz_room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub Command12_Click()
If Text1.Text = "" Then
MsgBox "Non è stata selezionata nessuna Room.", vbExclamation
Exit Sub
End If
If main.Winsock1.State = 7 Then
main.Winsock1.SendData "JOIN " & Text1.Text & " " & Text2.Text & vbCrLf
main.Enabled = True
Unload Me
Else
MsgBox "Non c'è nessuna connessione in atto.", vbExclamation
End If


End Sub

Private Sub Command14_Click()
main.Enabled = True
Unload Me

End Sub

Private Sub Command3_Click()
Dim a As String
Open App.Path + "\archivi\arcroom.arc" For Input As #2
Do While Not EOF(2)
Input #2, a
If a = Text1.Text Then
MsgBox "Room già esistente.": Close #2: Exit Sub
End If
Loop
Close #2
Open App.Path + "\archivi\arcroom.arc" For Append As #1
Write #1, Text1
Close #1
List1.AddItem Text1

End Sub

Private Sub Command4_Click()
Dim b1 As String
Open App.Path + "\archivi\arcroom.tmp" For Output As #2
Open App.Path + "\archivi\arcroom.arc" For Input As #1
Do While Not EOF(1)
Input #1, b1
If b1 <> List1.Text Then Write #2, b1
Loop
Close #1
Close #2
Kill App.Path + "\archivi\arcroom.arc"
FileCopy App.Path + "\archivi\arcroom.tmp", App.Path + "\archivi\arcroom.arc"
Kill App.Path + "\archivi\arcroom.tmp"
List1.RemoveItem List1.ListIndex

End Sub

Private Sub Form_Load()
Dim a1 As String
Open App.Path + "\archivi\arcroom.arc" For Input As #1
Do While Not EOF(1)
Input #1, a1
List1.AddItem a1
Loop
Close #1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
main.Enabled = True
Unload Me

End Sub

Private Sub List1_DblClick()
Text1.Text = List1.Text
End Sub
