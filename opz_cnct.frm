VERSION 5.00
Begin VB.Form opz_cnct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impostazioni di Connessione"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ClipControls    =   0   'False
   Icon            =   "opz_cnct.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "Esci"
      Height          =   375
      Left            =   2760
      TabIndex        =   22
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Applica"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Server"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   3615
      Begin VB.CommandButton Command8 
         Caption         =   "aggiungi"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "elimina"
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.ListBox List2 
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
         Height          =   900
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "default"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text3 
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
         Left            =   2880
         TabIndex        =   10
         Text            =   "6667"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "porta"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Impostazioni personali"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   3615
      Begin VB.CommandButton Command11 
         Caption         =   "default"
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Frame Frame6 
         Caption         =   "Nome reale"
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   2175
         Begin VB.TextBox Text5 
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
            TabIndex        =   19
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "E-mail"
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2175
         Begin VB.TextBox Text4 
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
            TabIndex        =   18
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nick secondari"
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3615
      Begin VB.CommandButton Command4 
         Caption         =   "elimina"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "aggiungi"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
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
         Height          =   900
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nick principale"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "default"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1095
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
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "opz_cnct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command11_Click()
Open App.Path + "\default\impst.dft" For Output As #1
Write #1, Text4, Text5
Close #1

End Sub

Private Sub Command12_Click()
main.Label1 = Text2.Text
main.Label2 = Text3.Text
main.Label3 = Text1.Text
main.Label5 = Text5.Text
main.Label4 = Text4.Text
main.Enabled = True
Unload Me
End Sub

Private Sub Command14_Click()
main.Enabled = True
Unload Me
End Sub

Private Sub Command2_Click()
Open App.Path + "\default\nick.dft" For Output As #1
Write #1, Text1
Close #1
End Sub

Private Sub Command3_Click()
Dim a As String
If Text1.Text <> "" Then
Open App.Path + "\archivi\arcnick.arc" For Input As #2
Do While Not EOF(2)
Line Input #2, a
If a = Text1.Text Then
MsgBox "Nick già esistente.": Close #2: Exit Sub
End If
Loop
Close #2
Open App.Path + "\archivi\arcnick.arc" For Append As #1
Print #1, Text1
Close #1
List1.AddItem Text1
End If
End Sub

Private Sub Command4_Click()
Dim b1 As String
Open App.Path + "\archivi\arcnick.tmp" For Output As #2
Open App.Path + "\archivi\arcnick.arc" For Input As #1
Do While Not EOF(1)
Line Input #1, b1
If b1 <> List1.Text Then Print #2, b1
Loop
Close #1
Close #2
Kill App.Path + "\archivi\arcnick.arc"
FileCopy App.Path + "\archivi\arcnick.tmp", App.Path + "\archivi\arcnick.arc"
Kill App.Path + "\archivi\arcnick.tmp"
List1.RemoveItem List1.ListIndex

End Sub

Private Sub Command5_Click()
Open App.Path + "\default\server.dft" For Output As #1
Write #1, Text2, Text3
Close #1

End Sub

Private Sub Command7_Click()
Dim b1, b2 As String
Open App.Path + "\archivi\arcserver.tmp" For Output As #2
Open App.Path + "\archivi\arcserver.arc" For Input As #1
Do While Not EOF(1)
Input #1, b1, b2
If b1 <> List2.Text Then Write #2, b1, b2
Loop
Close #1
Close #2
Kill App.Path + "\archivi\arcserver.arc"
FileCopy App.Path + "\archivi\arcserver.tmp", App.Path + "\archivi\arcserver.arc"
Kill App.Path + "\archivi\arcserver.tmp"
List2.RemoveItem List2.ListIndex
End Sub

Private Sub Command8_Click()
Dim a, a1 As String
If Text2.Text <> "" Then
Open App.Path + "\archivi\arcserver.arc" For Input As #2
Do While Not EOF(2)
Input #2, a, a1
If a = Text2.Text Then
MsgBox "Server già esistente.": Close #2: Exit Sub
End If
Loop
Close #2
Open App.Path + "\archivi\arcserver.arc" For Append As #1
Write #1, Text2, Text3
Close #1
List2.AddItem Text2
End If
End Sub

Private Sub Form_Load()
Dim a1, a2 As Integer
Dim b1, b2 As String
Open App.Path + "\archivi\arcserver.arc" For Input As #1
Do While Not EOF(1)
Input #1, b1, b2
List2.AddItem b1
Loop
Close #1
Open App.Path + "\archivi\arcnick.arc" For Input As #1
Do While Not EOF(1)
Line Input #1, b1
List1.AddItem b1
Loop
Close #1

Text2.Text = main.Label1
Text3.Text = main.Label2
Text1.Text = main.Label3
Text5.Text = main.Label5
Text4.Text = main.Label4
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
main.Enabled = True
Unload Me

End Sub

Private Sub List1_DblClick()
Text1.Text = List1.Text
End Sub

Private Sub List2_DblClick()
Dim b1, b2 As String
Open App.Path + "\archivi\arcserver.arc" For Input As #1
Do While Not EOF(1)
Input #1, b1, b2
If b1 = List2.Text Then
Text2.Text = b1: Text3.Text = b2: Close #1: Exit Sub
End If
Loop
Close #1

End Sub

