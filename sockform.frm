VERSION 5.00
Begin VB.Form sockform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sock"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Invia dati"
      Height          =   975
      Left            =   0
      TabIndex        =   24
      Top             =   4200
      Width           =   4815
      Begin VB.CommandButton Command4 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   3720
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame10 
         Caption         =   "Stringa"
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   3495
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   3255
         End
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Login utente"
      Height          =   1575
      Left            =   0
      TabIndex        =   16
      Top             =   2640
      Width           =   4815
      Begin VB.CheckBox Check2 
         Caption         =   "Valori da Bla"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.Frame Frame8 
         Caption         =   "Mail"
         Height          =   615
         Left            =   2400
         TabIndex        =   21
         Top             =   840
         Width           =   1215
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Name"
         Height          =   615
         Left            =   2400
         TabIndex        =   19
         Top             =   240
         Width           =   1215
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Nick"
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2175
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connessione / Disconnessione"
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   4815
      Begin VB.CommandButton Command5 
         Caption         =   "Chiudi Sock"
         Height          =   375
         Left            =   3720
         TabIndex        =   31
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Valori da Bla"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   3720
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Porta"
         Height          =   615
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   1215
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Server"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2175
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1935
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stato"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.OptionButton Option1 
         Caption         =   "Chiuso"
         Height          =   255
         Index           =   8
         Left            =   3720
         TabIndex        =   30
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Inserisci"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Errore"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Connessione in chiusura"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   7
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Connesso"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Host risolto"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Risoluzione Host in corso"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Connessione in corso"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In attesa"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Aperto"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "sockform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim zix As String
Dim zix1 As String
zix = "ServerName.Caption"
zix1 = "ServerPort.Caption"
If Check1.Value = 0 Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Sock.Connect"
ScriptEditor.editor.SelText = " " & Chr$(34) & Text1.Text & Chr$(34) & ", " & Text2.Text
Else
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Sock.Connect"
ScriptEditor.editor.SelText = " " & zix & ", " & zix1
End If
End Sub

Private Sub Command2_Click()
Dim zix As String
Dim zix1 As String
Dim zix2 As String
zix = "UserNick.caption"
zix1 = "UserMail.caption"
zix2 = "UserName.caption"
If Check2.Value = 0 Then
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Sock.sendData " & Chr$(34) & "User " & Text5.Text & " " & Chr$(34) & " & Sock.LocalHostName & " & Chr$(34) & " " & Chr$(34) & " & Sock.RemoteHost &" & Chr$(34) & " :" & Text4.Text & Chr$(34) & " & vbcrlf" & vbCrLf
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Sock.sendData " & Chr$(34) & "NICK " & Text3.Text & Chr$(34) & " & vbCrlf" & vbCrLf
Else
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Sock.sendData " & Chr$(34) & "User " & zix1 & " " & Chr$(34) & " & Sock.LocalHostName & " & Chr$(34) & " " & Chr$(34) & " & Sock.RemoteHost &" & Chr$(34) & " :" & zix2 & Chr$(34) & " & vbcrlf" & vbCrLf
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Sock.sendData " & Chr$(34) & "NICK " & zix & Chr$(34) & " & vbCrlf" & vbCrLf
End If

End Sub

Private Sub Command3_Click()
If Option1(0).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 1 Then " & vbCrLf
End If
If Option1(1).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 2 Then " & vbCrLf
End If
If Option1(2).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 6 Then "
End If
If Option1(3).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 4 Then "
End If
If Option1(4).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 5 Then "
End If
If Option1(5).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 7 Then "
End If
If Option1(6).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 8 Then "
End If
If Option1(7).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 9 Then "
End If
If Option1(8).Value = True Then
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "If Sock.State = 9 Then "
End If


End Sub

Private Sub Command4_Click()
ScriptEditor.editor.SelBold = True
ScriptEditor.editor.SelColor = vbBlue
ScriptEditor.editor.SelText = "Sock.sendData " & Chr$(34) & Text6.Text & Chr$(34) & " & vbCrlf" & vbCrLf
End Sub

Private Sub Command5_Click()
    ScriptEditor.editor.SelBold = True
    ScriptEditor.editor.SelColor = vbBlue
    ScriptEditor.editor.SelText = "Sock.Close" & vbCrLf

End Sub

Private Sub Form_Load()
Me.Icon = ScriptEditor.ImageList1.ListImages(10).Picture
End Sub
