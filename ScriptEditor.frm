VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ScriptEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script Editor"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9960
   ClipControls    =   0   'False
   Icon            =   "ScriptEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1560
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":0CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":100E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":1952
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":1DA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":20C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":3DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":3F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ScriptEditor.frx":47FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1429
      ButtonWidth     =   1667
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sock"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SpyForm"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "IconCenter"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ToolBar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Timer"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "NoticeForm"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "PopMenu"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "UserForm"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SubRoutine"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "VBScript"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   9735
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2520
         Top             =   5280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin RichTextLib.RichTextBox editor 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   9340
         _Version        =   393217
         BackColor       =   14737632
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         TextRTF         =   $"ScriptEditor.frx":6FAE
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
   Begin VB.Menu m0 
      Caption         =   "File"
      Begin VB.Menu m0a 
         Caption         =   "Apri File"
      End
      Begin VB.Menu m0b 
         Caption         =   "Salva File"
      End
      Begin VB.Menu m0cc 
         Caption         =   "-"
      End
      Begin VB.Menu m0e 
         Caption         =   "Cancella tutto"
      End
      Begin VB.Menu m0ccc 
         Caption         =   "-"
      End
      Begin VB.Menu m0d 
         Caption         =   "Ceck Script"
      End
      Begin VB.Menu m0c 
         Caption         =   "-"
      End
      Begin VB.Menu m0z 
         Caption         =   "Esci"
      End
   End
   Begin VB.Menu m3 
      Caption         =   "Tutorial"
      Begin VB.Menu m3a 
         Caption         =   "Codici RAW"
      End
      Begin VB.Menu m3b 
         Caption         =   "Comandi IRC"
      End
      Begin VB.Menu m3c 
         Caption         =   "Comandi IRCX"
      End
   End
End
Attribute VB_Name = "ScriptEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub editor_KeyPress(KeyAscii As Integer)
editor.SelBold = True
editor.SelColor = vbBlack
End Sub
Private Sub Form_Load()

'>>'script.AddObject "Timer01", Timer1, True
'>>'script.AddObject "Timer02", Timer2, True
'>>'script.AddObject "Timer03", Timer3, True
'>>'script.AddObject "Sock", Winsock1, True
'>>'script.AddObject "SpyStatus", status, True
'>>'script.AddObject "NoticeForm", Notice, True
'>>'script.AddObject "Toolbar", Toolbar1, True
'>>'script.AddObject "Iconcenter", ImageList1, True
'>>'script.AddObject "UserPopMenu1", Popmenu1, True'>>
'>>'script.AddObject "UserPopMenu2", Popmenu2, True
'>>'script.AddObject "UserPopMenu3", Popmenu3, True
'>>'script.AddObject "UserForm1", UserForm, True
'script.AddObject "ServerName", Label1, True
'script.AddObject "ServerPort", Label2, True
'script.AddObject "UserNick", Label3, True
''script.AddObject "RoomControl", Room1, True
'script.AddObject "Usermail", Label4, True
'script.AddObject "Username", Label5, True
'script.AddCode mainscript.Text

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
main.Enabled = True
End Sub


Private Sub m0a_Click()
CommonDialog1.Filter = "INI |*.ini|TXT |*.txt|(tutti) |*.*"
CommonDialog1.DialogTitle = "Apertura file script"
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.Flags = cdlOFNNoChangeDir
CommonDialog1.InitDir = App.Path + "\script"
CommonDialog1.FileName = ""
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
editor.LoadFile CommonDialog1.FileName, rtfText
Frame1.Caption = CommonDialog1.FileName
End If
End Sub

Private Sub m0b_Click()
CommonDialog1.Filter = "INI |*.ini|TXT |*.txt|(tutti) |*.*"
CommonDialog1.DialogTitle = "Salvataggio file script"
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.InitDir = App.Path + "\script"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
editor.SaveFile CommonDialog1.FileName, rtfText
Frame1.Caption = CommonDialog1.FileName
End If

End Sub

Private Sub m0d_Click()
   Dim Msg
On Error Resume Next
 main.script.AddCode editor.Text

If main.script.Error.Number <> 0 Then
     Msg = "Errore dello script." & vbNewLine
      Msg = Msg & "Errore " & main.script.Error.Number
      Msg = Msg & vbCrLf
      Msg = Msg & " (" & main.script.Error.Description & ")"
      Msg = Msg & " nella riga: "
      Msg = Msg & main.script.Error.Line
      MsgBox Msg
      main.script.Reset
Exit Sub
End If

main.script.Reset
MsgBox "Codice privo di errori di sintassi."
End Sub
Private Sub m0e_Click()
editor.Text = ""
Me.Frame1.Caption = ""
Me.ImageList2.ListImages.Clear
End Sub
Private Sub m0z_Click()
main.Enabled = True
Unload Me
End Sub
Private Sub m3a_Click()
info1.RichTextBox1.LoadFile App.Path + "\tutorial\numerics.txt", rtfText
info1.Show
End Sub
Private Sub m3b_Click()
info1.RichTextBox1.LoadFile App.Path + "\tutorial\irc.txt", rtfText
info1.Show

End Sub
Private Sub m3c_Click()
info1.RichTextBox1.LoadFile App.Path + "\tutorial\ircx.txt", rtfText
info1.Show

End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
sockform.Show

Case 2
statusform.Show

Case 3
iconform.Show

Case 4
toolform.Show

Case 5
timerform.Show

Case 10
SubForm.Show

End Select
End Sub
