VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48BC9515-0F01-11D5-899A-525400EB4A52}#7.0#0"; "xroom.ocx"
Object = "{A5FABD73-D4BC-11D4-899A-525400EB4A52}#31.0#0"; "ircEngine.ocx"
Begin VB.Form main 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Bla"
   ClientHeight    =   4455
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   1560
   ClipControls    =   0   'False
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   1560
   Begin VB.Timer Timer3 
      Left            =   840
      Top             =   3000
   End
   Begin VB.Timer Timer2 
      Left            =   840
      Top             =   2640
   End
   Begin xroom.Room Room1 
      Left            =   1200
      Top             =   1080
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin ircEng.ircEngine ircEngine1 
      Left            =   1200
      Top             =   1440
      _ExtentX        =   397
      _ExtentY        =   397
   End
   Begin VB.Timer Timer1 
      Left            =   840
      Top             =   2280
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":125E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":15FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "main.frx":1A16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   741
      ButtonWidth     =   661
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox mainscript 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"main.frx":22F2
   End
   Begin MSScriptControlCtl.ScriptControl script 
      Left            =   0
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.Label Label5 
      Caption         =   "Bla"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Bla"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "BlaUser"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "6667"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "irc.ozmatrix.com"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu m0 
      Caption         =   "File"
      Begin VB.Menu m0a 
         Caption         =   "Carica Script"
      End
      Begin VB.Menu m0x 
         Caption         =   "-"
      End
      Begin VB.Menu m0b 
         Caption         =   "Resetta Script"
         Enabled         =   0   'False
      End
      Begin VB.Menu m0xx 
         Caption         =   "-"
      End
      Begin VB.Menu m0z 
         Caption         =   "Esci"
      End
   End
   Begin VB.Menu m1 
      Caption         =   "Utilità"
      Begin VB.Menu m1a 
         Caption         =   "Script Editor"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim a, a1 As String
cekrefresh = False
Me.Left = Screen.Width - 1680
Me.Top = 0
Open App.Path + "\default\nick.dft" For Input As #1
Input #1, a
Close #1
Label3.Caption = a

Open App.Path + "\default\server.dft" For Input As #1
Input #1, a, a1
Close #1
Label1.Caption = a
Label2.Caption = a1

Open App.Path + "\default\impst.dft" For Input As #1
Input #1, a, a1
Close #1
Label4.Caption = a
Label5.Caption = a1


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
script.Reset
Winsock1.Close
Unload Me
End Sub

Private Sub ircEngine1_NickAction(Nick As String, Destination As String, Msg As String)
Room1.RoomMsg Nick, LCase(Destination), Msg
End Sub

Private Sub ircEngine1_NickChange(OldNick As String, NewNick As String)
Room1.NickChangeNick OldNick, NewNick
If OldNick = Label3.Caption Then Label3.Caption = NewNick
End Sub

Private Sub ircEngine1_NickJoin(Nick As String, Room As String, FullName As String)
Room1.ShowRoom LCase(Room)
Room1.NickJoin Nick, LCase(Room)
End Sub


Private Sub ircEngine1_NickMode(Nick As String, Room As String, Destination As String, q As Boolean, o As Boolean, v As Boolean, DataString As String)
If q = True Then Room1.NickChangeOpt Destination, LCase(Room), ".": Exit Sub
If o = True Then Room1.NickChangeOpt Destination, LCase(Room), "@": Exit Sub
If v = True Then Room1.NickChangeOpt Destination, LCase(Room), "+": Exit Sub
Room1.NickChangeOpt Destination, LCase(Room), ""
End Sub

Private Sub ircEngine1_NickMsg(Nick As String, Room As String, Msg As String)
Dim t As Integer
    If LCase(Room) = LCase(Label3.Caption) Then
        For t = 1 To 100
            If QueryName(t) = LCase(Nick) Then Room1.QueryMsg Nick, LCase(Nick), Msg: Exit Sub
        Next t
            
        For t = 1 To 100
            If QueryName(t) = "" Then
            QueryName(t) = LCase(Nick)
            Room1.ShowQuery LCase(Nick)
            Room1.QueryMsg Nick, LCase(Nick), Msg
            Exit Sub
            End If
        Next t
     Exit Sub
     End If

Room1.RoomMsg Nick, LCase(Room), Msg
script.Run "Roommessage", Nick, LCase(Room), Msg
End Sub

Private Sub ircEngine1_NickNotice(Nick As String, Destination As String, Msg As String)
Notice.RichTextBox1.SelStart = Len(Notice.RichTextBox1.Text)
Notice.RichTextBox1.SelText = Nick & " :" & Msg & vbCrLf
End Sub

Private Sub ircEngine1_NickPart(Nick As String, Room As String)
Room1.NickPart Nick, LCase(Room)
End Sub

Private Sub ircEngine1_NickQuit(Nick As String, Msg As String)
Room1.NickQuit Nick
End Sub


Private Sub ircEngine1_RefreshNick(Room As String, stato As Boolean)
If stato = True Then
cekrefresh = True
Winsock1.SendData "NAMES " & LCase(Room) & vbCrLf
End If
End Sub

Private Sub ircEngine1_RoomNames(Room As String, Nicks() As String, numnick As Long)
For t = 0 To numnick
If Left$(Nicks(t), 1) = "." Then
Room1.NickJoin Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room)
Room1.NickChangeOpt Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room), "."
End If
If Left$(Nicks(t), 1) = "@" Then
Room1.NickJoin Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room)
Room1.NickChangeOpt Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room), "@"
End If
If Left$(Nicks(t), 1) = "+" Then
Room1.NickJoin Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room)
Room1.NickChangeOpt Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room), "+"
End If
If Left$(Nicks(t), 1) <> "." And Left$(Nicks(t), 1) <> "@" And Left$(Nicks(t), 1) <> "+" Then
Room1.NickJoin Nicks(t), LCase(Room)
End If
Next t

' refresh
If cekrefresh = True Then
For t = 0 To numnick
If Left$(Nicks(t), 1) = "." Then
Room1.NickChangeOpt Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room), "."
End If
If Left$(Nicks(t), 1) = "@" Then
Room1.NickChangeOpt Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room), "@"
End If
If Left$(Nicks(t), 1) = "+" Then
Room1.NickChangeOpt Right$(Nicks(t), Len(Nicks(t)) - 1), LCase(Room), "+"
End If
If Left$(Nicks(t), 1) <> "." And Left$(Nicks(t), 1) <> "@" And Left$(Nicks(t), 1) <> "+" Then
Room1.NickChangeOpt Nicks(t), LCase(Room), ""
End If
Next t
Room1.RoomRefresh LCase(Room)
cekrefresh = False
End If

End Sub


Private Sub ircEngine1_RoomSetting(Room As String, Setting As String)
Room1.RoomSet LCase(Room), Setting
End Sub

Private Sub ircEngine1_RoomTopic(Room As String, Topic As String)
Room1.RoomTopic LCase(Room), Topic
Room1.RoomMsg "", LCase(Room), Topic
End Sub

Private Sub ircEngine1_ServerCommand(nicka As String, command As String, nickb As String, par1 As String, par2 As String, par3 As String, DataString As String)
script.Run "Servercomand", nicka, command, nickb, par1, par2, par3, DataString
End Sub

Private Sub ircEngine1_ServerMsg(DataString As String)
script.Run "ServerMsg", DataString
End Sub

Private Sub ircEngine1_ServerNumeric(Server As String, num As String, Nick As String, canale As String, DataString As String)
script.Run "Servernumeric", Server, num, Nick, canale, DataString
End Sub

Private Sub ircEngine1_WhoPing(Name As String)
Winsock1.SendData "PONG " & Right$(Name, Len(Name) - 1) & vbCrLf
End Sub

Private Sub m0a_Click()
CommonDialog1.Filter = "INI |*.ini|TXT |*.txt|(tutti) |*.*"
CommonDialog1.DialogTitle = "Apertura file script"
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.Flags = cdlOFNNoChangeDir
CommonDialog1.FileName = ""
CommonDialog1.InitDir = App.Path + "\script"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
mainscript.LoadFile CommonDialog1.FileName, rtfText
script.AddObject "Timer01", Timer1, True
script.AddObject "Timer02", Timer2, True
script.AddObject "Timer03", Timer3, True
script.AddObject "Sock", Winsock1, True
script.AddObject "SpyStatus", status, True
script.AddObject "NoticeForm", Notice, True
script.AddObject "Toolbar", Toolbar1, True
script.AddObject "Iconcenter", ImageList1, True
script.AddObject "UserPopMenu1", Popmenu1, True
script.AddObject "UserPopMenu2", Popmenu2, True
script.AddObject "UserPopMenu3", Popmenu3, True
script.AddObject "UserForm1", UserForm, True
script.AddObject "ServerName", Label1, True
script.AddObject "ServerPort", Label2, True
script.AddObject "UserNick", Label3, True
script.AddObject "RoomControl", Room1, True
script.AddObject "Usermail", Label4, True
script.AddObject "Username", Label5, True
script.AddCode mainscript.Text
script.Run "main"
m0a.Enabled = False
m0b.Enabled = True
m1a.Enabled = False
End If

End Sub

Private Sub m0b_Click()
script.Reset

For t = Toolbar1.Buttons.Count To 1 Step -1
Toolbar1.Buttons.Remove t
Next t
m0b.Enabled = False
m0a.Enabled = True
m1a.Enabled = True
End Sub

Private Sub m0z_Click()
script.Reset
Unload Me
End Sub

Private Sub m1a_Click()
ScriptEditor.Show
main.Enabled = False
End Sub

Private Sub Room1_PopmenuCommandLine(Room As String, Msg As String)
script.Run "PMenu1", LCase(Room), Msg
End Sub

Private Sub Room1_PopmenuListNick(Room As String, Nick As String)
script.Run "PMenu2", LCase(Room), Nick
End Sub

Private Sub Room1_PopmenuTextRoom(Room As String)
script.Run "PMenu3", LCase(Room)
End Sub

Private Sub Room1_QueryPart(Query As String)
Dim t As Integer
       
       For t = 1 To 100
           If QueryName(t) = LCase(Query) Then
           Room1.UnloadQuery LCase(Query)
           QueryName(t) = ""
           Exit For
           End If
        Next t
            
End Sub

Private Sub Room1_RoomPart(Room As String)
script.Run "RoomPart", LCase(Room)
End Sub

Private Sub Room1_UserMsg(Room As String, Msg As String)
If Left$(Msg, 1) <> "\" Then
Room1.RoomMsg Label3.Caption, LCase(Room), Msg
Winsock1.SendData "PRIVMSG " & Room & " :" & Msg & vbCrLf
Else
Winsock1.SendData Right$(Msg, Len(Msg) - 1) & vbCrLf
End If
End Sub

Private Sub Room1_UserQueryMsg(Room As String, Msg As String)
If Left$(Msg, 1) <> "\" Then
Room1.QueryMsg Label3.Caption, LCase(Room), Msg
Winsock1.SendData "PRIVMSG " & Room & " :" & Msg & vbCrLf
Else
Winsock1.SendData Right$(Msg, Len(Msg) - 1) & vbCrLf
End If

End Sub

Private Sub Timer1_Timer()
script.Run "Timer_01"
End Sub

Private Sub Timer2_Timer()
script.Run "Timer_02"
End Sub

Private Sub Timer3_Timer()
script.Run "Timer_03"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
script.Run "TButton" + Right$(Str(Button.Index), Len(Str(Button.Index) - 1))
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
opz_cnct.Show
main.Enabled = False
Case 2
opz_room.Show
main.Enabled = False
'If Winsock1.State <> 0 Then Winsock1.Close
'Winsock1.Connect ServerName, ServerPort
End Select
End Sub

Private Sub Winsock1_Connect()
script.Run "Sockconnect"
'Winsock1.SendData "User " & UserMail & " " & Winsock1.LocalHostName & " " & Winsock1.RemoteHost & " :" & UserName & vbCrLf
'Winsock1.SendData "NICK " & UserNick & vbCrLf
'Winsock1.SendData "ircx" & vbCrLf
'If cek_invi = 1 Then TCP1.SendData "MODE " & NickName & " +i" & vbCrLf

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim DataString As String
Winsock1.GetData DataString
ircEngine1.Winsockstring DataString

'script.Run "Sockdatarrival", DataString
'Call cekdati(Datastring)
End Sub

