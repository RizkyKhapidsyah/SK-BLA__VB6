VERSION 5.00
Begin VB.Form Intro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bla"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ForeColor       =   &H00000000&
   Icon            =   "Intro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      Height          =   2655
      Left            =   0
      Picture         =   "Intro.frx":0E42
      ScaleHeight     =   173
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label protectsystemcopyright2001 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   3855
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Intro.protectsystemcopyright2001.Caption = "Ideato e sviluppato da: Undertacker e Pixel"
End Sub

Private Sub pic_Click()
main.Show
Unload Me
End Sub

