VERSION 5.00
Begin VB.Form Impostazioni 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impostazioni"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2670
   Icon            =   "Impostazioni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   2670
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Chiudi"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Mail"
      Height          =   615
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   1695
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Text5"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Name"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Text4"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Applica"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Nick"
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   1695
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Porta"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   0
      Width           =   975
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Impostazioni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
main.Label1 = Text1.Text
main.Label2 = Text2.Text
main.Label3 = Text3.Text
main.Label5 = Text4.Text
main.Label4 = Text5.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = main.Label1
Text2.Text = main.Label2
Text3.Text = main.Label3
Text4.Text = main.Label5
Text5.Text = main.Label4

End Sub
