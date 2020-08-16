VERSION 5.00
Begin VB.Form MPMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T-Snake  2 - Multiplayer"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   255
   End
End
Attribute VB_Name = "MPMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MPForm.Winsock1.RemoteHost = "81.168.125.66"
    MPForm.Winsock1.RemotePort = 10101
    MPForm.Winsock1.Connect
End Sub

Private Sub Command2_Click()
    Me.Hide
    MainMenu.Show
    MainMenu.Enabled = True
End Sub
