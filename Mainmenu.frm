VERSION 5.00
Begin VB.Form MainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T-Snake 2.0 Main Menu"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Custom Maps"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Options"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Top Scores"
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "New Game"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   3975
      Begin VB.CommandButton SPButton 
         Caption         =   "Start Singleplayer"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton MPButton 
         Caption         =   "Multiplayer"
         Height          =   735
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox PlayerName 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Player name:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Me.Enabled = False
    CMForm.Show
End Sub

Private Sub Command2_Click()
    Me.Enabled = False
    OptionForm.Show
End Sub

Private Sub Form_Load()

    MainForm.Show
    MainForm.Enabled = False
    MainForm.Timer1.Enabled = False
    MainForm.LoadGFX
    Me.Show
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MenuQuit = True
    Unload MainForm
End Sub

Private Sub MPButton_Click()
    MainForm.Hide
    MPForm.Show
    MPForm.Enabled = False
    MPMenu.Show
    Me.Enabled = False
    Me.Hide
End Sub

Private Sub SPButton_Click()

    MainForm.Enabled = True
    Me.Hide
    MainForm.Show
    MainForm.ResetGame
    
End Sub

Private Sub Command3_Click()

End Sub
