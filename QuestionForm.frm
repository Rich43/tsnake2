VERSION 5.00
Begin VB.Form QuestionForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Main Menu"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "QuestionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    YN = "Yes"
    MainForm.FormUnload
    Unload Me
End Sub

Private Sub Command2_Click()
    MainForm.Enabled = True
    YN = "No"
    MainForm.FormUnload
    Unload Me
End Sub

Private Sub Command3_Click()
    MainForm.Timer1.Interval = 0
    Me.Enabled = False
    MainMenu.Show
    MainMenu.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Enabled = True
    MainForm.Enabled = False
    MainForm.Timer1.Enabled = False
    Me.Caption = "Quit?"
    Label1.Caption = "Are you sure you want to quit?"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    YN = "No"
    MainForm.FormUnload
    Unload Me
    IsCustomMap = False
End Sub
