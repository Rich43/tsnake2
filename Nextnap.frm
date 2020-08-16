VERSION 5.00
Begin VB.Form NextmapForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   0
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   3855
   End
End
Attribute VB_Name = "NextmapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Seconds As Integer

Private Sub Form_Load()

    MainForm.Timer1.Interval = 0
    MainForm.Enabled = False
    Seconds = 3
    Label1.Caption = "Next Map in 3 seconds..."

End Sub

Private Sub Timer1_Timer()

    Seconds = Seconds - 1
    Label1.Caption = "Next Map in " & Seconds & " seconds..."
    
    If Seconds = 0 Then
        MainForm.Enabled = True
        MainForm.Timer1.Interval = 15
        
            If IsCustomMap = False Then
                MainForm.NextMap
            Else
                MainForm.LoadCustomMap
            End If
            
        Unload Me
    End If
    
End Sub
