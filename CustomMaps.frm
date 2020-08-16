VERSION 5.00
Begin VB.Form CMForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Maps"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Map Pack Editor"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Map Editor"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Single Maps"
      Height          =   3135
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      Begin VB.FileListBox File2 
         Height          =   285
         Left            =   1680
         MultiSelect     =   1  'Simple
         Pattern         =   "*.TSM"
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox List2 
         Height          =   2205
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start Game"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   2640
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Packs"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   1680
         Pattern         =   "*.TSS"
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ListBox List1 
         Height          =   2205
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Game"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   2640
         Width           =   1815
      End
   End
End
Attribute VB_Name = "CMForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If List1.List(List1.ListIndex) = "" Then
    MsgBox "Select a Map Pack from the list first."
    GoTo ExitSub
End If

    IsFirstCM = True
    CustomMaps = MapPaths(List1.ListIndex)
    CMCounter = 0
    IsCustomMap = True
    MainForm.LoadCustomMap
    Me.Hide
    MainMenu.Hide
    MainForm.Enabled = True
    MainForm.SetFocus
    MainForm.Timer1.Interval = 15
    
ExitSub:

End Sub

Private Sub Command2_Click()

If List2.List(List2.ListIndex) = "" Then
    MsgBox "Select a map from the list first."
    GoTo ExitSub
End If

SPCMmap = List2.List(List2.ListIndex)

    IsFirstCM = True
    CustomMaps = List2.List(List2.ListIndex)
    
    CMCounter = 0
    IsSPCM = True
    MainForm.LoadCustomSPMap
    Me.Hide
    MainMenu.Hide
    MainForm.Enabled = True
    MainForm.SetFocus
    MainForm.Timer1.Interval = 15
    
ExitSub:

End Sub

Private Sub Command3_Click()
On Error GoTo NoFile
    Shell ("MapEditor\TSEdit20.exe"), vbNormalFocus
    Me.Show
    Exit Sub
NoFile:
MsgBox "Map Editor not found!", vbCritical, "Error"
End Sub

Private Sub Command4_Click()
On Error GoTo NoFile
    Shell ("MapEditor\TLSE.exe"), vbNormalFocus
    Me.Show
    Exit Sub
NoFile:
MsgBox "Map Pack Maker not found!", vbCritical, "Error"
End Sub

Private Sub Command5_Click()
    MainMenu.Enabled = True
    Me.Hide
        
End Sub

Private Sub Form_Load()
Dim MapContent As String
Dim MapSplit() As String


    File2.Path = ("Custom_Maps\")
    LoadCSPMList
    File1.Path = ("Map_Packs\")

    ReDim Preserve MapPaths(0 To File1.ListCount)
        
        For U = 0 To File1.ListCount - 1
        
            Open "Map_Packs\" & File1.List(U) For Input As #1
            
                Do Until EOF(1)
                
                       Line Input #1, MapContent
                       
                       MapSplit() = Split(MapContent, ";")

                       List1.AddItem MapSplit(0)
                       
                       For B = 1 To UBound(MapSplit())
                       
                             MapPaths(U) = MapPaths(U) & vbCrLf & MapSplit(B)
                             
                             
                       Next
                       

                Loop
                     
            Close #1
        Next
        

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    MainMenu.Enabled = True
    Unload Me
    
End Sub

Function LoadCSPMList()

Dim NameStr() As String

    For T = 0 To File2.ListCount - 1
        NameStr() = Split(File2.List(T), ".")
        List2.AddItem NameStr(0)
    Next
    
End Function

