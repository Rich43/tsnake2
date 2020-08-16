VERSION 5.00
Begin VB.Form OptionForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   2775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Themes"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command2 
         Caption         =   "Ok"
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "OptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ApplyTheme
End Sub

Private Sub Command2_Click()
    ApplyTheme
    MainMenu.Enabled = True
    Me.Hide
End Sub

Function ApplyTheme()

Select Case List1.List(List1.ListIndex)

    Case "Default Theme"
            MainForm.EndSprite.Picture = LoadPicture("GFX\Snakes\Def_End_Sprite.bmp")
            MainForm.EndMask.Picture = LoadPicture("GFX\Snakes\Def_End_Mask.bmp")
        
            MainForm.TailSprite.Picture = LoadPicture("GFX\Snakes\Def_Tail_Sprite.bmp")
            MainForm.TailMask.Picture = LoadPicture("GFX\Snakes\Def_Tail_Mask.bmp")
            
            MainForm.HeadSprite.Picture = LoadPicture("GFX\Snakes\Def_Head_Sprite.bmp")
            MainForm.HeadMask.Picture = LoadPicture("GFX\Snakes\Def_Head_Mask.bmp")
        
            MainForm.FoodSprite.Picture = LoadPicture("GFX\Objects\Def_Food_Sprite.bmp")
            MainForm.FoodMask.Picture = LoadPicture("GFX\Objects\Def_Food_Mask.bmp")
            
            MainForm.GemSprite.Picture = LoadPicture("GFX\Objects\Def_Gem_Sprite.bmp")
            MainForm.GemMask.Picture = LoadPicture("GFX\Objects\Def_Gem_Mask.bmp")
            
            MainForm.WallSprite.Picture = LoadPicture("GFX\Walls\Def_Wall_Sprite.bmp")
            MainForm.WallMask.Picture = LoadPicture("GFX\Walls\Def_Wall_Mask.bmp")
        
            MainForm.Picture = LoadPicture("GFX\Backgrounds\Default.gif")
            
    Case ""

End Select



End Function

Private Sub Form_Load()
List1.AddItem "Default Theme"
End Sub
