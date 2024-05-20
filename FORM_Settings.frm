VERSION 5.00
Begin VB.Form FORM_Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FracGen - Settings"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   Icon            =   "FORM_Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      Caption         =   "RGB Modifier Settings"
      Height          =   1335
      Left            =   3000
      TabIndex        =   55
      Top             =   5280
      Width           =   2775
      Begin VB.ComboBox DROPDOWN_RGB_B 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_RGB_G 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_RGB_R 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LABEL_RGB_B 
         Caption         =   "B*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   960
         Width           =   375
      End
      Begin VB.Label LABEL_RGB_G 
         Caption         =   "G*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   600
         Width           =   375
      End
      Begin VB.Label LABEL_RGB_R 
         Caption         =   "R*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Circle Settings (ABS for negatives)"
      Height          =   1695
      Left            =   3000
      TabIndex        =   47
      Top             =   960
      Width           =   2775
      Begin VB.ComboBox DROPDOWN_ZIndex 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_r 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_yc 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_xc 
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LABEL_ZIndex 
         Caption         =   "zI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label LABEL_r 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   375
      End
      Begin VB.Label LABEL_yc 
         Caption         =   "yc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   600
         Width           =   375
      End
      Begin VB.Label LABEL_xc 
         Caption         =   "xc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Line Settings (x1,y1,z1) - (x2, y2, z2)"
      Height          =   2415
      Left            =   3000
      TabIndex        =   43
      Top             =   2760
      Width           =   2775
      Begin VB.ComboBox DROPDOWN_LINE_z2 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_LINE_y2 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_LINE_x2 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_LINE_z1 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_LINE_y1 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_LINE_x1 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LABEL_LINE_z2 
         Caption         =   "z2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label LABEL_LINE_y2 
         Caption         =   "y2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label LABEL_LINE_x2 
         Caption         =   "x2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label LABEL_LINE_z1 
         Caption         =   "z1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   255
      End
      Begin VB.Label LABEL_LINE_y1 
         Caption         =   "y1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   255
      End
      Begin VB.Label LABEL_LINE_x1 
         Caption         =   "x1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Screen Size"
      Height          =   975
      Left            =   120
      TabIndex        =   40
      Top             =   4560
      Width           =   2775
      Begin VB.ComboBox DROPDOWN_ySize 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox DROPDOWN_xSize 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LABEL_ySize 
         Caption         =   "y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   255
      End
      Begin VB.Label LABEL_xSize 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Set Type"
      Height          =   735
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox DROPDOWN_FractalSet 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FORM_Settings.frx":030A
         Left            =   120
         List            =   "FORM_Settings.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Colour Scheme"
      Height          =   615
      Left            =   120
      TabIndex        =   36
      Top             =   960
      Width           =   2775
      Begin VB.ComboBox DROPDOWN_ColourScheme 
         Height          =   315
         ItemData        =   "FORM_Settings.frx":039B
         Left            =   120
         List            =   "FORM_Settings.frx":03B4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Julia Set Specific"
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   5640
      Width           =   2775
      Begin VB.TextBox TEXTBOX_cy 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox TEXTBOX_cx 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LABEL_cy 
         Caption         =   "cy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   255
      End
      Begin VB.Label LABEL_cx 
         Caption         =   "cx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scale"
      Height          =   1695
      Left            =   120
      TabIndex        =   30
      Top             =   2760
      Width           =   2775
      Begin VB.TextBox TEXTBOX_y2 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   8
         Text            =   "Text6"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox TEXTBOX_y1 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox TEXTBOX_x2 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox TEXTBOX_x1 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LABEL_y2 
         Caption         =   "y2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label LABEL_y1 
         Caption         =   "y1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   255
      End
      Begin VB.Label LABEL_x2 
         Caption         =   "x2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   255
      End
      Begin VB.Label LABEL_x1 
         Caption         =   "x1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton BUTTON_Update 
      Caption         =   "Update Settings"
      Height          =   375
      Left            =   2280
      TabIndex        =   26
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Frame Variables 
      Caption         =   "Loop"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2775
      Begin VB.TextBox TEXTBOX_M 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox TEXTBOX_K 
         Height          =   285
         Left            =   360
         MaxLength       =   20
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label LABEL_M 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
      Begin VB.Label LABEL_K 
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Label LABEL_Validation 
      Alignment       =   2  'Center
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   6720
      Width           =   5775
   End
End
Attribute VB_Name = "FORM_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'000000000000000000000000000000000000000000000000000000000000000000000
'000000000000000000011000000000011000000000000000000000000000000000000
'000000011111100000101000110000111000011000000000001111110000000000000
'000000100001100001010000110001110000110000000000001000111100000000000
'000011000001100010100001110011010000110000000000001100001100000000000
'000011000011000111000000100010100001110000001110001100111000010000000
'000110000011001110100000100011000000100000110100001111000001111000000
'000110000000001101100111100110000101100000110001111100000010011011000
'000011000000111110111001111001111001100011111110011111001100010100000
'000011111111001100000001100001100001111110001000010001110000111000000
'000000011000000000000000000000000001111000000000010000011100111000000
'000000000000000000000000000000000000000000000000000000000001011000000
'000000000000000000000000000000000000000000000000000000000010110000000
'(c) 2002 by Riley T. Perry - Chillers of Entropy

'-> If the comments below look garbled then change font to COURIER NEW

'                                                 ,  ,
'                                                / \/ \
'                                              (/ //_ \_
'     .-._                                      \||  .  \
'      \  '-._                            _,:__.-"/---\_ \
' ______/___  '.    .--------------------'~-'--.)__( , )\ \
'`'--.___  _\  /    | Settings Form           ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 22/3/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * The settings form for the application.                 *
' *--------------------------------------------------------*

Option Explicit
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   BUTTON_Update_Click()                  |
' |  /   \  ---------------------                  |
' | |\_.  | Validate and assing new settings       |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub BUTTON_Update_Click()

    '*------------------------------------------*
    '*          Validate new settings           *
    '*------------------------------------------*
   
    Dim Validated As Boolean
    
    Validated = True
    LABEL_Validation = ""

    '***** Check if values are numeric and M & K >=1 ****

    If Not (IsNumeric(TEXTBOX_M.Text) _
            And IsNumeric(TEXTBOX_K.Text) _
            And IsNumeric(TEXTBOX_cx.Text) _
            And IsNumeric(TEXTBOX_cy.Text) _
            And IsNumeric(TEXTBOX_x1.Text) _
            And IsNumeric(TEXTBOX_x2.Text) _
            And IsNumeric(TEXTBOX_y1.Text) _
            And IsNumeric(TEXTBOX_y2.Text)) Then
            
        '**** Invalid! ****
            
        LABEL_Validation = "All textbox values must be numeric"
        Validated = False
        
    Else
    
        '**** All int or double values are numeric, check M & K ****
    
        If CInt(TEXTBOX_M.Text) < 1 Or CInt(TEXTBOX_K.Text) < 1 Or CInt(TEXTBOX_M.Text) > CONSTANT_MAX_M Then
        
            '**** Invalid! ****
           
            LABEL_Validation = "M and K values must be >=1 and M must be less than 300"
            Validated = False
        
        End If
        
    End If
    
    '**** Check List indexes ****
    
    If DROPDOWN_LINE_x1.ListIndex = -1 _
       Or DROPDOWN_LINE_y1.ListIndex = -1 _
       Or DROPDOWN_LINE_z1.ListIndex = -1 _
       Or DROPDOWN_LINE_x2.ListIndex = -1 _
       Or DROPDOWN_LINE_y2.ListIndex = -1 _
       Or DROPDOWN_LINE_z2.ListIndex = -1 _
       Or DROPDOWN_r.ListIndex = -1 _
       Or DROPDOWN_xc.ListIndex = -1 _
       Or DROPDOWN_yc.ListIndex = -1 _
       Or DROPDOWN_ZIndex.ListIndex = -1 Then

        '**** Invalid! ****
           
        LABEL_Validation = "Line index fault - select new values"
        Validated = False
        
    End If
     
    '*------------------------------------------*
    '*       If valid assign new settings       *
    '*------------------------------------------*
   
    If Validated Then
    
        '**** Int or double settings ****
    
        Settings.M = CInt(TEXTBOX_M.Text)
        Settings.K = CInt(TEXTBOX_K.Text)
        Settings.cx = CDbl(TEXTBOX_cx.Text)
        Settings.cy = CDbl(TEXTBOX_cy.Text)
        Settings.x1 = CDbl(TEXTBOX_x1.Text)
        Settings.x2 = CDbl(TEXTBOX_x2.Text)
        Settings.y1 = CDbl(TEXTBOX_y1.Text)
        Settings.y2 = CDbl(TEXTBOX_y2.Text)
        
        '**** Dropdown settings ****
        
        Settings.ColourScheme = DROPDOWN_ColourScheme.ListIndex
        Settings.FractalSet = DROPDOWN_FractalSet.ListIndex
        
        Settings.xSize = DROPDOWN_xSize.ListIndex + 1
        Settings.ySize = DROPDOWN_ySize.ListIndex + 1
        
        Settings.LINE_x1 = DROPDOWN_LINE_x1.ListIndex
        Settings.LINE_y1 = DROPDOWN_LINE_y1.ListIndex
        Settings.LINE_z1 = DROPDOWN_LINE_z1.ListIndex
        Settings.LINE_x2 = DROPDOWN_LINE_x2.ListIndex
        Settings.LINE_y2 = DROPDOWN_LINE_y2.ListIndex
        Settings.LINE_z2 = DROPDOWN_LINE_z2.ListIndex
        
        Settings.ZIndex = DROPDOWN_ZIndex.ListIndex
        Settings.xc = DROPDOWN_xc.ListIndex
        Settings.yc = DROPDOWN_yc.ListIndex
        Settings.r = DROPDOWN_r.ListIndex

        Settings.RGB_R = DROPDOWN_RGB_R.ListIndex
        Settings.RGB_G = DROPDOWN_RGB_G.ListIndex
        Settings.RGB_B = DROPDOWN_RGB_B.ListIndex
        
        '**** Close settings form ****
        
        FORM_Settings.Visible = False
        
    End If
 
End Sub


'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Form_Load()                            |
' |  /   \  -----------                            |
' | |\_.  | Assign old settings                    |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub Form_Load()

    Dim i As Integer

    '*------------------------------------------*
    '*            Assign old settings           *
    '*------------------------------------------*
   
    '**** Load settings values into text boxes ****

    TEXTBOX_M.Text = CStr(Settings.M)
    TEXTBOX_K.Text = CStr(Settings.K)
    TEXTBOX_cx.Text = CStr(Settings.cx)
    TEXTBOX_cy.Text = CStr(Settings.cy)
    TEXTBOX_x1.Text = CStr(Settings.x1)
    TEXTBOX_x2.Text = CStr(Settings.x2)
    TEXTBOX_y1.Text = CStr(Settings.y1)
    TEXTBOX_y2.Text = CStr(Settings.y2)
    
    '**** Load settings into text dropdown boxes ****
    
    DROPDOWN_ColourScheme.ListIndex = Settings.ColourScheme
    DROPDOWN_FractalSet.ListIndex = Settings.FractalSet

    '**** Load numeric values into x and y size dropdown boxes ****

    For i = 1 To 500
    
        DROPDOWN_xSize.List(i - 1) = CStr(i)
        DROPDOWN_ySize.List(i - 1) = CStr(i)
        
    Next
   
    '**** Default values for x and y size ***
    
    DROPDOWN_ySize.ListIndex = Settings.ySize - 1
    DROPDOWN_xSize.ListIndex = Settings.xSize - 1
    
    '**** Load numeric values into line, circle, and RGB modifier dropdown boxes ****

    For i = 0 To 255
    
        DROPDOWN_LINE_x1.List(i) = CStr(i)
        DROPDOWN_LINE_y1.List(i) = CStr(i)
        DROPDOWN_LINE_z1.List(i) = CStr(i)
        DROPDOWN_LINE_x2.List(i) = CStr(i)
        DROPDOWN_LINE_y2.List(i) = CStr(i)
        DROPDOWN_LINE_z2.List(i) = CStr(i)
        
        DROPDOWN_ZIndex.List(i) = CStr(i)
        DROPDOWN_xc.List(i) = CStr(i)
        DROPDOWN_yc.List(i) = CStr(i)
        
        DROPDOWN_RGB_R.List(i) = CStr(i)
        DROPDOWN_RGB_G.List(i) = CStr(i)
        DROPDOWN_RGB_B.List(i) = CStr(i)
        
    Next
   
    '**** Load numeric values into circle radius dropdown box ****

    For i = 0 To 126
    
        DROPDOWN_r.List(i) = CStr(i + 1)
        
    Next
  
    '**** Default values for the line ***
    
    DROPDOWN_LINE_x1.ListIndex = Settings.LINE_x1
    DROPDOWN_LINE_y1.ListIndex = Settings.LINE_y1
    DROPDOWN_LINE_z1.ListIndex = Settings.LINE_z1
    DROPDOWN_LINE_x2.ListIndex = Settings.LINE_x2
    DROPDOWN_LINE_y2.ListIndex = Settings.LINE_y2
    DROPDOWN_LINE_z2.ListIndex = Settings.LINE_z2
 
    '**** Default values for the circle ***
    
    DROPDOWN_ZIndex.ListIndex = Settings.ZIndex
    DROPDOWN_xc.ListIndex = Settings.xc
    DROPDOWN_yc.ListIndex = Settings.yc
    DROPDOWN_r.ListIndex = Settings.r
    
    '**** Default values for the RGB modifiers ****
    
    DROPDOWN_RGB_R.ListIndex = Settings.RGB_R
    DROPDOWN_RGB_G.ListIndex = Settings.RGB_G
    DROPDOWN_RGB_B.ListIndex = Settings.RGB_B

End Sub

