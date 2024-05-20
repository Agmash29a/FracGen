VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FORM_MainScreen 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FracGen"
   ClientHeight    =   7560
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "FORM_MainScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FORM_MainScreen.frx":030A
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   504
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog DIALOG_Save 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox PICTURE_Set 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   7560
      Left            =   0
      Picture         =   "FORM_MainScreen.frx":1752A
      ScaleHeight     =   480.769
      ScaleMode       =   0  'User
      ScaleWidth      =   500
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   7560
   End
   Begin VB.PictureBox PICTURE_Title 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   7560
      Left            =   0
      Picture         =   "FORM_MainScreen.frx":17C52
      ScaleHeight     =   480.769
      ScaleMode       =   0  'User
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      Begin VB.Label LABEL_Generating 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Generating"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   65.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.Label LABEL_ByMe 
         BackStyle       =   0  'Transparent
         Caption         =   "FracGen - 2002 by Riley T. Perry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   75
         TabIndex        =   1
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuitemGenerate 
         Caption         =   "Generate"
      End
      Begin VB.Menu mnuitemSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuItemSaveImage 
         Caption         =   "Save Image"
      End
   End
End
Attribute VB_Name = "FORM_MainScreen"
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
'`'--.___  _\  /    | Main Form               ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 27/3/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * The main form for the application.                     *
' *--------------------------------------------------------*

'Option Explicit --> removed for file flags
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Form_Load()                            |
' |  /   \  -----------                            |
' | |\_.  | Assign default settings                |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub Form_Load()

    '*------------------------------------------*
    '*          Assign default settings         *
    '*------------------------------------------*
   
    Settings.ColourScheme = CONSTANT_ColourScheme
    Settings.FractalSet = CONSTANT_FractalSet
    Settings.M = CONSTANT_M
    Settings.K = CONSTANT_K
    Settings.cx = CONSTANT_cx
    Settings.cy = CONSTANT_cy
    Settings.x1 = CONSTANT_x1
    Settings.x2 = CONSTANT_x2
    Settings.y1 = CONSTANT_y1
    Settings.y2 = CONSTANT_y2
    Settings.xSize = CONSTANT_xScreenLength
    Settings.ySize = CONSTANT_yScreenLength
    
    Settings.LINE_x1 = CONSTANT_LINE_x1
    Settings.LINE_y1 = CONSTANT_LINE_y1
    Settings.LINE_z1 = CONSTANT_LINE_z1
    Settings.LINE_x2 = CONSTANT_LINE_x2
    Settings.LINE_y2 = CONSTANT_LINE_y2
    Settings.LINE_z2 = CONSTANT_LINE_z2
    
    Settings.xc = CONSTANT_xc
    Settings.yc = CONSTANT_yc
    Settings.r = CONSTANT_r
    Settings.ZIndex = CONSTANT_ZIndex

    Settings.RGB_R = CONSTANT_RGB_R
    Settings.RGB_G = CONSTANT_RGB_G
    Settings.RGB_B = CONSTANT_RGB_B
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   mnuitemGenerate_Click()                |
' |  /   \  -----------------------                |
' | |\_.  | Executed when Generate option is       |
' |\|  | /| selected                              /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub mnuitemGenerate_Click()
    
    '*------------------------------------------*
    '*  Generate set when menu option selected  *
    '*------------------------------------------*
   
    GenerateSet (Settings.FractalSet)
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   GenerateSet()                          |
' |  /   \  -------------                          |
' | |\_.  | Generates a madelbrot or Julia set     |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> SetType - The type of set to generate
'   `---'
'
Private Sub GenerateSet(ByVal SetType As ENUM_SetType)

    '*------------------------------------------*
    '*         Clear screen and refresh         *
    '*------------------------------------------*
   
    PICTURE_Set.Cls
    LABEL_Generating.Visible = True
    PICTURE_Set.Visible = False
    PICTURE_Title.Refresh
    

    '*------------------------------------------*
    '*          Set up Mandelbrot Object        *
    '*------------------------------------------*

    Dim MySet As New CLASS_Mandelbrot
    Dim c As New CLASS_ComplexNumber
            
    '**** Julia set value for variable "c" ****
    
    c.x = Settings.cx
    c.y = Settings.cy
    
    Set MySet.c = c
    
    '**** Other values from settings and Consts ****
    
    MySet.M = Settings.M
    MySet.K = Settings.K
    MySet.x1 = Settings.x1
    MySet.x2 = Settings.x2
    MySet.y1 = Settings.y1
    MySet.y2 = Settings.y2
    
    MySet.XWidth = Settings.xSize
    MySet.YWidth = Settings.ySize
    
    MySet.LINE_x1 = Settings.LINE_x1
    MySet.LINE_y1 = Settings.LINE_y1
    MySet.LINE_z1 = Settings.LINE_z1
    MySet.LINE_x2 = Settings.LINE_x2
    MySet.LINE_y2 = Settings.LINE_y2
    MySet.LINE_z2 = Settings.LINE_z2
    
    MySet.xc = Settings.xc
    MySet.yc = Settings.yc
    MySet.r = Settings.r + 1
    MySet.ZIndex = Settings.ZIndex
    
    MySet.RGB_R = Settings.RGB_R
    MySet.RGB_G = Settings.RGB_G
    MySet.RGB_B = Settings.RGB_B
    
    '*------------------------------------------*
    '*          Call Generator Function         *
    '*------------------------------------------*
           
    Select Case SetType
    
        Case MandelbrotSet
        
            '**** Mandelbrot set ****
        
            Call MySet.MandelbrotGenerator(FORM_MainScreen.PICTURE_Set.hdc, Settings.ColourScheme)
        
        Case JuliaSet_zSqrPlusc
       
            '**** Julia set (z->z^2+c) ****
       
            Call MySet.JuliaGenerator_zSqrPlusc(FORM_MainScreen.PICTURE_Set.hdc, Settings.ColourScheme)
               
        Case JuliaSet_eTozPlusc
    
            '**** Julia set (z->e^z+c) ****
    
            Call MySet.JuliaGenerator_eToz(FORM_MainScreen.PICTURE_Set.hdc, Settings.ColourScheme)
        
        Case JuliaSet_SinzPluseTozPlusc

            '**** Julia set (z->Sin(z)+e^z+c) ****
    
            Call MySet.JuliaGenerator_SinzPluseToz(FORM_MainScreen.PICTURE_Set.hdc, Settings.ColourScheme)
        
        Case JuliaSet_ceToz

            '**** Julia set (z->ce^z) ****
    
            Call MySet.JuliaGenerator_ceToz(FORM_MainScreen.PICTURE_Set.hdc, Settings.ColourScheme)
               
    End Select

    '*------------------------------------------*
    '*           Clean up and refresh           *
    '*------------------------------------------*
    
    PICTURE_Set.Refresh
    PICTURE_Set.Visible = True
    
    Set MySet = Nothing
    Set c = Nothing

End Sub


'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   mnuitemSettings_Click()                |
' |  /   \  -----------------------                |
' | |\_.  | Show settings form                     |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub mnuitemSettings_Click()

    '*------------------------------------------*
    '*            Show settings form            *
    '*------------------------------------------*
    
    FORM_Settings.Visible = True

End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Form_Unload()                          |
' |  /   \  -------------                          |
' | |\_.  | Close and clean up                     |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub Form_Unload(Cancel As Integer)
    
    '*------------------------------------------*
    '*                  Close                   *
    '*------------------------------------------*
 
    Unload FORM_Settings
    
    End

End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   mnuItemSaveImage_Click()               |
' |  /   \  ------------------------               |
' | |\_.  | Saves the fractal as a bitmap          |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
Private Sub mnuItemSaveImage_Click()
   
    '*------------------------------------------*
    '*               Save Image                 *
    '*------------------------------------------*
 
    '**** create and set cancelled bool ****
 
    Dim cancelled As Boolean
 
    cancelled = True
 
    '**** Trap File Error ****
    
    On Error GoTo Error_Handler
    
    '**** Dialog box options ****
    
    DIALOG_Save.DefaultExt = "bmp"
    DIALOG_Save.Filter = "Bitmap files|*.bmp"
    DIALOG_Save.FilterIndex = 1
    DIALOG_Save.Flags = cdlOHideReadOnly Or cdlOFNPathMustExist Or _
        cdlOFNOverwritePrompt Or cdlOFNNoReadOnlyReturn
    DIALOG_Save.DialogTitle = "Select the image file"
    DIALOG_Save.CancelError = True
    
    '**** Show save dialog box ****
    
    DIALOG_Save.ShowSave
    
    '**** Box was not cancelled ****
    
    cancelled = False
    
    '**** Save file ****
    
    SavePicture PICTURE_Set.Image, DIALOG_Save.FileName
 
    Exit Sub
    
Error_Handler:
    
    '**** show error message box if not cancelled ****
    
    If Not cancelled Then
    
        Dim result As VbMsgBoxResult
        result = MsgBox("Invalid File Operation", , "File Error")
        
    End If
    
End Sub

