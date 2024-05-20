Attribute VB_Name = "MODULE_Globals"
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
'`'--.___  _\  /    | Global Variables        ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 27/3/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Globals.                                               *
' *--------------------------------------------------------*

'**** Public classes ****

Public Settings As New CLASS_Settings

'**** Constants (default values) ****

Public Const CONSTANT_M As Integer = 200
Public Const CONSTANT_MAX_M As Integer = 1999

Public Const CONSTANT_K As Integer = 100

Public Const CONSTANT_cx As Double = -0.194
Public Const CONSTANT_cy As Double = 0.6557

Public Const CONSTANT_x1 As Double = -2
Public Const CONSTANT_x2 As Double = 2

Public Const CONSTANT_y1 As Double = -2
Public Const CONSTANT_y2 As Double = 2

Public Const CONSTANT_ColourScheme As Integer = 0
Public Const CONSTANT_FractalSet As Integer = 0

Public Const CONSTANT_xScreenLength As Integer = 500
Public Const CONSTANT_yScreenLength As Integer = 500

Public Const CONSTANT_LINE_x1 As Integer = 0
Public Const CONSTANT_LINE_y1 As Integer = 0
Public Const CONSTANT_LINE_z1 As Integer = 0

Public Const CONSTANT_LINE_x2 As Integer = 255
Public Const CONSTANT_LINE_y2 As Integer = 255
Public Const CONSTANT_LINE_z2 As Integer = 255

Public Const CONSTANT_yc As Integer = 127
Public Const CONSTANT_xc As Integer = 127

Public Const CONSTANT_ZIndex As Integer = 127

Public Const CONSTANT_r As Integer = 126 'radius - 1

Public Const CONSTANT_RGB_R As Integer = 1
Public Const CONSTANT_RGB_G As Integer = 1
Public Const CONSTANT_RGB_B As Integer = 1


'**** Public types ****

Public Enum ENUM_SetType
    MandelbrotSet = 0
    JuliaSet_zSqrPlusc = 1
    JuliaSet_eTozPlusc = 2
    JuliaSet_SinzPluseTozPlusc = 3
    JuliaSet_ceToz = 4
End Enum

