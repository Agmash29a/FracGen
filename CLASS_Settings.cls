VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLASS_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
'`'--.___  _\  /    | Settings Class          ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 27/3/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Settings for the program.                              *
' *--------------------------------------------------------*

Option Explicit

'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.                                          |
' |  /   \          Get and set methods            |
' | |\_.  |         -------------------            |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
'local variable(s) to hold property value(s)
Private mvarx1 As Double 'local copy
Private mvarx2 As Double 'local copy
Private mvary1 As Double 'local copy
Private mvary2 As Double 'local copy
Private mvarM As Integer 'local copy
Private mvarK As Integer 'local copy
Private mvarcx As Double 'local copy
Private mvarcy As Double 'local copy
Private mvarColourScheme As Integer 'local copy
Private mvarFractalSet As Integer 'local copy
Private mvarxSize As Integer 'local copy
Private mvarySize As Integer 'local copy
Private mvarLINE_x1 As Integer 'local copy
Private mvarLINE_x2 As Integer 'local copy
Private mvarLINE_y1 As Integer 'local copy
Private mvarLINE_y2 As Integer 'local copy
Private mvarLINE_z1 As Integer 'local copy
Private mvarLINE_z2 As Integer 'local copy
Private mvarZIndex As Integer 'local copy
Private mvarxc As Integer 'local copy
Private mvaryc As Integer 'local copy
Private mvarr As Integer 'local copy
Private mvarRGB_R As Integer 'local copy
Private mvarRGB_G As Integer 'local copy
Private mvarRGB_B As Integer 'local copy


Public Property Let RGB_B(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RGB_B = 5
    mvarRGB_B = vData
End Property


Public Property Get RGB_B() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RGB_B
    RGB_B = mvarRGB_B
End Property


Public Property Let RGB_G(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RGB_G = 5
    mvarRGB_G = vData
End Property


Public Property Get RGB_G() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RGB_G
    RGB_G = mvarRGB_G
End Property


Public Property Let RGB_R(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RGB_R = 5
    mvarRGB_R = vData
End Property


Public Property Get RGB_R() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RGB_R
    RGB_R = mvarRGB_R
End Property


Public Property Let r(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.r = 5
    mvarr = vData
End Property


Public Property Get r() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.r
    r = mvarr
End Property


Public Property Let yc(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.yc = 5
    mvaryc = vData
End Property


Public Property Get yc() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.yc
    yc = mvaryc
End Property


Public Property Let xc(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xc = 5
    mvarxc = vData
End Property


Public Property Get xc() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xc
    xc = mvarxc
End Property


Public Property Let ZIndex(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ZIndex = 5
    mvarZIndex = vData
End Property


Public Property Get ZIndex() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ZIndex
    ZIndex = mvarZIndex
End Property


Public Property Let LINE_z2(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LINE_z2 = 5
    mvarLINE_z2 = vData
End Property


Public Property Get LINE_z2() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LINE_z2
    LINE_z2 = mvarLINE_z2
End Property


Public Property Let LINE_z1(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LINE_z1 = 5
    mvarLINE_z1 = vData
End Property


Public Property Get LINE_z1() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LINE_z1
    LINE_z1 = mvarLINE_z1
End Property


Public Property Let LINE_y2(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LINE_y2 = 5
    mvarLINE_y2 = vData
End Property


Public Property Get LINE_y2() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LINE_y2
    LINE_y2 = mvarLINE_y2
End Property


Public Property Let LINE_y1(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LINE_y1 = 5
    mvarLINE_y1 = vData
End Property


Public Property Get LINE_y1() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LINE_y1
    LINE_y1 = mvarLINE_y1
End Property



Public Property Let LINE_x2(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LINE_x2 = 5
    mvarLINE_x2 = vData
End Property


Public Property Get LINE_x2() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LINE_x2
    LINE_x2 = mvarLINE_x2
End Property


Public Property Let LINE_x1(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LINE_x1 = 5
    mvarLINE_x1 = vData
End Property


Public Property Get LINE_x1() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LINE_x1
    LINE_x1 = mvarLINE_x1
End Property


Public Property Let ySize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ySize = 5
    mvarySize = vData
End Property


Public Property Get ySize() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ySize
    ySize = mvarySize
End Property


Public Property Let xSize(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.xSize = 5
    mvarxSize = vData
End Property


Public Property Get xSize() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.xSize
    xSize = mvarxSize
End Property


Public Property Let FractalSet(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FractalSet = 5
    mvarFractalSet = vData
End Property


Public Property Get FractalSet() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FractalSet
    FractalSet = mvarFractalSet
End Property


Public Property Let ColourScheme(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ColourScheme = 5
    mvarColourScheme = vData
End Property


Public Property Get ColourScheme() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ColourScheme
    ColourScheme = mvarColourScheme
End Property


Public Property Let cy(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cy = 5
    mvarcy = vData
End Property


Public Property Get cy() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cy
    cy = mvarcy
End Property


Public Property Let cx(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.cx = 5
    mvarcx = vData
End Property


Public Property Get cx() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cx
    cx = mvarcx
End Property


Public Property Let K(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.K = 5
    mvarK = vData
End Property


Public Property Get K() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.K
    K = mvarK
End Property


Public Property Let M(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.M = 5
    mvarM = vData
End Property


Public Property Get M() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.M
    M = mvarM
End Property



Public Property Let y2(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.y2 = 5
    mvary2 = vData
End Property


Public Property Get y2() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.y2
    y2 = mvary2
End Property


Public Property Let y1(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.y1 = 5
    mvary1 = vData
End Property


Public Property Get y1() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.y1
    y1 = mvary1
End Property


Public Property Let x2(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.x2 = 5
    mvarx2 = vData
End Property


Public Property Get x2() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.x2
    x2 = mvarx2
End Property


Public Property Let x1(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.x1 = 5
    mvarx1 = vData
End Property


Public Property Get x1() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.x1
    x1 = mvarx1
End Property



