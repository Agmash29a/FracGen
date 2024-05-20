VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLASS_Mandelbrot"
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
'`'--.___  _\  /    | Mandelbrot              ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 27/3/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Class generates and draws Mandelbrot and Julia sets    *
' *--------------------------------------------------------*

Option Explicit

'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.                                          |
' |  /   \          Types and variables            |
' | |\_.  |         -------------------            |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       |
'  \     /
'   `---'
'
'*------------------------------------------*
'*               Public types               *
'*------------------------------------------*
        
'**** The colour scheme to be used ****

Public Enum ENUM_ColourScheme
    PurpleHaze = 0
    RhapsodyInBlue = 1
    JustTheSet = 2
    LineThroughColorSpace = 3
    CircleThroughColorSpace = 4
    Waves = 5
    RGBModifier = 6
End Enum

'*------------------------------------------*
'*            Private variables             *
'*------------------------------------------*
       
'**** boolean for even or odd iteration ****

Private IsEven As Boolean

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
Private mvarM As Integer 'local copy
Private mvarK As Integer 'local copy
Private mvarx1 As Double 'local copy
Private mvarx2 As Double 'local copy
Private mvary1 As Double 'local copy
Private mvary2 As Double 'local copy
Private mvarXWidth As Integer 'local copy
Private mvarYWidth As Integer 'local copy
Private mvarc As CLASS_ComplexNumber 'local copy
Private mvarLINE_x1 As Integer 'local copy
Private mvarLINE_x2 As Integer 'local copy
Private mvarLINE_y1 As Integer 'local copy
Private mvarLINE_y2 As Integer 'local copy
Private mvarLINE_z1 As Integer 'local copy
Private mvarLINE_z2 As Integer 'local copy
Private mvarr As Integer 'local copy
Private mvarxc As Integer 'local copy
Private mvaryc As Integer 'local copy
Private mvarZIndex As Integer 'local copy
Private mvarRGB_R As Integer 'local copy
Private mvarRGB_G As Integer 'local copy
Private mvarRGB_B As Integer 'local copy

Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte


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


Public Property Set c(ByVal vData As CLASS_ComplexNumber)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.c = Form1
    Set mvarc = vData
End Property


Public Property Get c() As CLASS_ComplexNumber
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.c
    Set c = mvarc
End Property


Public Property Let YWidth(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.YWidth = 5
    mvarYWidth = vData
End Property


Public Property Get YWidth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.YWidth
    YWidth = mvarYWidth
End Property


Public Property Let XWidth(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.XWidth = 5
    mvarXWidth = vData
End Property


Public Property Get XWidth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.XWidth
    XWidth = mvarXWidth
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
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   MandelbrotGenerator()                  |
' |  /   \  ---------------------                  |
' | |\_.  | Generates and draws a Mandelbrot set   |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Handle - handle of control to draw to
'   `---'   2.>> ColourScheme - Colours to use with set
'
Public Sub MandelbrotGenerator(ByRef Handle As OLE_HANDLE, ByVal ColourScheme As ENUM_ColourScheme)

    '*------------------------------------------*
    '*               Declarations               *
    '*------------------------------------------*
    
    '**** Instance of object to do complex arithmetic ****
  
    Dim ComplexArithmetic As New CLASS_ComplexArithmetic
 
    '**** Complex numbers ****
      
    Dim c As New CLASS_ComplexNumber
    Dim z As New CLASS_ComplexNumber

    '**** Difference between 2 adjacent points in complex plane vars. ****
      
    Dim xDifferenceBetweenAdjacent As Double
    Dim yDifferenceBetweenAdjacent As Double
    
    '**** Points on screen ****
    
    Dim x As Integer
    Dim y As Integer
          
    '**** Number of iterations per complex number ****
          
    Dim count As Integer

    '*------------------------------------------*
    '* Difference between 2 adjacent points in  *
    '* complex plane                            *
    '*------------------------------------------*
        
    xDifferenceBetweenAdjacent = StepValue(x1, x2, XWidth)
    yDifferenceBetweenAdjacent = StepValue(y1, y2, YWidth)
    
    '*------------------------------------------*
    '*               Generate Set               *
    '*------------------------------------------*
        
    'for (each pixel)
    '{
    
    For y = 0 To YWidth - 1
    
        For x = 0 To XWidth - 1
        
            'calculate c
            
            Set c = MapPixeltoComplexNumber(x, y, xDifferenceBetweenAdjacent, yDifferenceBetweenAdjacent)
            
            'count <- 0
            
            count = 0
            
            'z <- 0 + 0i
            
            z.x = 0
            z.y = 0
            
            'do
            '{
            
            While ((ComplexArithmetic.Magnitude(z) ^ 2 < M) And (count < K))
            
                'z <- z^2 + c
            
                Set z = ComplexArithmetic.Add(ComplexArithmetic.Multiply(z, z), c)
                
                'count <- count + 1
                
                count = count + 1
              
            Wend
          
            '} while (|z|^2 < M and count < K);
            
            'colour pixel according to the value of count
            
            Call DrawColouredPixel(count, x, y, Handle, ColourScheme)
            
        Next
    
    Next
    
    '}
    
    '*------------------------------------------*
    '*                Clean up                  *
    '*------------------------------------------*
    
    Set ComplexArithmetic = Nothing
    Set z = Nothing
    Set c = Nothing
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   JuliaGenerator_zSqrPlusc()             |
' |  /   \  --------------------------             |
' | |\_.  | Generates and draws a Julia set with   |
' |\|  | /| z -> z^2+c                            /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Handle - handle of control to draw to
'   `---'   2.>> ColourScheme - Colours to use with set
'
Public Sub JuliaGenerator_zSqrPlusc(ByRef Handle As OLE_HANDLE, ByVal ColourScheme As ENUM_ColourScheme)

    '*------------------------------------------*
    '*               Declarations               *
    '*------------------------------------------*
    
    '**** Instance of object to do complex arithmetic ****
  
    Dim ComplexArithmetic As New CLASS_ComplexArithmetic
 
    '**** Complex numbers ****
      
    Dim z As New CLASS_ComplexNumber

    '**** Difference between 2 adjacent points in complex plane vars. ****
      
    Dim xDifferenceBetweenAdjacent As Double
    Dim yDifferenceBetweenAdjacent As Double
    
    '**** Points on screen ****
    
    Dim x As Integer
    Dim y As Integer
          
    '**** Number of iterations per complex number ****
          
    Dim count As Integer

    '*------------------------------------------*
    '* Difference between 2 adjacent points in  *
    '* complex plane                            *
    '*------------------------------------------*
        
    xDifferenceBetweenAdjacent = StepValue(x1, x2, XWidth)
    yDifferenceBetweenAdjacent = StepValue(y1, y2, YWidth)
    
    '*------------------------------------------*
    '*               Generate Set               *
    '*------------------------------------------*
        
    'for (each pixel)
    '{
    
    For y = 0 To YWidth - 1
    
        For x = 0 To XWidth - 1
        
            'calculate z0
        
            Set z = MapPixeltoComplexNumber(x, y, xDifferenceBetweenAdjacent, yDifferenceBetweenAdjacent)
            
            'count <- 0
            
            count = 0
            
            'do
            '{
            
            While ((ComplexArithmetic.Magnitude(z) ^ 2 < M) And (count < K))
            
                'z <- z^2 + c
            
                Set z = ComplexArithmetic.Add(ComplexArithmetic.Multiply(z, z), c)
                
                'count <- count + 1
                
                count = count + 1
              
            Wend
          
            '} while (|z|^2 < M and count < K);
            
            'colour pixel according to the value of count
            
            Call DrawColouredPixel(count, x, y, Handle, ColourScheme)
            
        Next
    
    Next
 
    '}
    
    '*------------------------------------------*
    '*                Clean up                  *
    '*------------------------------------------*
    
    Set ComplexArithmetic = Nothing
    Set z = Nothing
            
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   JuliaGenerator_eToz()                  |
' |  /   \  ---------------------                  |
' | |\_.  | Generates and draws a Julia set with   |
' |\|  | /| z -> e^z+c                            /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Handle - handle of control to draw to
'   `---'   2.>> ColourScheme - Colours to use with set
'
Public Sub JuliaGenerator_eToz(ByRef Handle As OLE_HANDLE, ByVal ColourScheme As ENUM_ColourScheme)

    '*------------------------------------------*
    '*               Declarations               *
    '*------------------------------------------*
    
    '**** Instance of object to do complex arithmetic ****
  
    Dim ComplexArithmetic As New CLASS_ComplexArithmetic
 
    '**** Complex numbers ****
      
    Dim z As New CLASS_ComplexNumber

    '**** Difference between 2 adjacent points in complex plane vars. ****
      
    Dim xDifferenceBetweenAdjacent As Double
    Dim yDifferenceBetweenAdjacent As Double
    
    '**** Points on screen ****
    
    Dim x As Integer
    Dim y As Integer
          
    '**** Number of iterations per complex number ****
          
    Dim count As Integer

    '*------------------------------------------*
    '* Difference between 2 adjacent points in  *
    '* complex plane                            *
    '*------------------------------------------*
        
    xDifferenceBetweenAdjacent = StepValue(x1, x2, XWidth)
    yDifferenceBetweenAdjacent = StepValue(y1, y2, YWidth)
    
    '*------------------------------------------*
    '*               Generate Set               *
    '*------------------------------------------*
        
    'for (each pixel)
    '{
    
   For y = 0 To YWidth - 1
    
        For x = 0 To XWidth - 1
        
            'calculate z0
        
            Set z = MapPixeltoComplexNumber(x, y, xDifferenceBetweenAdjacent, yDifferenceBetweenAdjacent)
            
            'count <- 0
            
            count = 0
            
            'do
            '{
            
            '--> Another exit condition: While ((exp(ComplexArithmetic.Magnitude(z)) < M) And (count < K))
           
            While ((ComplexArithmetic.Magnitude(z) ^ 2 < M) And (count < K))
            
                'z <- e^z + c
            
                Set z = ComplexArithmetic.Add(ComplexArithmetic.eToz(z), c)
               
                'count <- count + 1
                
                count = count + 1
              
            Wend
          
            '} while (|z|^2 < M and count < K);
            
            'colour pixel according to the value of count
             
            Call DrawColouredPixel(count, x, y, Handle, ColourScheme)
           
        Next
    
    Next

    '}
     
    '*------------------------------------------*
    '*                Clean up                  *
    '*------------------------------------------*
    
    Set ComplexArithmetic = Nothing
    Set z = Nothing
          
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   JuliaGenerator_SinzPluseToz()          |
' |  /   \  -----------------------------          |
' | |\_.  | Generates and draws a Julia set with   |
' |\|  | /| z -> Sin(z) + e^z+c                   /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Handle - handle of control to draw to
'   `---'   2.>> ColourScheme - Colours to use with set
'
Public Sub JuliaGenerator_SinzPluseToz(ByRef Handle As OLE_HANDLE, ByVal ColourScheme As ENUM_ColourScheme)

    '*------------------------------------------*
    '*               Declarations               *
    '*------------------------------------------*
    
    '**** Instance of object to do complex arithmetic ****
  
    Dim ComplexArithmetic As New CLASS_ComplexArithmetic
 
    '**** Complex numbers ****
      
    Dim z As New CLASS_ComplexNumber

    '**** Difference between 2 adjacent points in complex plane vars. ****
      
    Dim xDifferenceBetweenAdjacent As Double
    Dim yDifferenceBetweenAdjacent As Double
    
    '**** Points on screen ****
    
    Dim x As Integer
    Dim y As Integer
          
    '**** Number of iterations per complex number ****
          
    Dim count As Integer

    '*------------------------------------------*
    '* Difference between 2 adjacent points in  *
    '* complex plane                            *
    '*------------------------------------------*
        
    xDifferenceBetweenAdjacent = StepValue(x1, x2, XWidth)
    yDifferenceBetweenAdjacent = StepValue(y1, y2, YWidth)
    
    '*------------------------------------------*
    '*               Generate Set               *
    '*------------------------------------------*
        
    'for (each pixel)
    '{
    
   For y = 0 To YWidth - 1
    
        For x = 0 To XWidth - 1
        
            'calculate z0
        
            Set z = MapPixeltoComplexNumber(x, y, xDifferenceBetweenAdjacent, yDifferenceBetweenAdjacent)
            
            'count <- 0
            
            count = 0
            
            'do
            '{
            
            While ((ComplexArithmetic.Magnitude(z) ^ 2 < M) And (count < K))
            
                'z <- Sin(z) + e^z + c
            
                Set z = ComplexArithmetic.Add(ComplexArithmetic.eToz(z), ComplexArithmetic.SinZ(z))
                Set z = ComplexArithmetic.Add(z, c)
                
                'count <- count + 1
                
                count = count + 1
              
            Wend
          
            '} while (|z|^2 < M and count < K);
            
            'colour pixel according to the value of count
             
            Call DrawColouredPixel(count, x, y, Handle, ColourScheme)
           
        Next
    
    Next

    '}
     
    '*------------------------------------------*
    '*                Clean up                  *
    '*------------------------------------------*
    
    Set ComplexArithmetic = Nothing
    Set z = Nothing
          
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   JuliaGenerator_ceToz()                 |
' |  /   \  ----------------------                 |
' | |\_.  | Generates and draws a Julia set with   |
' |\|  | /| z -> ce^z                             /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Handle - handle of control to draw to
'   `---'   2.>> ColourScheme - Colours to use with set
'
Public Sub JuliaGenerator_ceToz(ByRef Handle As OLE_HANDLE, ByVal ColourScheme As ENUM_ColourScheme)

    '*------------------------------------------*
    '*               Declarations               *
    '*------------------------------------------*
    
    '**** Instance of object to do complex arithmetic ****
  
    Dim ComplexArithmetic As New CLASS_ComplexArithmetic
 
    '**** Complex numbers ****
      
    Dim z As New CLASS_ComplexNumber

    '**** Difference between 2 adjacent points in complex plane vars. ****
      
    Dim xDifferenceBetweenAdjacent As Double
    Dim yDifferenceBetweenAdjacent As Double
    
    '**** Points on screen ****
    
    Dim x As Integer
    Dim y As Integer
          
    '**** Number of iterations per complex number ****
          
    Dim count As Integer

    '*------------------------------------------*
    '* Difference between 2 adjacent points in  *
    '* complex plane                            *
    '*------------------------------------------*
        
    xDifferenceBetweenAdjacent = StepValue(x1, x2, XWidth)
    yDifferenceBetweenAdjacent = StepValue(y1, y2, YWidth)
    
    '*------------------------------------------*
    '*               Generate Set               *
    '*------------------------------------------*
        
    'for (each pixel)
    '{
    
   For y = 0 To YWidth - 1
    
        For x = 0 To XWidth - 1
        
            'calculate z0
        
            Set z = MapPixeltoComplexNumber(x, y, xDifferenceBetweenAdjacent, yDifferenceBetweenAdjacent)
            
            'count <- 0
            
            count = 0
            
            'do
            '{
            
            While ((ComplexArithmetic.Magnitude(z) ^ 2 < M) And (count < K))
            
                'z <- ce^z
            
                Set z = ComplexArithmetic.Multiply(c, ComplexArithmetic.eToz(z))
                
                'count <- count + 1
                
                count = count + 1
              
            Wend
          
            '} while (|z|^2 < M and count < K);
            
            'colour pixel according to the value of count
             
            Call DrawColouredPixel(count, x, y, Handle, ColourScheme)
           
        Next
    
    Next

    '}
     
    '*------------------------------------------*
    '*                Clean up                  *
    '*------------------------------------------*
    
    Set ComplexArithmetic = Nothing
    Set z = Nothing
          
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   DrawColouredPixel()                    |
' |  /   \  -------------------                    |
' | |\_.  | Draw coloured pixel at (x,y)           |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> count - Number of iterations
'   `---'   2.>> x - Pixel index in x direction
'           3.>> y - Pixel index in y direction
'           4.>> Handle - handle of control to draw to
'           5.>> ColourScheme - Colours to use with set
'
Public Sub DrawColouredPixel(ByVal count As Integer, ByVal x As Integer, ByVal y As Integer, ByRef Handle As OLE_HANDLE, ByVal ColourScheme As ENUM_ColourScheme)
  
    '**** return value of pixel set function ****
    
    Dim isSet As Byte
    
    '**** Line scalar ****
    
    Dim u As Double
    
    '**** Line values ****
    
    Dim LINE_x As Double
    Dim LINE_y As Double
    Dim LINE_z As Double
    
    Dim CIRCLE_x As Double
    Dim CIRCLE_y As Double
    
    '**** theta value for trig functions ****
                
    Dim theta As Double
                    
    '*------------------------------------------*
    '* Colour pixel black if max iterations     *
    '* reached, otherwise use colour scheme     *
    '*------------------------------------------*
      
    If count = K Then
    
        '**** Max iterations reached, colour black ****

        isSet = SetPixelV(Handle, x, y, 0)

    Else
    
        '*------------------------------------------*
        '*          Choose colour scheme            *
        '*------------------------------------------*
     
        Select Case ColourScheme
        
            Case PurpleHaze
            
                '**** "Deep" purple colour scheme ****
            
                isSet = SetPixelV(Handle, x, y, RGB(Round((count Mod 100) * 1.5), Round((count Mod 100) * 0.5), Round((count Mod 100) * 2.5)))
                
            Case RhapsodyInBlue
        
                '**** "Deep" blue colour scheme ****
        
                isSet = SetPixelV(Handle, x, y, RGB(0, 0, Round((count Mod 100) * 2.5)))
                
            Case JustTheSet
            
                '**** White - do nothing ****
                
            Case LineThroughColorSpace
                
                '*------------------------------------------*
                '*    Arbitrary line though colour space    *
                '*                                          *
                '* x = x1 + (x2 - x1)u                      *
                '* y = y1 + (y2 - y1)u                      *
                '* z = z1 + (z2 - z1)u                      *
                '*                                          *
                '* 0 <= u <= 1                              *
                '* 0 <= (x,y,z) <= 1                        *
                '*------------------------------------------*
        
                '**** Determine u mapping to count ****
             
                u = count / (K / 255) / 255
     
                '**** determine x, y, and z vaues ****
                
                LINE_x = LINE_x1 / 255 + (((LINE_x2 / 255) - (LINE_x1 / 255)) * u)
                LINE_y = LINE_y1 / 255 + (((LINE_y2 / 255) - (LINE_y1 / 255)) * u)
                LINE_z = LINE_z1 / 255 + (((LINE_z2 / 255) - (LINE_z1 / 255)) * u)
     
                '**** Draw ****
     
                isSet = SetPixelV(Handle, x, y, RGB(Round(LINE_x * 255), Round(LINE_y * 255), Round(LINE_z * 255)))
                
            Case CircleThroughColorSpace
                       
                '*------------------------------------------*
                '*             Arbitrary circle             *
                '*                                          *
                '* x = xc + r * cos (theta)                 *
                '* y = yc + r * sin (theta)                 *
                '*------------------------------------------*
                
                '**** determine x and y values, (1 / K) is stepsize (after converting to rads) ****
            
                theta = ((count / K) * 360) / 180 * 3.141592654
                    
                CIRCLE_x = Abs(xc + (r * Cos(theta)))
                CIRCLE_y = Abs(yc + (r * Sin(theta)))
                
                '**** Draw ****
                
                isSet = SetPixelV(Handle, x, y, RGB(Round(CIRCLE_x), Round(CIRCLE_y), ZIndex))
                      Case Waves
                
                '*------------------------------------------*
                '* alternate between circle and line        *
                '* depending on value of count and K        *
                '*------------------------------------------*
     
                If (count Mod 2) = 1 And count < (K / 2) Then
                                
                    '*------------------------------------------*
                    '*             Arbitrary circle             *
                    '*                                          *
                    '* x = xc + r * cos (theta)                 *
                    '* y = yc + r * sin (theta)                 *
                    '*------------------------------------------*
                
                    '**** determine x and y values, (1 / K) is stepsize (after converting to rads) ****

                    theta = ((count / K) * 360) / 180 * 3.141592654
                    
                    CIRCLE_x = Abs(xc + (r * Cos(theta)))
                    CIRCLE_y = Abs(yc + (r * Sin(theta)))
                
                    '**** Draw ****
                
                    isSet = SetPixelV(Handle, x, y, RGB(Round(CIRCLE_x), Round(CIRCLE_y), ZIndex))
            
                
                Else
               
                    '*------------------------------------------*
                    '*    Arbitrary line though colour space    *
                    '*               (with waves)               *
                    '*                                          *
                    '* x = x1 + (x2 - x1)u                      *
                    '* y = y1 + (y2 - y1)u                      *
                    '* z = z1 + (z2 - z1)u                      *
                    '*                                          *
                    '* 0 <= u <= 1                              *
                    '* 0 <= (x,y,z) <= 1                        *
                    '*------------------------------------------*
        
                    '**** Determine u mapping to count ****
             
                    u = count / (K / 255) / 255
                
                    '**** determine x, y, and z vaues ****
                   
                    LINE_x = LINE_x1 / 255 + (((LINE_x2 / 255) - (LINE_x1 / 255)) * u)
                    LINE_y = LINE_y1 / 255 + (((LINE_y2 / 255) - (LINE_y1 / 255)) * u)
                    LINE_z = LINE_z1 / 255 + (((LINE_z2 / 255) - (LINE_z1 / 255)) * u)
        
                    '**** Draw ****
        
                    isSet = SetPixelV(Handle, x, y, RGB(Round(LINE_x * 255), Round(LINE_y * 255), Round(LINE_z * 255)))
                
                End If
                
                IsEven = Not IsEven
                
            Case RGBModifier
                       
                '*------------------------------------------*
                '*   Modify count value by a multiplier     *
                '*------------------------------------------*
               
                '**** Draw ****
                
                isSet = SetPixelV(Handle, x, y, RGB(count * RGB_R Mod 256, count * RGB_G Mod 256, count * RGB_B Mod 256))
                
        End Select
    
    End If
    
End Sub
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   StepValue()                            |
' |  /   \  -----------                            |
' | |\_.  | Find difference between 2 adjacent     |
' |\|  | /| points in complex plane               /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> xy1 - Delta x (or delta y) min
'   `---'   2.>> xy2 - Delta x (or delta y) max
'           3.>> WidthLength - Size of view window in x or y direction
'
'           Returns:
'           1.<< Resulting step value
'
Private Function StepValue(ByVal xy1 As Double, ByVal xy2 As Double, ByVal WidthLength As Integer) As Double
    
    '*------------------------------------------*
    '*             Find step value              *
    '*------------------------------------------*
    
    StepValue = (xy2 - xy1) / WidthLength

End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   MapPixeltoComplexNumber()              |
' |  /   \  -------------------------              |
' | |\_.  | Find complex number that corresponds   |
' |\|  | /| to a pixel                            /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> x - Delta x (or delta y) min
'   `---'   2.>> y - Delta x (or delta y) max
'           3.>> StepValuex - Difference between 2 adjacent points on complex plane in x direction
'           3.>> StepValuey - Difference between 2 adjacent points on complex plane in y direction
'
'           Returns:
'           1.<< Resulting complex number
'
Private Function MapPixeltoComplexNumber(ByVal x As Integer, ByVal y As Integer, ByVal StepValuex As Double, ByVal StepValuey As Double) As CLASS_ComplexNumber

    '*------------------------------------------*
    '*       Match pixel to complex number      *
    '*------------------------------------------*
    
    Dim dummy As New CLASS_ComplexNumber
    
    dummy.x = x1 + (StepValuex * x)
    dummy.y = y2 - (StepValuey * y)

    '*------------------------------------------*
    '*                 Clean up                 *
    '*------------------------------------------*
        
    Set MapPixeltoComplexNumber = dummy
    Set dummy = Nothing
     
End Function


