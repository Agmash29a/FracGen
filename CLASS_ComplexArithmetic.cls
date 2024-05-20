VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLASS_ComplexArithmetic"
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
'`'--.___  _\  /    | ComplexArithmetic       ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 27/3/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Varoius functions on complex numbers.                  *
' *--------------------------------------------------------*

Option Explicit

'**** Constants ****

Private Const CONSTANT_e As Double = 2.718281828

'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Multiply()                             |
' |  /   \  ----------                             |
' | |\_.  | Multiplies 2 complex numbers           |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> SetType - The type of set to generate
'   `---'   2.>> z1 - Multiplicand
'           3.>> z2 - Multiplier
'
'           Returns:
'           1.<< Result (z1 * z2)
'
Public Function Multiply(ByVal z1 As CLASS_ComplexNumber, ByVal z2 As CLASS_ComplexNumber) As CLASS_ComplexNumber

    '*------------------------------------------*
    '*          Multiply the 2 numbers          *
    '*------------------------------------------*
    
    Dim dummy As New CLASS_ComplexNumber
    
    dummy.x = (z1.x * z2.x) - (z1.y * z2.y)
    dummy.y = (z1.x * z2.y) + (z2.x * z1.y)

    '*------------------------------------------*
    '*                 Clean up                 *
    '*------------------------------------------*
    
    Set Multiply = dummy
    Set dummy = Nothing
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Magnitude()                            |
' |  /   \  -----------                            |
' | |\_.  | Calculate magnitude of a complex no.   |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> z - Complex number as argument
'   `---'
'           Returns:
'           1.<< Magnitude of argument
'
Public Function Magnitude(ByVal z As CLASS_ComplexNumber) As Double

    '*------------------------------------------*
    '*          Calculate the Magnitude         *
    '*------------------------------------------*
    
    Magnitude = Sqr((z.x * z.x) + (z.y * z.y))

End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   Add()                                  |
' |  /   \  -----                                  |
' | |\_.  | Adds 2 complex numbers                 |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> z1 - Addend
'   `---'   2.>> z2 - Augend
'
'           Returns:
'           1.<< Sum of z1 and z2
'
Public Function Add(ByVal z1 As CLASS_ComplexNumber, ByVal z2 As CLASS_ComplexNumber) As CLASS_ComplexNumber

    '*------------------------------------------*
    '*             Add the 2 numbers            *
    '*------------------------------------------*
    
    Dim dummy As New CLASS_ComplexNumber
    
    dummy.x = z1.x + z2.x
    dummy.y = z1.y + z2.y

    '*------------------------------------------*
    '*                 Clean up                 *
    '*------------------------------------------*
    
    Set Add = dummy
    Set dummy = Nothing
    
End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   eToz()                                 |
' |  /   \  ------                                 |
' | |\_.  | Calculate e^z                          |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> z - Complex number as exponent
'   `---'
'           Returns:
'           1.<< Result of e^z
'
Public Function eToz(ByVal z As CLASS_ComplexNumber) As CLASS_ComplexNumber

    '*------------------------------------------*
    '*              Calculate e^z               *
    '*------------------------------------------*
    
    Dim dummy As New CLASS_ComplexNumber
    
    '**** e^(a + bi) = e^a * cos(b) + e^a * sin(b) ****
    
    dummy.x = (Exp(z.x)) * (Cos(z.y))
    dummy.y = (Exp(z.x)) * (Sin(z.y))

    '*------------------------------------------*
    '*                 Clean up                 *
    '*------------------------------------------*
    
    Set eToz = dummy
    Set dummy = Nothing

End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   DegToRad()                             |
' |  /   \  ----------                             |
' | |\_.  | Convert Degrees to Rads                |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> Degrees - number of degrees
'   `---'
'           Returns:
'           1.<< Degrees as Rads
'
Public Function DegToRad(ByVal Degrees As Double) As Double

    '*------------------------------------------*
    '*           Convert to Radians             *
    '*------------------------------------------*
    
    DegToRad = Degrees / 180 * 3.141592654

End Function
'                                            .---.
'                                           /  .  \
'                                          |\_/|   |
'                                          |   |  /|
'   .--------------------------------------------' |
'  /  .-.   SinZ()                                 |
' |  /   \  ------                                 |
' | |\_.  | Calculate Sin(z)                       |
' |\|  | /|                                       /
' | `---' |--------------------------------------'
' \       | Parameters:
'  \     /  1.>> z - Complex number as parameter
'   `---'
'           Returns:
'           1.<< Result of Sin(z)
'
Public Function SinZ(ByVal z As CLASS_ComplexNumber) As CLASS_ComplexNumber

    '*------------------------------------------*
    '*             Calculate Sin(z)             *
    '*------------------------------------------*
    
    Dim dummy As New CLASS_ComplexNumber
    
    '**** sin(a + bi) = sin(x)cosh(y) + i*cos(x)sinh(y) ****
    
    dummy.x = (Sin(z.x) * ((Exp(z.y) + Exp(-z.y)) / 2))
    dummy.y = (Cos(z.x) * ((Exp(z.y) - Exp(-z.y)) / 2))

    '*------------------------------------------*
    '*                 Clean up                 *
    '*------------------------------------------*
    
    Set SinZ = dummy
    Set dummy = Nothing

End Function
