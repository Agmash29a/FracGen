VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLASS_ComplexNumber"
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
'`'--.___  _\  /    | ComplexNumber             ,'    \)|\ `\|
'     /_.-' _\ \ _:,_                               " ||   (
'   .'__ _.' \'-/,`-~`                                |/
'       '. ___.> /=,| 27/3/2002 - Riley T. Perry      |
'        / .-'/_ )  '---------------------------------'
'        )'  ( /(/             Riley@deliverance.com.au
'             \\ "
'              '=='
'
' *--------------------------------------------------------*
' * Class encapsulating a complex number.                  *
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
Private mvarx As Double 'local copy
Private mvary As Double 'local copy

Public Property Let y(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.y = 5
    mvary = vData
End Property


Public Property Get y() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.y
    y = mvary
End Property


Public Property Let x(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.x = 5
    mvarx = vData
End Property


Public Property Get x() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.x
    x = mvarx
End Property



