Attribute VB_Name = "Module2"
Option Explicit
Dim z As Double 'this is a global variable to this module
Private zz As Double    'this is also a global variable for this module, it is the same as dim
Public zzz As Double    'this is public variable that can be used by other modules


Sub math_example()
    Dim x As Double 'this is a local variable
    Dim y As Double
    Dim result As String
    
    x = 67
    y = 33
    result = "your answer is: " & x + y
    MsgBox result
    

End Sub
