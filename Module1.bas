Attribute VB_Name = "Module1"
Option Explicit

Function toFarenheit(degrees)

toFarenheit = (9 / 5) * degrees + 32

End Function


Function toCentigrade(degrees)

toCentigrade = (5 / 9) * (degrees - 32)

End Function

Sub test()

' this is just random text

Dim answer   'all variables need to be declared since option explicit is listed

answer = toCentigrade(55)
MsgBox answer

End Sub
