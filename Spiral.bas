Attribute VB_Name = "Module1"
Option Explicit
Option Base 0
Option Compare Binary
DefInt A-Z

' Constants for making a spiral
Global Const Clockwise = True
Global Const AntiClockwise = False
Global Const Line = True
Global Const Dot = False
Global Const Pi = 3.14159265358979

Function DrawSpiral(Radius As Single, Picture As PictureBox, Increment As Single, x As Single, y As Single, Color As Long, StartAngle As Single, Direction As Boolean, AngleStep As Single, DotLine As Boolean, Optional ButtonOff As Control) As Boolean

' Checking to see that variables passed through are all present and correct.
If Radius = 0 Then DrawSpiral = False: Exit Function
If Increment = 0 Then DrawSpiral = False: Exit Function

' Variables
Dim I As Single 'Looper
Dim XInc As Double
Dim YInc As Double 'Incremental values of X and Y
Dim Angle As Double 'The Angle calculations

' Make the control "unavailable" if nescessary
On Error Resume Next
ButtonOff.Enabled = False
On Error GoTo 0

' Put the original dot on the screen, so that I can line draw from it.
Picture.PSet (x, y), Color
' Initialise all of the variables
Angle = StartAngle

' Draw the Spiral =Þ
For I = 1 To Radius Step Increment

' Calculate the new angle to put the point at
If Direction = Clockwise Then
Angle = Angle - AngleStep
Else
Angle = Angle + AngleStep
End If
Angle = Angle

' Generate the XInc and YInc Values
XInc = (I * Cos(Angle * (Pi / 180)))
YInc = (I * Sin(Angle * (Pi / 180)))

' Put the dot/line on the screen
If DotLine = Line Then
Picture.Line -(x + XInc, y - YInc), Color
Else
Picture.PSet (x + XInc, y - YInc), Color
End If

DoEvents

Next I

' Make the control "available" if nescessary
On Error Resume Next
ButtonOff.Enabled = True
On Error GoTo 0

DrawSpiral = True

End Function
