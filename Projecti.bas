Attribute VB_Name = "Module1"
Global tgwidth As Integer
Global g As Single
Global visc As Single
Global wind As Single
Global target As Integer
Global count As Integer
Global dir As Integer



'A program to demonstrate projectile motion
'HTC Enterprises 1996

Function shoot()
ReDim t(3)
ReDim sx(3)
ReDim sy(3)
ReDim winda(3)
'ReDim visca(3)

step = 0.05
count = count + 1
Form1.Label9.Caption = "Tries = " & count
v = CInt(Form1.Text2.Text)
If Form1.Text1.Text = "" Then ang = 0 Else ang = CInt(Form1.Text1.Text)
ang = ang * (3.141592653592 / 180)
Y = v * Sin(ang)
X = v * Cos(ang)
t1 = -2 * step

While sy(2) < 3000 And t1 < 1000
t1 = t1 + step
For b = 0 To 3
t(b) = t1 + b * step
winda(b) = 0.5 * wind * t(b) ^ 2
'visca(b) = 0.5 * visc * t(b) ^ 2
sx(b) = X * t(b) - winda(b) ' - visca(b)
sy(b) = Y * t(b) - 0.5 * g * t(b) ^ 2 ' - visca(b)
Next b
'v = (sx(3) - sx(2)) / step
'visca(2) = 0.001 * visc * v * t(2) ^ 2
sy(2) = (3000 - sy(2))
sy(3) = (3000 - sy(3))

'If sx(1) > (sx(0)) Then
'dir = -1
'ElseIf sx(1) < (sx(0)) Then
'dir = 1
'Else
'dir = 0
'End If

'sx(2) = X * t(2) - winda(2) + dir * visca(2)
'sx(3) = X * t(3) - winda(3) + dir * visca(2)


For b = 0 To 3
    sx(b) = sx(b) Mod 7395
    If sx(b) < 0 Then sx(b) = sx(b) + 7395
Next b
If Abs(sx(2) - sx(3)) < 6000 Then Form1.Picture1.Line (sx(2), sy(2))-(sx(3), sy(3)), QBColor(5)
'
Wend
    If g > 0 Then sx(2) = Int(sx(2)) Else sx(2) = "Infinite"
    If g > 0 Then toth = Int(Y ^ 2 / (2 * g)) Else toth = "Infinite"
    Form1.Label3.Caption = "Initial y Velocity: " & Y & "    Initial x Velocity: " & X
    Form1.Label5.Caption = "Total Distance Travelled: " & sx(2) & "    Maximum Height Achieved: " & toth
    If g > 0 Then If (sx(2) > target) And (sx(2) < target + tgwidth) Then Form4.Visible = True
End Function
Function immediate()
Randomize
target = Rnd() * 6000 + 1000
Form1.Picture1.Line (0, 3000)-(8000, 3000), QBColor(2)                    'this command draws the ground
Form1.Image2.Left = target               'just a little target practise
tgwidth = 450
Form1.Image2.Width = tgwidth
g = 9.8
visc = 0
wind = 0
count = 0
If wind >= 0 Then Form1.Label4.Caption = "Wind = " & wind & " <<<<" Else Form1.Label4.Caption = "Wind = " & Abs(wind) & " >>>>"
Form1.Label7.Caption = "Gravity = " & g
Form1.Label8.Caption = "" '"Viscosity = " & visc
End Function
