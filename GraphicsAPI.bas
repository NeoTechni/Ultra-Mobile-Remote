Attribute VB_Name = "GraphicsAPI"
Option Explicit

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Const GradientSteps As Long = 32
Public Const Brightness As Long = 128

Public Type GradientCache
    StartColor As OLE_COLOR
    EndColor As OLE_COLOR
    
    Cache(GradientSteps) As OLE_COLOR
    
    LightColor As OLE_COLOR
    DarkColor As OLE_COLOR
End Type

Public GradientList() As GradientCache, GradientCount As Long, AntiAliasing As Boolean

Public Const PI = 3.14159265358979
Public Const PI2 = PI * 2
Public Const NotPi = PI / 180
Public Const InvPi = 180 / PI

Public Function IsGradientReversed(GradientID As Long, StartColor As Long, EndColor As Long) As Boolean
    With GradientList(GradientID)
        IsGradientReversed = (.StartColor = EndColor) And (.EndColor = StartColor)
    End With
End Function
Public Function FindGradient(StartColor As Long, EndColor As Long) As Long
    Dim temp As Long
    FindGradient = -1
    For temp = 0 To GradientCount - 1
        With GradientList(temp)
            If .StartColor = StartColor And .EndColor = EndColor Then
                FindGradient = temp
                Exit For
            ElseIf .StartColor = EndColor And .EndColor = StartColor Then
                FindGradient = temp
                Exit For
            End If
        End With
    Next
End Function
Public Function CacheGradient(ByVal StartColor As Long, ByVal EndColor As Long, Optional Brightness As Byte = 64) As Long
    Dim dR As Double, dG As Double, dB As Double, cR As Double, cG As Double, cB As Double, aR As Double, aG As Double, aB As Double, temp As Long
    Dim color As Long, R As Long, G As Long, B As Long
    
    color = FindGradient(StartColor, EndColor)
    If color > -1 Then
        CacheGradient = color
        Exit Function
    End If
    
    CacheGradient = GradientCount
    GradientCount = GradientCount + 1
    ReDim Preserve GradientList(GradientCount)
    
    aR = Red(EndColor)
    aG = Green(EndColor)
    aB = Blue(EndColor)
    
    R = MinMax(aR - Brightness, 0, 255)
    G = MinMax(aG - Brightness, 0, 255)
    B = MinMax(aB - Brightness, 0, 255)
    GradientList(GradientCount - 1).DarkColor = RGB(R, G, B)
    
    cR = Red(StartColor)
    cG = Green(StartColor)
    cB = Blue(StartColor)
    
    R = MinMax(cR + Brightness, 0, 255)
    G = MinMax(cG + Brightness, 0, 255)
    B = MinMax(cB + Brightness, 0, 255)
    GradientList(GradientCount - 1).LightColor = RGB(R, G, B)
    
    dR = AlphaIncrement(CInt(cR), Red(EndColor), GradientSteps)
    dG = AlphaIncrement(CInt(cG), Green(EndColor), GradientSteps)
    dB = AlphaIncrement(CInt(cB), Blue(EndColor), GradientSteps)
    
    color = StartColor
    For temp = 0 To GradientSteps
        cR = cR - dR
        cG = cG - dG
        cB = cB - dB
        
        cR = MinMax(CLng(cR), 0, 255)
        cG = MinMax(CLng(cG), 0, 255)
        cB = MinMax(CLng(cB), 0, 255)
        
        color = RGB(CInt(cR), CInt(cG), CInt(cB))
        GradientList(GradientCount - 1).Cache(temp) = color
    Next
    GradientList(GradientCount - 1).StartColor = StartColor
    GradientList(GradientCount - 1).EndColor = EndColor
End Function
Public Function GetGradientColor(GradientID As Long, ByVal step As Long, Reversed As Boolean) As Long
    If Reversed Then step = GradientSteps - step
    If step = -1 Then
        If Reversed Then
            GetGradientColor = GradientList(GradientID).EndColor
        Else
            GetGradientColor = GradientList(GradientID).StartColor
        End If
    ElseIf step = GradientSteps + 1 Then
        If Reversed Then
            GetGradientColor = GradientList(GradientID).StartColor
        Else
            GetGradientColor = GradientList(GradientID).EndColor
        End If
    Else
        GetGradientColor = GradientList(GradientID).Cache(step)
    End If
End Function

Public Function findXY(X As Single, Y As Single, Distance As Single, Angle As Double, Optional isx As Boolean = True) As Single
    If isx Then findXY = X + Sin(Angle) * Distance Else findXY = Y + Cos(Angle) * Distance
End Function
Public Function DegreesToRadians(Degrees As Long) As Double 'Converts Degrees to Radians.
    DegreesToRadians = (Degrees Mod 360) * NotPi
End Function

Public Function AlterBrightness(ByVal color As OLE_COLOR, Brightness As Long, Optional ForceChange As Boolean) As Long
    Dim R As Long, G As Long, B As Long
    color = SysToLNG(color)
    R = MinMax(Red(color) + Brightness, 0, 255, ForceChange)
    G = MinMax(Green(color) + Brightness, 0, 255, ForceChange)
    B = MinMax(Blue(color) + Brightness, 0, 255, ForceChange)
    AlterBrightness = RGB(R, G, B)
End Function

Public Function MinMax(Number As Long, Minimum As Long, Maximum As Long, Optional ForceChange As Boolean) As Long
    MinMax = Number
    If ForceChange Then
        If Number < Minimum Then MinMax = Number + Maximum
        If Number > Maximum Then MinMax = Number Mod Maximum
    Else
        If Number < Minimum Then MinMax = Minimum
        If Number > Maximum Then MinMax = Maximum
    End If
End Function



'COLOR
Public Function Red(color As Long)
    Red = color Mod 256
End Function
Public Function Green(color As Long)
    Green = ((color And &HFF00) / 256) Mod 256
End Function
Public Function Blue(color As Long)
    Blue = (color And &HFF0000) / 65536
End Function
Public Function AlphaBlend(ColorA As Long, ColorB As Long, Alpha As Double) As Long
    Dim R As Long, G As Long, B As Long
    R = blend(Red(ColorA), Red(ColorB), Alpha)
    G = blend(Green(ColorA), Green(ColorB), Alpha)
    B = blend(Blue(ColorA), Blue(ColorB), Alpha)
    AlphaBlend = RGB(R, G, B)
End Function
Public Function blend(ColorA As Long, ColorB As Long, Alpha As Double) As Long
    blend = Abs((ColorA - ColorB) * Alpha + ColorB) Mod 256
End Function
Public Function SysToLNG(ByVal lColor As Long) As Long 'Special thanks to redbird77 for this code and realizing what the bug was
    SysToLNG = lColor ' If hi-bit is set, then it is a system color.
    If (lColor And &H80000000) Then SysToLNG = GetSysColor(lColor And &HFFFFFF)
End Function

Public Function AlphaIncrement(ColorA As Long, ColorB As Long, Steps As Long) As Double
    AlphaIncrement = (ColorA - ColorB) / Steps
End Function




'Gradient code
Public Sub DrawRhomboid(color As OLE_COLOR, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, HasL As Boolean, HasR As Boolean, PointUp As Boolean, Rotate As Boolean)
    Dim AddLeft As Long, AddWidth As Long, AddHeight As Long, AddTop As Long, temp As Long, Steps As Long
    
    Steps = Height
    If Rotate Then Steps = Width

    If Rotate Then
        If PointUp Then
            If HasL Then
                AddHeight = -1
                AddTop = 1
            End If
            If HasR Then
                Height = Height + Steps
                AddHeight = AddHeight - 1
            End If
        Else
            If HasL Then
                Height = Height - Steps
                Y = Y + Steps
                AddTop = -1
                AddHeight = 1
            End If
            If HasR Then
                Height = Height - Steps
                AddHeight = AddHeight + 1
            End If
        End If
        AddLeft = 1
    Else
        If PointUp Then
            If HasL Then
                AddLeft = 1
                AddWidth = -1
            End If
            If HasR Then
                Width = Width + Steps
                AddWidth = AddWidth - 1
            End If
        Else
            If HasL Then ' AddLeft = 1: AddWidth = -1
                Width = Width - Steps
                X = X + Steps
                AddLeft = -1
                AddWidth = 1
            End If
            If HasR Then 'Width = Width + Steps: AddWidth = AddWidth - 1
                Width = Width - Steps
                AddWidth = AddWidth + 1
            End If
        End If
        AddTop = 1
    End If
    
    dest.DrawWidth = 1
    For temp = 1 To Steps
        If Rotate Then
            dest.Line (X, Y)-(X, Y + Height - 1), color
        Else
            dest.Line (X, Y)-(X + Width - 1, Y), color
        End If
        
        Width = Width + AddWidth
        X = X + AddLeft
        Y = Y + AddTop
        Height = Height + AddHeight
    Next
End Sub


Public Sub DrawGradientSemiCircle(hDC As Long, GID As Long, IsRev As Boolean, X As Long, Y As Long, StartRadius As Long, EndRadius As Long, Factor As Double, ByVal StartAngle As Long, Width As Long)
    Const TwoPi As Double = 2 * PI
    Dim AngleSteps As Long, temp As Long, temp2 As Long, EndAngle As Long, color As Long, AngleWidth As Long
    Dim CAngle As Long, IncAngle As Long
    
    EndAngle = StartAngle + Width
    AngleSteps = Width / (GradientSteps + 1)
    AngleWidth = Abs(AngleSteps)
    
    CAngle = StartAngle
    IncAngle = AngleWidth
    If Width < 0 Then
        CAngle = CAngle + AngleWidth
        AngleWidth = Abs(AngleWidth)
        IncAngle = -AngleWidth
    End If
    
    For temp = -1 To GradientSteps + 1
        color = GetGradientColor(GID, temp, IsRev)
        If CAngle < 0 Then CAngle = CAngle + 360
        If CAngle < 345 Then DrawSemiCircle X, Y, EndRadius, CAngle, AngleWidth, -1, color, , , StartRadius
        CAngle = CAngle + IncAngle
    Next
End Sub

Public Function DrawCurvedGradientSquare(hDC As Long, ColorA As Long, ColorB As Long, X As Long, Y As Long, Width As Long, Height As Long) As Long
    Const Start As Double = 1.5 * PI
    Const Finish As Double = 2 * PI
    
    Dim Radius As Double, cmiddle As Double, FillColor As Long, temp As Long
    Dim GID As Long, IsReversed As Boolean, Size As Long, Steps As Long
    
    Steps = 1
    Size = Distance(0, 0, CSng(Width), CSng(Height))
    
    Radius = Size / GradientSteps
    GID = CacheGradient(ColorA, ColorB)
    IsReversed = IsGradientReversed(GID, ColorA, ColorB)
    DrawCurvedGradientSquare = GID
    
    If Radius < 1 Then
        Steps = 1 / Radius
        Radius = Radius * Steps
        
    End If
    
    FillColor = ColorA
    cmiddle = Radius / 2
    dest.DrawWidth = Radius + 1
    dest.Circle (X, Y), 0, FillColor, Start, Finish
    For temp = 0 To GradientSteps Step Steps
        FillColor = GetGradientColor(GID, temp, IsReversed)
        dest.Circle (X, Y), cmiddle, FillColor, Start, Finish  ', Factor
        cmiddle = cmiddle + Radius
    Next
End Function



















Public Sub DrawGradientSquare(hDC As Long, ColorA As Long, ColorB As Long, X As Long, Y As Long, Width As Long, Height As Long)
    Dim dR As Double, dG As Double, dB As Double, cR As Double, cG As Double, cB As Double, aR As Double, aG As Double, aB As Double, temp As Long
    
    aR = Red(ColorB)
    aG = Green(ColorB)
    aB = Blue(ColorB)
    
    cR = Red(ColorA)
    cG = Green(ColorA)
    cB = Blue(ColorA)
    
    dR = AlphaIncrement(CInt(cR), Red(ColorB), Height)
    dG = AlphaIncrement(CInt(cG), Green(ColorB), Height)
    dB = AlphaIncrement(CInt(cB), Blue(ColorB), Height)
        
    For temp = Y To Y + Height - 1
        'DrawGradientLine hdc, RGB(CInt(cR), CInt(cG), CInt(cB)), RGB(CInt(aR), CInt(aG), CInt(aB)), X, temp, Width
        
        cR = cR - dR
        cG = cG - dG
        cB = cB - dB
        
        aR = aR + dR
        aG = aG + dG
        aB = aB + dB
    Next
End Sub

Public Sub DrawGradientVerticalLine(hDC As Long, X As Long, Y As Long, Height As Long, GID As Long, Optional IsRev As Boolean, Optional Width As Long = 1)
    Dim dR As Double, dG As Double, dB As Double, cR As Double, cG As Double, cB As Double, temp As Long, ColorA As Long, ColorB As Long
    Dim step As Long
    'If scrY < 0 Or scrY > getHeight Then Exit Sub
    
    With GradientList(GID)
    
    For temp = Y To Y + Height - 1
        'SetPixelV hdc, X, temp, RGB(CInt(cR), CInt(cG), CInt(cB))
    
        cR = cR - dR
        cG = cG - dG
        cB = cB - dB
    Next
    
    End With
    
    Exit Sub
    For temp = 2 To Width
        'BitBlt hdc, X + temp - 1, Y, 1, Height, hdc, X, Y, vbSrcCopy
    Next
End Sub

Public Sub DrawGradientLine(hDC As Long, ByVal X As Double, ByVal Y As Double, Width As Long, GID As Long, Optional IsRev As Boolean, Optional Height As Long = 1, Optional Rotate As Boolean)
    Dim temp As Long, step As Long, Inc As Double, color As Long
    step = -1
    If Rotate Then
        Inc = Height / GradientSteps
    Else
        Inc = Width / GradientSteps
    End If
    
    For temp = 0 To GradientSteps
        color = GetGradientColor(GID, temp, IsRev)
        If Rotate Then
            dest.Line (X, Y)-(X + Width - 1, Y + Inc - 1), color, B
            Y = Y + Inc
        Else
            dest.Line (X, Y)-(X + Inc - 1, Y + Height - 1), color, B
            X = X + Inc
        End If
    Next
End Sub






Public Function GetXYIntercept(X1 As Single, Y1 As Single, Angle As Double, X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, ByRef X As Single, ByRef Y As Single) As Boolean
    Dim X2 As Single, Y2 As Single
    Const Distance As Single = 100 'Example number
    X2 = findXY(X1, Y1, Distance, Angle, True)
    Y2 = findXY(X1, Y1, Distance, Angle, False)
    GetXYIntercept = LineLineIntercept(X1, Y1, X2, Y2, X3, Y3, X4, Y4, CLng(X), CLng(Y))
End Function
'Intersections, obtained elsewhere
Public Function LineLineIntercept(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, X3 As Single, Y3 As Single, X4 As Single, Y4 As Single, ByRef X As Long, ByRef Y As Long) As Boolean
    Dim a1 As Single, a2 As Single, b1 As Single, b2 As Single, c1 As Single, c2 As Single, denom As Single
    'Translated from Pascal, lost source
    a1 = Y2 - Y1
    b1 = X1 - X2
    c1 = X2 * Y1 - X1 * Y2 '  { a1*x + b1*y + c1 = 0 is line 1 }

    a2 = Y4 - Y3
    b2 = X3 - X4
    c2 = X4 * Y3 - X3 * Y4 '  { a2*x + b2*y + c2 = 0 is line 2 }

    denom = a1 * b2 - a2 * b1

    If denom <> 0 Then
        LineLineIntercept = True
        X = (b1 * c2 - b2 * c1) / denom
        Y = (a2 * c1 - a1 * c2) / denom
    End If
End Function
Public Function LineCircleIntersept(Cx As Single, Cy As Single, Radius As Long, X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, ByRef ix1 As Single, ByRef iy1 As Single, Optional ByRef ix2 As Single, Optional ByRef iy2 As Single, Optional OneResult As Boolean) As Integer
    Dim dX As Single, dY As Single, A As Single, B As Single, C As Single, det As Single, t As Single
    'http://www.vb-helper.com/howto_line_circle_intersections.html
    dX = X2 - X1
    dY = Y2 - Y1

    A = dX * dX + dY * dY
    B = 2 * (dX * (X1 - Cx) + dY * (Y1 - Cy))
    C = (X1 - Cx) * (X1 - Cx) + (Y1 - Cy) * (Y1 - Cy) - Radius * Radius

    det = B * B - 4 * A * C
    If (A <= 0.0000001) Or (det < 0) Then
        ' No real solutions.
    ElseIf det = 0 Then
        ' One solution.
        LineCircleIntersept = 1
        t = -B / (2 * A)
        ix1 = X1 + t * dX
        iy1 = Y1 + t * dY
    Else
        ' Two solutions.
        LineCircleIntersept = 2
        t = (-B + Sqr(det)) / (2 * A)
        ix1 = X1 + t * dX
        iy1 = Y1 + t * dY
        If Not OneResult Then 'Check if I only need 1 result
            t = (-B - Sqr(det)) / (2 * A)
            ix2 = X1 + t * dX
            iy2 = Y1 + t * dY
        End If
    End If
End Function


Public Function CorrectAngle(ByVal Angle As Single) As Single
    Angle = Angle Mod 360
    Do Until Angle >= 0
        Angle = Angle + 360
    Loop
    CorrectAngle = Angle
End Function

Public Function GetAngle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Long
    GetAngle = CorrectAngle(AngleBySection(X1, Y1, X2, Y2, RadiansToDegrees(Angle(X1, Y1, X2, Y2))) - 180)
End Function

Public Function RadiansToDegrees(ByVal Radians As Double) As Double 'Converts Radians to Degrees.
    RadiansToDegrees = Radians * InvPi
End Function

Public Function Angle(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Double
    On Error Resume Next
    Angle = Atn((Y2 - Y1) / (X1 - X2))
End Function

Public Function AngleBySection(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single, ByVal Angle As Long) As Double
    Angle = Abs(Angle)
    AngleBySection = 90 'Corrected
    If X1 < X2 Then 'the point is at the left of Center
        If Y1 = Y2 Then
            AngleBySection = 270 'Corrected
        ElseIf Y1 < Y2 Then
            If 270 + Angle = 360 Then
                AngleBySection = 0 'Corrected
            Else
                AngleBySection = 270 + Angle 'Corrected
            End If
        ElseIf Y1 > Y2 Then
            AngleBySection = 270 - Angle 'Corrected
        End If
    Else
    
        If X1 > X2 Then 'the point is at the right of Center
            If Y1 > Y2 Then
                AngleBySection = 90 + Angle 'Corrected
            ElseIf Y1 < Y2 Then
                AngleBySection = 90 - Angle 'Corrected
            End If
        Else
    
            If X1 = X2 Then
                If Y1 < Y2 Then
                    AngleBySection = 0 'Corrected
                ElseIf Y1 > Y2 Then
                    AngleBySection = 180 'Corrected
                End If
            End If
    
        End If

    End If
End Function

Public Function Distance(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
    On Error Resume Next
    If Y2 - Y1 = 0 Then Distance = Abs(X2 - X1): Exit Function
    If X2 - X1 = 0 Then Distance = Abs(Y2 - Y1): Exit Function
    Distance = Abs(Y2 - Y1) / Sin(Atn(Abs(Y2 - Y1) / Abs(X2 - X1)))
End Function

