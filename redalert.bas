Attribute VB_Name = "Module1"
Option Explicit

Public Sub DrawCondition(X As Long, Y As Long, Width As Long, Color As Long)
    Exit Sub

    Const BarHeight As Single = 0.2, BarWidth As Single = 0.1, BarCurve As Single = 10
    Const BeamWidth As Single = 0.6, BeamY As Single = 0.3, BeamHeight As Double = 0.02
    
    Dim MiddleY As Long, MiddleX As Long, BHeight As Long, BWidth As Long, CWidth As Long, Height As Long, Degree As Long, DegreeWidth As Long
    Height = Width
    DrawSquare X, Y, Width, Height, vbWhite
    
    'middle of the red alert square
    MiddleX = X + (Width \ 2) - 1
    MiddleY = Y + (Height \ 2) - 1
    
    'block dimensions
    BHeight = Height * BarHeight 'height of the block
    BWidth = Width * BarWidth 'width of the block not including the curve
    CWidth = BHeight / BarCurve 'width of the curve
    
    'left block
    DrawSquare X + CWidth, MiddleY - (BHeight / 2), BWidth, BHeight, Color, Color
    DrawSemiCircle X + CWidth, MiddleY, (BHeight / 2), 90, 180, Color, Color, , BarCurve
    
    'right block
    DrawSquare X + Width - 1 - BWidth - CWidth, MiddleY - (BHeight / 2), BWidth, BHeight, Color, Color
    DrawSemiCircle X + Width - 1 - CWidth, MiddleY, (BHeight / 2), 270, 180, Color, Color, , BarCurve
    
    'beam dimensions
    BWidth = Width * BeamWidth
    BHeight = Height * BeamHeight
    Degree = 360 - GetAngle(CSng(MiddleX), CSng(MiddleY), MiddleX - (BWidth \ 2), CSng(Y))
    DegreeWidth = 360 - GetAngle(CSng(MiddleX), CSng(MiddleY), CSng(X), MiddleY - (Height * BeamY / 2))
    
    'top beam
    DrawSquare MiddleX - (BWidth \ 2), MiddleY - (Height * BeamY / 2), BWidth, BHeight, Color, Color
    CWidth = Distance(0, 0, BWidth \ 2, Height * BeamY / 2)
    DrawSemiCircle MiddleX, MiddleY, Width \ 2, 90 + Degree, DegreeWidth, vbBlue, vbBlue, , , CWidth
    
    'bottom beam
    DrawSquare MiddleX - (BWidth \ 2), MiddleY + (Height * BeamY / 2) - BHeight, BWidth, BHeight, Color, Color
End Sub




Public Function Distance(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Single
    On Error Resume Next
    If Y2 - Y1 = 0 Then Distance = Abs(X2 - X1): Exit Function
    If X2 - X1 = 0 Then Distance = Abs(Y2 - Y1): Exit Function
    Distance = Abs(Y2 - Y1) / Sin(Atn(Abs(Y2 - Y1) / Abs(X2 - X1)))
End Function
