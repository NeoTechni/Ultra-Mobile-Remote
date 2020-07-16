Attribute VB_Name = "CircularLCAR"
Option Explicit

Public Const cLCAR_Yellow As Long = 6750104 'rgb(152,255,102)
Public Const cLCAR_Green As Long = 231942 'rgb(6,138,3)
Public Const cLCAR_LightBlue As Long = 16764313 'rgb(153,205,255)
Public Const cLCAR_Blue As Long = 16646144 'rgb(0,0,254)

'Private Const Resolution As Long = 1
Private Const ColsPerQuadrant As Long = 10 'must be an even number ' * Resolution
Private Const Rows As Long = 10 '* Resolution
Private Const Width As Single = 360 / (ColsPerQuadrant * 4)
Private Const LineWidth As Single = Width / 8 '4 per grid with equal whitespace

Public Enum LineType
    NoLine
    Color1Line
    Color2Line
End Enum

Public Enum GridType
    Blank
    aCircle 'O
    aSquare '[]
    SemiCircle '(c
    Lines '=
    Bar 'u
    GridLine
End Enum

Public Type GridSegment
    SegmentType As GridType
    
    Top As Single
    Bottom As Single
    Left As Single
    Right As Single
    
    Color As Long
    BlinkColor As Long
    Blinking As Boolean
End Type

Public Type CirLCAR
    Grid(1 To Rows, 1 To ColsPerQuadrant * 4) As GridSegment
End Type

Public CircleID As Long, CircleName As String, CircleRow As Long, CircleCol As Long, CircleTool As Long, isdown2 As Boolean
Public CircleMode As GridType, CircleDiameter As Single, CircleLeft As Single, CircleRight As Single, CircleTop As Single, CircleBottom As Single
Public CircleColor As Long, CircleBlinkColor As Long, CircleLines(0 To 3) As Long

Dim temp As CirLCAR

Public Sub ShutdownEngine()
    Form1.TimerBlink.Enabled = False
    Form1.TimerEffects.Enabled = False
    Form1.Cls
End Sub

Public Function CircleColsPerQuadrant() As Long
    CircleColsPerQuadrant = ColsPerQuadrant
End Function

Public Function CircleCols() As Long
    CircleCols = ColsPerQuadrant * 4
End Function
Public Function CircleRows() As Long
    CircleRows = Rows
End Function
Public Function CircleColWidth() As Single
    CircleColWidth = Width
End Function

Public Sub DrawCirLCAR(Circ As CirLCAR, X As Long, Y As Long, Radius As Long, Optional Blink As Boolean)
    Dim Row As Long, Col As Long, CurrAngle As Single, Height As Long, Color As Long, TempColor As Long
    Dim cStart As Long, cFinish As Long, temp As Double, temp2 As Long, X2 As Long, Y2 As Long, Col2 As Long
    
    Height = Radius / Rows
    cFinish = Height
    For Row = 1 To Rows
        If isRotated Then CurrAngle = 180 Else CurrAngle = 90      '- Width
        
        For Col = 1 To ColsPerQuadrant * 4
            If isRotated Then
                Col2 = (Col + ColsPerQuadrant) Mod ColsPerQuadrant * 4
            Else
                Col2 = Col
            End If
            With Circ.Grid(Row, Col2)
                If Blink And .Blinking Then Color = .BlinkColor Else Color = .Color
                If RedAlert And Color <> vbBlack Then If Blink Then Color = LCAR_White Else Color = LCAR_Red
                
                Select Case .SegmentType
                    Case aCircle
                        temp = DegreesToRadians(CorrectAngle(CurrAngle + 90 - (Width / 2) - (.Left * Width)))            '-
                        X2 = findXY(CSng(X), CSng(Y), cStart + (Height / 2) + (Height * .Bottom), temp, True)
                        Y2 = findXY(CSng(X), CSng(Y), cStart + (Height / 2) + (Height * .Bottom), temp, False)
                        DrawSemiCircle X2, Y2, .Top * (Height / 2), 0, 360, Color, Color, , , 0
                    Case aSquare 'size is relative to radius, not height/width
                        temp2 = -1
                        Select Case Col
                            Case 1, ColsPerQuadrant * 4:                            temp2 = 0 'Top
                            Case ColsPerQuadrant * 3, ColsPerQuadrant * 3 + 1:      temp2 = 6 'left
                            Case ColsPerQuadrant, ColsPerQuadrant + 1:              temp2 = 2 'Right
                            Case ColsPerQuadrant * 2, ColsPerQuadrant * 2 + 1:      temp2 = 4 'Bottom
                            
                            Case ColsPerQuadrant / 2, ColsPerQuadrant / 2 + 1:      temp2 = 1 'top right
                            Case ColsPerQuadrant * 1.5, ColsPerQuadrant * 1.5 + 1:  temp2 = 3
                            Case ColsPerQuadrant * 2.5, ColsPerQuadrant * 2.5 + 1:  temp2 = 5
                            Case ColsPerQuadrant * 3.5, ColsPerQuadrant * 3.5 + 1:  temp2 = 7
                        End Select
                        If temp2 > -1 Then CirLCAR_DrawSquare X, Y, Height * .Right, Height * .Left, temp2, Color, Height * (Row - 1)
                    Case SemiCircle
                        'DrawSemiCircle X, Y, cStart + (.Top * Height), CurrAngle + (.Left * Width), (.Right * Width) - (.Left * Width), Color, Color, 2, , cStart + (.Bottom * Height)
                        If .Left < 0 Then
                            DrawSemiCircle X, Y, cStart + (.Top * Height), CurrAngle - (.Right * Width), (.Right - .Left) * Width, Color, Color, 2, , cStart + (.Bottom * Height), , 1
                            'DrawSemiCircle X, Y, cStart + (.Top * Height), CurrAngle + (.Left * Width) - ((.Right - .Left) * (Width)), (.Right - .Left) * Width, Color, Color, 2, , cStart + (.Bottom * Height), , 1
                        Else
                            DrawSemiCircle X, Y, cStart + (.Top * Height), CurrAngle - ((.Right + .Left) * Width) + (.Left * Width), (.Right - .Left) * Width, Color, Color, 2, , cStart + (.Bottom * Height), , 1
                        End If
                    Case Lines 'yes I could compress this but I don't have the time
                        temp2 = CLng(CurrAngle) '+ Width
                        CirDrawLine X, Y, temp2, cStart, Height, .Top, .Color, .BlinkColor, 2
                        temp2 = temp2 - LineWidth * 2
                        CirDrawLine X, Y, temp2, cStart, Height, .Bottom, .Color, .BlinkColor, 2
                        temp2 = temp2 - LineWidth * 2
                        CirDrawLine X, Y, temp2, cStart, Height, .Left, .Color, .BlinkColor, 2
                        temp2 = temp2 - LineWidth * 2
                        CirDrawLine X, Y, temp2, cStart, Height, .Right, .Color, .BlinkColor, 2
                    Case Bar
                        temp2 = cStart + (Height * 0.5)
                        'DrawSemiCircle X, Y, temp2,  CurrAngle + (((1 - .Top) / 2) * Width),  Width * .Top, Color, Color, Height / 2, 1, temp2
                        DrawSemiCircle X, Y, temp2, CurrAngle - ((.Left + .Right) * Width), (.Right * Width) - (.Left * Width), Color, Color, Height / 2, , temp2
                    Case GridLine
                        'DrawSemiCircle X, Y, cStart, CurrAngle - Width, Width, Color, Color, 1, 1, cStart
                        CirDrawLine X, Y, CLng(CurrAngle), cStart, Height, 1, Color, Color, 1
                        'CirDrawLine X, Y, CLng(CurrAngle) - Width, cStart, Height, 1, Color, Color, 1
                        DrawSemiCircle X, Y, cStart + Height, CurrAngle - Width, Width, Color, Color, 1, 1, cStart + Height
                End Select
            End With
            CurrAngle = CurrAngle - Width
            If CurrAngle < 0 Then CurrAngle = CurrAngle + 360
        Next
        
        cStart = cStart + Height
        cFinish = cStart + Height
    Next
End Sub

Public Function CirDrawLine(X As Long, Y As Long, Angle As Long, Radius As Long, Length As Long, State As Single, Color As Long, BlickColor As Long, DrawWidth As Long)
    Dim temp As Long, temp2 As Long, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, Radians As Double
    If Angle < 0 Then Angle = Angle + 360
    If State > 0 Then
        If RedAlert Then
            If State = 1 Then temp = LCAR_Red Else temp = LCAR_White
        Else
            If State = 1 Then temp = Color Else temp = BlickColor
        End If
        
        Radians = DegreesToRadians(CorrectAngle(Angle + 90))
        X1 = findXY(CSng(X), CSng(Y), CSng(Radius), Radians, True)
        Y1 = findXY(CSng(X), CSng(Y), CSng(Radius), Radians, False)
        X2 = findXY(CSng(X), CSng(Y), CSng(Radius) + Length, Radians, True)
        Y2 = findXY(CSng(X), CSng(Y), CSng(Radius) + Length, Radians, False)
        
        temp2 = dest.DrawWidth
        dest.DrawWidth = DrawWidth
        dest.Line (X1, Y1)-(X2, Y2), temp
        dest.DrawWidth = temp2
    End If
End Function

Public Sub CirLCAR_DrawSquare(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, Length As Long, ByVal Angle As Long, Color As OLE_COLOR, Optional Start As Long)
    Dim temp As Long, temp2 As Long, X2 As Long, Y2 As Long, NewAngle As Long
    temp = Width / 2
    If Width Mod 1 = 0 Then Width = Width + 1
    Select Case Angle
        Case 0 '  |   up
            DrawSquare X - temp, Y - Length - Start, Width, Length, Color, Color
        Case 6 '-     left
            DrawSquare X - Length - Start, Y - temp, Length, Width, Color, Color
        Case 2 '    - right
            DrawSquare X + Start, Y - temp, Length, Width, Color, Color
        Case 4 '  |   down
            DrawSquare X - temp, Y + Start, Width, Length, Color, Color
        Case Else
            Select Case Angle
                Case 1: NewAngle = 135 ' /  up right
                Case 3: NewAngle = 45 ' \  down right
                Case 5: NewAngle = 315 '/   down left
                Case 7: NewAngle = 225  '\   up left
            End Select
            If isRotated Then NewAngle = CLng(CorrectAngle(NewAngle + 90))
            
            temp = Width / 3
            
            For temp2 = 0 To Length - 1
                X2 = findXY(CSng(X), CSng(Y), Start + temp2, DegreesToRadians(NewAngle), True)
                Y2 = findXY(CSng(X), CSng(Y), Start + temp2, DegreesToRadians(NewAngle), False)
                Select Case Angle
                    Case 1 ' /  up right
                        dest.Line (X2 - temp, Y2 - temp)-(X2 + temp + 1, Y2 + temp + 1), Color
                        dest.Line (X2 - temp + 1, Y2 - temp)-(X2 + temp + 2, Y2 + temp + 1), Color
                    Case 3 'down right
                        dest.Line (X2 + temp - 1, Y2 - temp + 1)-(X2 - temp, Y2 + temp), Color
                        dest.Line (X2 + temp - 2, Y2 - temp + 1)-(X2 - temp - 1, Y2 + temp), Color
                    Case 5 'down left
                        dest.Line (X2 - temp, Y2 - temp)-(X2 + temp + 1, Y2 + temp + 1), Color
                        dest.Line (X2 - temp - 1, Y2 - temp)-(X2 + temp, Y2 + temp + 1), Color
                    Case 7 'up left
                        dest.Line (X2 + temp - 1, Y2 - temp - 1)-(X2 - temp - 2, Y2 + temp), Color
                        dest.Line (X2 + temp - 2, Y2 - temp - 1)-(X2 - temp - 3, Y2 + temp), Color
                        'dest.Line (X2 + temp, Y2 - temp + 1)-(X2 - temp + 1, Y2 + temp), Color
                End Select
            Next
    End Select
End Sub

Public Function CirLCAR_SetBlank(Circ As CirLCAR, Row As Long, Col As Long) As Boolean
    If Row < 1 Or Row > Rows Then Exit Function
    If Col < 1 Or Col > ColsPerQuadrant * 4 Then Exit Function
    With Circ.Grid(Row, Col)
        .SegmentType = Blank
        .BlinkColor = LCAR_Black
        .Blinking = False
        .Color = LCAR_Black
        .Bottom = 0
        .Left = 0
        .Top = 0
        .Right = 0
    End With
    CirLCAR_SetBlank = True
End Function

Public Function CirLCAR_SetBar(Circ As CirLCAR, Row As Long, Col As Long, Optional Left As Single = 0, Optional Right As Single = 1, Optional Color As Long = cLCAR_Yellow, Optional BlinkColor As Long = -1) As Boolean
    If Row < 1 Or Row > Rows Then Exit Function
    If Col < 1 Or Col > ColsPerQuadrant * 4 Then Exit Function
    With Circ.Grid(Row, Col)
        .SegmentType = Bar
        .Color = Color
        .Left = Left
        .Right = Right
        If BlinkColor = -1 Then
            .Blinking = False
            .BlinkColor = LCAR_Black
        Else
            .Blinking = True
            .BlinkColor = BlinkColor
        End If
    End With
    CirLCAR_SetBar = True
End Function

Public Function CirLCAR_SetSemiCircle(Circ As CirLCAR, Row As Long, Col As Long, Top As Single, Bottom As Single, Left As Single, Right As Single, Optional Color As Long = cLCAR_Yellow, Optional BlinkColor As Long = -1) As Boolean
    If Row < 1 Or Row > Rows Then Exit Function
    If Col < 1 Or Col > ColsPerQuadrant * 4 Then Exit Function
    If Top < Bottom Or Left > Right Then Exit Function
    
    With Circ.Grid(Row, Col)
        .SegmentType = SemiCircle
        
        .Top = Top
        .Bottom = Bottom
        .Left = Left
        .Right = Right
        
        .Color = Color
        If BlinkColor = -1 Then
            .Blinking = False
            .BlinkColor = LCAR_Black
        Else
            .Blinking = True
            .BlinkColor = BlinkColor
        End If
    End With
    
    CirLCAR_SetSemiCircle = True
End Function
Public Function CirLCAR_SetCircle(Circ As CirLCAR, Row As Long, Col As Long, Diameter As Single, Optional Color As Long = cLCAR_Yellow, Optional BlinkColor As Long = -1, Optional Left As Single, Optional Bottom As Single) As Boolean
    If Row < 1 Or Row > Rows Then Exit Function
    If Col < 1 Or Col > ColsPerQuadrant * 4 Then Exit Function
    
    With Circ.Grid(Row, Col)
        .SegmentType = aCircle
        
        .Top = Diameter
        .Color = Color
        
        .Left = Left
        .Bottom = Bottom
        
        If BlinkColor = -1 Then
            .Blinking = False
            .BlinkColor = LCAR_Black
        Else
            .Blinking = True
            .BlinkColor = BlinkColor
        End If
    End With
    
    CirLCAR_SetCircle = True
End Function
Public Function CirLCAR_SetLines(Circ As CirLCAR, Row As Long, Col As Long, Optional Color1 As Long = cLCAR_Yellow, Optional Color2 As Long = cLCAR_Green, Optional Bar1 As LineType, Optional Bar2 As LineType, Optional Bar3 As LineType, Optional Bar4 As LineType) As Boolean
    If Row < 1 Or Row > Rows Then Exit Function
    If Col < 1 Or Col > ColsPerQuadrant * 4 Then Exit Function
    If Bar1 < 0 Or Bar1 > 2 Then Exit Function
    If Bar2 < 0 Or Bar2 > 2 Then Exit Function
    If Bar3 < 0 Or Bar3 > 2 Then Exit Function
    If Bar4 < 0 Or Bar4 > 2 Then Exit Function
    
    With Circ.Grid(Row, Col)
        .SegmentType = Lines
        
        .Top = Bar1
        .Bottom = Bar2
        .Left = Bar3
        .Right = Bar4
        
        .Blinking = False
        .Color = Color1
        .BlinkColor = Color2
    End With
    CirLCAR_SetLines = True
End Function

Public Function CirLCAR_SetGridline(Circ As CirLCAR, Row As Long, Col As Long, Optional Color As Long = cLCAR_Yellow, Optional BlinkColor As Long = -1) As Boolean
    If Row < 1 Or Row > Rows Then Exit Function
    If Col < 1 Or Col > ColsPerQuadrant * 4 Then Exit Function
    With Circ.Grid(Row, Col)
        .SegmentType = GridLine
        .Color = Color
        If BlinkColor = -1 Then
            .Blinking = False
            .BlinkColor = LCAR_Black
        Else
            .Blinking = True
            .BlinkColor = BlinkColor
        End If
    End With
End Function

Public Function CirLCAR_SetSquare(Circ As CirLCAR, Row As Long, Col As Long, Width As Single, Length As Single, Optional Color As Long = cLCAR_Yellow, Optional BlinkColor As Long = -1) As Boolean
    If Row < 1 Or Row > Rows Then Exit Function
    If Col < 1 Or Col > ColsPerQuadrant * 4 Then Exit Function
    With Circ.Grid(Row, Col)
        .SegmentType = aSquare
        .Color = Color
        If BlinkColor = -1 Then
            .Blinking = False
            .BlinkColor = LCAR_Black
        Else
            .Blinking = True
            .BlinkColor = BlinkColor
        End If
        
        .Left = Length
        .Right = Width
    End With
End Function

Public Function CirLCAR_SetAllGridlines(Circ As CirLCAR, Optional Color As Long = cLCAR_Yellow, Optional BlinkColor As Long = -1) As Boolean
    Dim temp As Long, temp2 As Long
    For temp = 1 To Rows
        For temp2 = 1 To ColsPerQuadrant * 4
            If Color = -1 Then
                CirLCAR_SetBlank Circ, temp, temp2
            Else
                CirLCAR_SetGridline Circ, temp, temp2, Color, BlinkColor
            End If
        Next
    Next
    CirLCAR_SetAllGridlines = True
End Function

Public Sub TestCirLCAR()
    Dim temp2 As Long

    CirLCAR_SetAllGridlines temp, LCAR_Red
    temp2 = LCAR_AddCircle("circGridlines", -205, 279, 200, True, 8) '305
    LCARCircleList(temp2).Circ = temp
    
    CirLCAR_SetAllGridlines temp, -1
    
    If False Then
        CirLCAR_SetSemiCircle temp, 4, 1, 1, 0, 0, 1
        CirLCAR_SetSemiCircle temp, 4, 2, 0.75, 0.25, 0, 1, cLCAR_Blue
        CirLCAR_SetSemiCircle temp, 4, 4, 1, 0, 0, 1
        CirLCAR_SetSemiCircle temp, 4, 5, 0.75, 0.25, 0, 1, cLCAR_LightBlue
    
        CirLCAR_SetCircle temp, 4, 6, 0.75
        CirLCAR_SetCircle temp, 4, 7, 0.75, cLCAR_Blue
        CirLCAR_SetCircle temp, 4, 8, 0.75, cLCAR_LightBlue
    
        CirLCAR_SetBar temp, 4, 9, 0.5
    
        CirLCAR_SetLines temp, 4, 10, , , Color1Line, Color2Line, Color1Line, Color2Line
    
        CirLCAR_SetGridline temp, 4, 11
    
        CirLCAR_SetSemiCircle temp, 6, 1, 1, 0, 0, 1
        CirLCAR_SetSemiCircle temp, 6, 2, 0.75, 0.25, 0, 1, cLCAR_Blue
        CirLCAR_SetSemiCircle temp, 6, 4, 1, 0, 0, 1
        CirLCAR_SetSemiCircle temp, 6, 5, 0.75, 0.25, 0, 1, cLCAR_LightBlue
    End If
    
    'DrawCirLCAR temp, 200, 200, 200
    temp2 = LCAR_AddCircle("circTest", -205, 279, 200, True, 8) '305
    LCARCircleList(temp2).Circ = temp
    
    CircleDiameter = 0.1
    CircleRight = 1
    CircleTop = 1
End Sub

Public Function SaveCirLCAR(Cir As CirLCAR) As String
        Const D As String = " "
        Dim Row As Long, Col As Long, tempstr As String, tempstr2 As String
        For Row = 1 To Rows
            tempstr2 = Empty
            For Col = 1 To ColsPerQuadrant * 4
                With Cir.Grid(Row, Col)
                    tempstr2 = tempstr2 & .SegmentType & D & LCAR_ColorIDfromColor(.Color) & D & LCAR_ColorIDfromColor(.BlinkColor) & D & .Top & D & .Bottom & D & .Left & D & .Right & D
                End With
            Next
            tempstr = tempstr & tempstr2
        Next
        SaveCirLCAR = Left(tempstr, Len(tempstr) - 1)
End Function

Public Function LoadCirLCAR(Cir As CirLCAR, Text As String, Optional StartCol As Long = 1, Optional EndCol As Long = -1) As Boolean
    Dim Row As Long, Col As Long, tempstr() As String, temp As Long, Required As Long
    If EndCol < StartCol Then EndCol = ColsPerQuadrant * 4
    Required = (7 * Rows * ColsPerQuadrant * 4) - 1
    tempstr = Split(Text, " ")
    Row = 1
    Col = 1
    If UBound(tempstr) <> Required Then Exit Function
    
    For temp = 0 To UBound(tempstr) Step 7
        If Col >= StartCol And Col <= EndCol Then
            With Cir.Grid(Row, Col)
                .SegmentType = Val(tempstr(temp))
                .Color = ColorList(Val(tempstr(temp + 1))).Color
                .BlinkColor = ColorList(Val(tempstr(temp + 2))).Color
                .Top = Val(tempstr(temp + 3))
                .Bottom = Val(tempstr(temp + 4))
                .Left = Val(tempstr(temp + 5))
                .Right = Val(tempstr(temp + 6))
            End With
        End If
        
        Col = Col + 1
        If Col > ColsPerQuadrant * 4 Then
            Col = 1
            Row = Row + 1
        End If
    Next
    LoadCirLCAR = True
End Function

Public Function LoadFile(Filename As String) As String
    On Error Resume Next
    If FileLen(Filename) = 0 Then Exit Function
    Dim temp As Long, tempstr As String, tempstr2 As String
    temp = FreeFile
    If Dir(Filename) <> Filename Then
        Open Filename For Input As temp
            Do Until EOF(temp)
                Line Input #temp, tempstr
                If tempstr2 <> Empty Then tempstr2 = tempstr2 & vbNewLine
                tempstr2 = tempstr2 & tempstr
                DoEvents
            Loop
            LoadFile = tempstr2
        Close temp
    End If
End Function
Public Function SaveFile(Filename As String, Contents As String) As Boolean
    On Error Resume Next
    Dim temp As Long
    temp = FreeFile
    If Filename Like "?:\*" Then
        Open Filename For Output As temp
            Print #temp, Contents
        Close temp
    End If
    SaveFile = True
End Function
