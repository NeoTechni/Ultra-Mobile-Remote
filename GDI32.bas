Attribute VB_Name = "GDI32"
Option Explicit

Public Declare Function LineTo Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Public Declare Function SelectObject Lib "GDI32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Public Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function PolyBezier Lib "gdi32.dll" (ByVal hdc As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Public Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal hdc As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Public Declare Function PolyPolygon Lib "gdi32.dll" (ByVal hdc As Long, lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long) As Long
Public Declare Function RectangleX Lib "GDI32" Alias "Rectangle" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Public Declare Function CreateHatchBrush Lib "GDI32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Public Declare Function BeginPath Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function EndPath Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function StrokePath Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function StrokeAndFillPath Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function SetGraphicsMode Lib "GDI32" (ByVal hdc As Long, ByVal iMode As Long) As Long
Public Declare Function SetWorldTransform Lib "GDI32" (ByVal hdc As Long, lpXform As XFORM) As Long
Public Declare Function SetMapMode Lib "GDI32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function DPtoLP Lib "GDI32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Public Declare Function GetWorldTransform Lib "GDI32" (ByVal hdc As Long, lpXform As XFORM) As Long
Public Declare Function GetMapMode Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function SetBkColor Lib "GDI32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetGraphicsMode Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function ModifyWorldTransform Lib "GDI32" (ByVal hdc As Long, lpXform As XFORM, ByVal iMode As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "GDI32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function PaintRgn Lib "GDI32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function PtInRegion Lib "GDI32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type XFORM
        eM11 As Single
        eM12 As Single
        eM21 As Single
        eM22 As Single
        eDx As Single
        eDy As Single
End Type
Public Const MWT_IDENTITY = 1
Public Const MWT_LEFTMULTIPLY = 2
Public Const MWT_RIGHTMULTIPLY = 3
Public Const MM_LOMETRIC = 2
Public Const MM_LOENGLISH = 4
Public Const GM_ADVANCED = 2
Public Const PS_GEOMETRIC = &H10000
Public Const PS_ENDCAP_FLAT = &H200
Public Const HS_BDIAGONAL = 3               '  /////
Public Const HS_CROSS = 4                   '  +++++
Public Const HS_DIAGCROSS = 5               '  xxxxx
Public Const HS_FDIAGONAL = 2               '  \\\\\
Public Const HS_HORIZONTAL = 0              '  -----
Public Const HS_SOLID = 8
Public Const PS_JOIN_BEVEL = &H1000
Public Const BS_SOLID = 0
Public Const PS_SOLID = 0
