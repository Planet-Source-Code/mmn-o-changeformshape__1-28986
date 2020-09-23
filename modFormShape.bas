Attribute VB_Name = "modFormShape"
Option Explicit

Enum MNShape            ' M(ikael)N(ordfelth)Shape
  Elliptic = 1
  Rectangle = 2
  RoundedRectangle = 3
  Polygon = 4
End Enum

Type PointAPI
  x As Long
  y As Long
End Type

Public Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" _
    (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, _
    ByVal RectY2 As Long, ByVal EllipseWidth As Long, _
    ByVal EllipseHeight As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As PointAPI, ByVal nCount As Long, ByVal nPolyFillMode _
    As Long) As Long

Declare Function CombineRgn Lib "gdi32" _
    (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
    ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
    (ByVal hWnd As Long, ByVal hRgn As Long, _
    ByVal blnRedraw As Boolean) As Long


Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long


Public Const RGN_AND = 1&
Public Const RGN_OR = 2&
Public Const RGN_XOR = 3&
Public Const RGN_DIFF = 4&
Public Const RGN_COPY = 5&
