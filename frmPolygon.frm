VERSION 5.00
Begin VB.Form frmPolygon 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Octagon hole with circle in it"
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   2400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   735
      Left            =   780
      TabIndex        =   0
      Top             =   840
      Width           =   795
   End
   Begin VB.Shape shpDiagonal 
      FillStyle       =   7  'Diagonal Cross
      Height          =   1455
      Left            =   360
      Top             =   480
      Width           =   1635
   End
End
Attribute VB_Name = "frmPolygon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Dim Points(9) As PointAPI
  Dim PointsRemove(9) As PointAPI
  
  Dim hRgnForm As Long
  Dim hRgnTemp As Long
  
    '# COORDINATES #'
  
   ' The region that comes out from these coordinates will be used to make an _
     octagon-hole in the Form
  PointsRemove(1).X = Me.Width * 0.5 / 15
  PointsRemove(1).Y = 0

  PointsRemove(2).X = Me.Width * 0.15 / 15
  PointsRemove(2).Y = Me.Height * 0.15 / 15

  PointsRemove(3).X = 0
  PointsRemove(3).Y = Me.Height * 0.5 / 15

  PointsRemove(4).X = Me.Width * 0.15 / 15
  PointsRemove(4).Y = Me.Height * 0.85 / 15

  PointsRemove(5).X = Me.Width * 0.5 / 15
  PointsRemove(5).Y = Me.Height / 15

  PointsRemove(6).X = Me.Width * 0.85 / 15
  PointsRemove(6).Y = Me.Height * 0.85 / 15

  PointsRemove(7).X = Me.Width / 15
  PointsRemove(7).Y = Me.Height * 0.5 / 15

  PointsRemove(8).X = Me.Width * 0.85 / 15
  PointsRemove(8).Y = Me.Height * 0.15 / 15

  PointsRemove(9).X = Me.Width * 0.5 / 15
  PointsRemove(9).Y = 0
  
    ' Set the variables to the specified regions
  hRgnForm = CreateRectRgn(0, 0, Me.Width / 15, Me.Height / 15)
  hRgnTemp = CreatePolygonRgn(PointsRemove(1), 9, 1)
  
    ' Combining the two regions using XOR (invert the second region on the first)
  CombineRgn hRgnForm, hRgnForm, hRgnTemp, RGN_XOR
  
    ' Will now make a little circle within everything (will be visible thanks to XOR)
  hRgnTemp = CreateEllipticRgn(Me.Width * 0.25 / 15, Me.Height * 0.25 / 15, _
                                Me.Width * 0.75 / 15, Me.Height * 0.75 / 15)
  
  CombineRgn hRgnForm, hRgnForm, hRgnTemp, RGN_XOR
  
  SetWindowRgn Me.hwnd, hRgnForm, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  EasyMove Me
End Sub
