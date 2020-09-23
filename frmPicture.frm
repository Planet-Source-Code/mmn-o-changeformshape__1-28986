VERSION 5.00
Begin VB.Form frmPicture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture-shape"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   2175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShape 
      Caption         =   "Elliptic"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   2760
      Width           =   675
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1500
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.PictureBox picTest 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2670
      Left            =   0
      Picture         =   "frmPicture.frx":0000
      ScaleHeight     =   2670
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdShape_Click()
  If cmdShape.Caption = "Elliptic" Then
    SetWindowRgn picTest.hwnd, _
                  CreateEllipticRgn(0, _
                                    0, _
                                    picTest.Width / 15, _
                                    picTest.Height / 15), _
                  True
    cmdShape.Caption = "RoundRect"
  Else
    SetWindowRgn picTest.hwnd, _
                  CreateRoundRectRgn(0, 0, _
                                      picTest.Width / 15, picTest.Height / 15, _
                                      50, 50), _
                  True
    cmdShape.Caption = "Elliptic"
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    ' The picturebox should have the BorderStyle set to '0 - None'
    SetWindowRgn picTest.hwnd, _
                  CreateRoundRectRgn(0, 0, _
                                      picTest.Width / 15, picTest.Height / 15, _
                                      50, 50), _
                  True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Show
End Sub
