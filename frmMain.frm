VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "ChangeFormShape"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPictureShaped 
      Caption         =   "Show Picture Shaped"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2220
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdPicture 
      Caption         =   "Show Picture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1140
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdPolygon 
      Caption         =   "Show Polygon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   14
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox chkShowResizeBars 
      Caption         =   "Show Resize Bars"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2460
      Width           =   1755
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   2460
      Width           =   555
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "&Set"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2400
      TabIndex        =   11
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtBend 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   600
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "50"
      ToolTipText     =   "Bend"
      Top             =   2100
      Width           =   615
   End
   Begin VB.OptionButton optRoundRect 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rounded Rectangle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Value           =   -1  'True
      Width           =   2040
   End
   Begin VB.OptionButton optRect 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rectangle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1500
      Width           =   1260
   End
   Begin VB.OptionButton optElliptic 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Elliptic"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   900
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picTop 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   3375
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton cmdClose 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Label lblDownRight 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      MousePointer    =   8  'Size NW SE
      TabIndex        =   4
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblRight 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   3240
      MousePointer    =   9  'Size W E
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblDown 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   5
      Top             =   2880
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LeftX As Single
Dim TopY As Single

Dim FormShape As MNShape
Dim intBend As Integer

Private mfAngle         As Single
Private mlColor1        As Long
Private mlColor2        As Long
Private mGradient       As New clsGradient

Private mFormPos As Form_ControlPos

Private Sub chkShowResizeBars_Click()
  lblDown.BackStyle = chkShowResizeBars.Value
  lblDownRight.BackStyle = chkShowResizeBars.Value
  lblRight.BackStyle = chkShowResizeBars.Value
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdMin_Click()
  frmMain.WindowState = vbMinimized
End Sub

Private Sub cmdPicture_Click()
  frmPicture.Show
End Sub

Private Sub cmdPictureShaped_Click()
  frmPictureShape.Show
End Sub

Private Sub cmdPolygon_Click()
  frmPolygon.Show
End Sub

Private Sub cmdRestore_Click()
  If frmMain.WindowState = vbMaximized Then
    frmMain.WindowState = vbNormal
    cmdRestore.Caption = "1"
    lblDown.Enabled = True
    lblDownRight.Enabled = True
    lblRight.Enabled = True
    Form_Resize
  Else
    frmMain.WindowState = vbMaximized
    cmdRestore.Caption = "2"
    lblDown.Enabled = False
    lblDownRight.Enabled = False
    lblRight.Enabled = False
    Form_Resize
  End If
End Sub

Private Sub cmdSet_Click()
  If optElliptic Then
    SetWindowRgn Me.hwnd, CreateEllipticRgn(0, 0, _
                  (Me.Width / 15), (Me.Height / 15)), _
                  True
    FormShape = Elliptic
  ElseIf optRect Then
  SetWindowRgn Me.hwnd, CreateRectRgn(0, 0, _
                (Me.Width / 15), (Me.Height / 15)), _
                True
    FormShape = Rectangle
  ElseIf optRoundRect Then
    intBend = Abs(Val(txtBend))     ' Abs() turns the number to positive
  SetWindowRgn Me.hwnd, CreateRoundRectRgn(0, 0, _
                (Me.Width / 15), (Me.Height / 15), _
                intBend, intBend), _
                True
    FormShape = RoundedRectangle
  End If
End Sub

Private Sub Form_Load()
  Set mFormPos = New Form_ControlPos
  
  lblDown.BackStyle = chkShowResizeBars.Value
  lblDownRight.BackStyle = chkShowResizeBars.Value
  lblRight.BackStyle = chkShowResizeBars.Value
  
  picTop.FontName = "Arial"
  picTop.Print frmMain.Caption
  
  FormShape = RoundedRectangle
  intBend = 50
  
  cmdSet_Click
  
  With mFormPos
    Call .Init(Me)
    Call .Add(picTop, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(cmdMin, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(cmdRestore, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(cmdClose, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(cmdPolygon, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(cmdPicture, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(cmdPictureShaped, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(optElliptic, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(optRect, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(optRoundRect, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(txtBend, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(chkShowResizeBars, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(cmdSet, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
    Call .Add(cmdExit, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)
  End With
  
  With mGradient
    .Angle = 90
    .Color1 = RGB(0, 0, 127)
    .Color2 = RGB(0, 0, 255)
    .Draw picTop    ' Also handled in Form_Resize
  End With
End Sub

Private Sub Form_Resize()
  If FormShape = Elliptic Then
    SetWindowRgn Me.hwnd, CreateEllipticRgn(0, 0, _
                  (Me.Width / 15), (Me.Height / 15)), _
                  True
  ElseIf FormShape = Rectangle Then
    SetWindowRgn Me.hwnd, CreateRectRgn(0, 0, _
                  (Me.Width / 15), (Me.Height / 15)), _
                  True
  ElseIf FormShape = RoundedRectangle Then
    SetWindowRgn Me.hwnd, CreateRoundRectRgn(0, 0, _
                  (Me.Width / 15), (Me.Height / 15), _
                  intBend, intBend), _
                True
  End If
  mGradient.Draw picTop
  picTop.CurrentX = 10 * 15
  picTop.CurrentY = picTop.Height / 2 - picTop.FontSize * 15 / 2
  picTop.Print frmMain.Caption
  picTop.Refresh          ' It must have the AutoRedraw property True (picTop)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub lblDownRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    LeftX = x
    TopY = y
  End If
End Sub

Private Sub lblDownRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If Button = vbLeftButton Then
    lblDownRight.Left = lblDownRight.Left + x - LeftX
    lblDownRight.Top = lblDownRight.Top + y - TopY
    Me.Width = lblDownRight.Left + lblDownRight.Width
    Me.Height = lblDownRight.Top + lblDownRight.Height
    lblRight.Left = lblRight.Left + x - LeftX
    lblRight.Height = Me.Height
    lblDown.Top = Me.Height - lblDown.Height
    lblDown.Width = Me.Width
    cmdSet_Click
  End If
End Sub

Private Sub lblRight_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    LeftX = x
    TopY = y
  End If
End Sub

Private Sub lblRight_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If Button = vbLeftButton Then
    lblRight.Left = lblRight.Left + x - LeftX
    lblRight.Height = Me.Height
    Me.Width = lblRight.Left + lblRight.Width
    lblDownRight.Top = Me.Height - lblDownRight.Height
    lblDownRight.Left = Me.Width - lblDownRight.Width
    lblDown.Top = Me.Height - lblDown.Height
    lblDown.Width = Me.Width
    cmdSet_Click
  End If
End Sub

Private Sub lblDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    LeftX = x
    TopY = y
  End If
End Sub

Private Sub lblDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  On Error Resume Next
  If Button = vbLeftButton Then
    lblDown.Top = lblDown.Top + y - TopY
    Me.Height = lblDown.Top + lblDown.Height
    lblRight.Height = Me.Height
    lblDownRight.Left = Me.Width - lblDownRight.Width
    lblDownRight.Top = Me.Height - lblDownRight.Height
    cmdSet_Click
  End If
End Sub

Private Sub optElliptic_Click()
  txtBend.Enabled = False
End Sub

Private Sub optRect_Click()
  txtBend.Enabled = False
End Sub

Private Sub optRoundRect_Click()
  txtBend.Enabled = True
End Sub

Private Sub picTop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  EasyMove Me
End Sub
