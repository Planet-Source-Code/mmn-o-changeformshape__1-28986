VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Form_ControlPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const c_PosMode = "PosMode"
Private Const c_Top = "Top"
Private Const c_Left = "Left"
Private Const c_Width = "Width"
Private Const c_Height = "Height"
Private Const c_FontSize = "FontSize"
Private Const c_MoveTop = "MoveTop"
Private Const c_MoveLeft = "MoveLeft"
Private Const c_MoveWidth = "MoveWidth"
Private Const c_MoveHeight = "MoveHeight"
Private Const c_MoveFontSize = "MoveFontSize"

Private WithEvents m_Form As Form
Attribute m_Form.VB_VarHelpID = -1

Private m_Width As Long
Private m_Height As Long
Private m_WindowState As FormWindowStateConstants

Private m_Col As Collection

Public Enum FormControlPosModes
    FormControl_None = 0
    FormControl_Absolute = 1
    FormControl_Proportional = 2
    FormControl_Half = 3
End Enum

Public Function Add(ctl As Control, _
                    Optional ByVal Left As FormControlPosModes = FormControl_None, _
                    Optional ByVal Top As FormControlPosModes = FormControl_None, _
                    Optional ByVal Width As FormControlPosModes = FormControl_None, _
                    Optional ByVal Height As FormControlPosModes = FormControl_None, _
                    Optional ByVal FontSize As Boolean = False) As Long

Dim lngReturn As Long
Dim mCtl As Form_Control

    On Error Resume Next

    Set mCtl = New Form_Control
    With mCtl
        Set .Control = ctl
        With .Properties
            
            .Add ctl.Left, c_Left
            .Add ctl.Top, c_Top
            .Add ctl.Width, c_Width
            .Add ctl.Height, c_Height
            .Add ctl.FontSize, c_FontSize
            
            .Add Left, c_MoveLeft
            .Add Top, c_MoveTop
            .Add Height, c_MoveHeight
            .Add Width, c_MoveWidth
            .Add FontSize, c_MoveFontSize
            
        End With
    
    End With
    
    m_Col.Add mCtl

End Function


Public Function AddProportional(ctl As Control) As Long

    AddProportional = Me.Add(ctl, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, FormControl_Proportional, True)

End Function

Public Function Init(frm As Form) As Long

    Set m_Col = New Collection

    Set m_Form = frm
    With m_Form
        m_Width = .ScaleWidth
        m_Height = .ScaleHeight
        m_WindowState = .WindowState
    End With
    
End Function


Private Sub m_Form_Resize()

Dim mCtl As Form_Control
Dim lngW As Long    ' Norm-Mode
Dim lngH As Long    ' Norm-Mode
Dim lngWH As Long   ' Half-Mode
Dim lngHH As Long   ' Half-Mode
Dim dblWP As Double ' Prop-Mode
Dim dblHP As Double ' Prop-Mode
Dim lMode As FormControlPosModes
    
    On Error Resume Next

    With m_Form
        lngW = .ScaleWidth - m_Width
        lngH = .ScaleHeight - m_Height
        lngHH = lngH / 2
        lngWH = lngW / 2
        dblWP = .ScaleWidth / m_Width
        dblHP = .ScaleHeight / m_Height
    End With

    ' Alle gemeldeten Control durchgehen
    For Each mCtl In m_Col
        With mCtl
            ' Links
            Select Case .Properties(c_MoveLeft)
                Case FormControl_Absolute
                    .Control.Left = .Properties(c_Left) + lngW
                Case FormControl_Proportional
                    .Control.Left = .Properties(c_Left) * dblWP
                Case FormControl_Half
                    .Control.Left = .Properties(c_Left) + lngWH
            End Select

            ' Oben
            Select Case .Properties(c_MoveTop)
                Case FormControl_Absolute
                    .Control.Top = .Properties(c_Top) + lngH
                Case FormControl_Proportional
                    .Control.Top = .Properties(c_Top) * dblHP
                Case FormControl_Half
                    .Control.Top = .Properties(c_Top) + lngHH
            End Select

            ' Breite
            Select Case .Properties(c_MoveWidth)
                Case FormControl_Absolute
                    .Control.Width = .Properties(c_Width) + lngW
                Case FormControl_Proportional
                    .Control.Width = .Properties(c_Width) * dblWP
                Case FormControl_Half
                    .Control.Width = .Properties(c_Width) + lngWH
            End Select

            ' Höhe
            Select Case .Properties(c_MoveHeight)
                Case FormControl_Absolute
                    .Control.Height = .Properties(c_Height) + lngH
                Case FormControl_Proportional
                    .Control.Height = .Properties(c_Height) * dblHP
                Case FormControl_Half
                    .Control.Height = .Properties(c_Height) + lngHH
            End Select
                
            ' Fontsize immer Proportional
            If .Properties(c_MoveFontSize) Then
                If .Properties(c_MoveHeight) <> FormControl_None Then
                    .Control.FontSize = .Properties(c_FontSize) * dblHP
                End If
            End If
        End With
    Next

    ' Fehler in VB, beim zurücksetzen auf vbNormal wird Ereignis falsch ausgelöst
    If m_WindowState <> m_Form.WindowState Then
        m_WindowState = m_Form.WindowState
        If m_WindowState = vbNormal Then
            Call m_Form_Resize
        End If
    End If

End Sub

