VERSION 5.00
Begin VB.Form frmPictureShape 
   BorderStyle     =   0  'None
   Caption         =   "PictureShapedForm"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPictureShape.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPictureShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ShapeForm As clsTransForm 'make a reference to the class

Private Sub Form_DblClick()
  Unload Me
End Sub

Private Sub Form_Load()
  Set ShapeForm = New clsTransForm 'instantiate the object from the class
  
    ' If the picture-shaped form is not showing correctly, remove 'FormShape.shp' _
      from the application's folder...
  
    ' I'm having trouble when trying to make something with green transparent, _
      but apparently I'm the only one that does...
  If Dir(App.Path & "\FormShape.shp") = "" Then   ' Save the FormRegion
    ShapeForm.ShapeMe Me, RGB(0, 0, 255), False, App.Path & "\FormShape.shp"
  Else  ' Only use below if you wish to recalculate everything...
    ShapeForm.ShapeMe Me, RGB(0, 0, 255), True, App.Path & "\FormShape.shp"
  End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  EasyMove Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set ShapeForm = Nothing 'destroy the object
  frmMain.Show
End Sub
