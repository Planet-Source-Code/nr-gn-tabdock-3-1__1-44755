Attribute VB_Name = "modBorder"
'
'  Color your Border
'  Code originally intended to produce 3D borders
'       for WIN 3.x Forms.
'  Reapplied by: linda
'  linda.69@mailcity.com
'
'  can be easily modified to draw anything on the
'  title bar of the form.
'
Option Explicit
'DefInt A-Z

' private declares
' Pen Styles
Private Const PS_SOLID         As Integer = 0
Private Const CLR_INVALID      As Integer = 0
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                          ByVal hdc As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                        ByVal nWidth As Long, _
                        ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                        ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                          ByVal x As Long, _
                          ByVal Y As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, _
                          ByVal hpal As Long, _
                          ByRef lpcolorref As Long) As Long
' ******************************************************************************
' Routine       : DrawBorder
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 02/10/0010:20:55
' Inputs        :
' Outputs       :
' Credits       : linda.69@mailcity.com (Color your Border demo)
' Modifications : Color translation OLE_COLOR to RGB
' Description   : draw a user defined color border
' ******************************************************************************
Private DummyToKeepDecCommentsInDEclarations As Boolean

Public Sub DrawBorder(frmTarget As Form, _
                      Color As OLE_COLOR)

  Dim hWindowDC As Long
  Dim hOldPen   As Long
  Dim nLeft     As Long
  Dim nRight    As Long
  Dim nTop      As Long
  Dim nBottom   As Long
  Dim Ret       As Long
  Dim hMyPen    As Long
  Dim WidthX    As Long
  Dim rgbColor  As Long

    ' translate
    rgbColor = TranslateColor(Color)
    ' border width
    WidthX = GetSystemMetrics(SM_CYBORDER) * 5
    ' get window DC
    hWindowDC = GetWindowDC(frmTarget.hwnd)   'this is outside the form
    ' create a pen
    hMyPen = CreatePen(PS_SOLID, WidthX, rgbColor)
    ' Initialize misc variables
    nLeft = 0
    nTop = 0
    nRight = frmTarget.width / Screen.TwipsPerPixelX
    nBottom = frmTarget.Height / Screen.TwipsPerPixelY
    ' select border pen
    hOldPen = SelectObject(hWindowDC, hMyPen)
    ' draw color around the border
    Ret = LineTo(hWindowDC, nLeft, nBottom)
    Ret = LineTo(hWindowDC, nRight, nBottom)
    Ret = LineTo(hWindowDC, nRight, nTop)
    Ret = LineTo(hWindowDC, nLeft, nTop)
    ' select old pen
    Ret = SelectObject(hWindowDC, hOldPen)
    Ret = DeleteObject(hMyPen)
    Ret = ReleaseDC(frmTarget.hwnd, hWindowDC)

End Sub

Public Sub DrawBorderUnDocked(frmTarget As Form, _
                              Color As OLE_COLOR)

  Dim hWindowDC As Long
  Dim hOldPen   As Long
  Dim nLeft     As Long
  Dim nRight    As Long
  Dim nTop      As Long
  Dim nBottom   As Long
  Dim Ret       As Long
  Dim hMyPen    As Long
  Dim WidthX    As Long
  Dim rgbColor  As Long

    ' translate
    rgbColor = TranslateColor(Color)
    ' border width
    WidthX = GetSystemMetrics(SM_CYBORDER) * 5
    ' get window DC
    hWindowDC = GetWindowDC(frmTarget.hwnd)   'this is outside the form
    ' create a pen
    hMyPen = CreatePen(PS_SOLID, WidthX, rgbColor)
    ' Initialize misc variables
    nLeft = 0
    nTop = -1
    nRight = (frmTarget.width / Screen.TwipsPerPixelX) - 1
    nBottom = (frmTarget.Height / Screen.TwipsPerPixelY) - 1
    ' select border pen
    hOldPen = SelectObject(hWindowDC, hMyPen)
    ' draw color around the border
    Ret = LineTo(hWindowDC, nLeft, nBottom)
    Ret = LineTo(hWindowDC, nRight, nBottom)
    Ret = LineTo(hWindowDC, nRight, nTop)
    Ret = LineTo(hWindowDC, nLeft, nTop)
    ' select old pen
    Ret = SelectObject(hWindowDC, hOldPen)
    Ret = DeleteObject(hMyPen)
    Ret = ReleaseDC(frmTarget.hwnd, hWindowDC)

End Sub

Public Function TranslateColor(ByVal clr As OLE_COLOR, _
                               Optional hpal As Long = 0) As Long

  ' ******************************************************************************
  ' Routine       : TranslateColor
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 02/10/0010:20:19
  ' Inputs        :
  ' Outputs       :
  ' Credits       : Extracted from VB KB Article
  ' Modifications :
  ' Description   : Converts an OLE_COLOR to RGB color
  ' ******************************************************************************

    If OleTranslateColor(clr, hpal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If

End Function

