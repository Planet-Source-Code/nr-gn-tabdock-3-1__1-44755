VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSplitDDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ======================================================================
' Class    : cSplitDDC
' Filename : cSplitDC.cls
' Author   : SP McMahon
' Date     : 07 July 1998
'
' A splitter class using the Desktop window to draw a
' splitter bar, therefore allowing splitting of MDI forms
' as well as standard forms.
' ======================================================================
' REMARKS
' This code was extract from vbaccelarator web site and modified
' to acomplish the requirements of this project
'
' Thanks! to SP McMahon
'
' Modifications
' -------------
'
' -  The splitter object (and further calls to it) was replaced
'    by Top and Left properties once the Userobject TabDockHost
'    is already the bounds for splitting actions
'
' -  The events FormMouseMove and FormMouseUp was put together
'    in the FormMouseDown Event to supply a immediate split
'    drawing for TabDock UserControl
'
' -  The Clipping area (MouseDown() event) is now handled by
'    the TabDockHost MouseDown() event once this event calculates
'    the rectangle of the docked window
'
' -  Removed Cursor clipping (1.6)
'
' -  Removed Border references. This project does not need (1.6)
'
' -  Window Rect was replace for a Rect passed as reference
'    which is the exact area we want to split (1.6)
'
' -  Created an offset based on start and end cursor position
' removed declarations to the modAPi32 for reducing project size
'*******************
Option Explicit
'-- Splitter Ammended Calls
'-- By: Gary Noble (Phantom Man(PSC))
Private Const R2_NOTXORPEN        As Integer = 10     '  DPxn
Private Const PATINVERT           As Long = &H5A0049  ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT           As Long = &H550009  ' (DWORD) dest = (NOT dest)
Private Type BITMAP '24 bytes
    bmType                            As Long
    bmWidth                           As Long
    bmHeight                          As Long
    bmWidthBytes                      As Long
    bmPlanes                          As Integer
    bmBitsPixel                       As Integer
    bmBits                            As Long
End Type
Private m_lPattern(0 To 3)        As Long
Private m_hBrush                  As Long
'-- End
'// some global declarations
Private bDraw                     As Boolean
Private rcCurrent                 As RECT
Private rcNew                     As RECT
Private rcWindow                  As RECT
Private m_Offset                  As Long
Public Enum eOrientationConstants
    espVertical = 1
    espHorizontal = 2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private espVertical, espHorizontal
#End If
Private m_hWnd                    As Long
Private m_eOrientation            As eOrientationConstants
Private m_bSplitting              As Boolean
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                          ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                          lpRect As RECT, _
                          ByVal hBrush As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, _
                          ByVal x As Long, _
                          ByVal Y As Long, _
                          ByVal nWidth As Long, _
                          ByVal nHeight As Long, _
                          ByVal dwRop As Long) As Long

Private Sub Class_Initialize()

  Dim i As Long

    m_eOrientation = espVertical
    For i = 0 To 3
        m_lPattern(i) = &HAAAA5555
    Next i
    createBrush

End Sub

Private Function ClassName(ByVal lhWnd As Long) As String

  Dim lLen As Long
  Dim sBuf As String

    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lhWnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If

End Function

Private Function createBrush() As Boolean

  Dim tbm As BITMAP
  Dim hBm As Long

    DestroyBrush
    ' Create a monochrome bitmap containing the desired pattern:
    tbm.bmType = 0
    tbm.bmWidth = 16
    tbm.bmHeight = 8
    tbm.bmWidthBytes = 2
    tbm.bmPlanes = 1
    tbm.bmBitsPixel = 1
    tbm.bmBits = VarPtr(m_lPattern(0))
    hBm = CreateBitmapIndirect(tbm)
    ' Make a brush from the bitmap bits
    m_hBrush = CreatePatternBrush(hBm)
    '// Delete the useless bitmap
    DeleteObject hBm

End Function

Private Sub DestroyBrush()

    If Not (m_hBrush = 0) Then
        DeleteObject m_hBrush
        m_hBrush = 0
    End If

End Sub

Public Property Get Offset() As Long

  '-- end code

    Offset = m_Offset

End Property

Public Property Let Orientation(ByVal eOrientation As eOrientationConstants)

    m_eOrientation = eOrientation

End Property

Public Property Get Orientation() As eOrientationConstants

    Orientation = m_eOrientation

End Property

Private Sub pDrawSplitter(oRect As RECT)

  Dim lHDC      As Long
  Dim hOldBrush As Long

    '-- Ammended Splitter Calls
    lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    hOldBrush = SelectObject(lHDC, m_hBrush)
    PatBlt lHDC, oRect.Left, oRect.Top, oRect.Right - oRect.Left, oRect.Bottom - oRect.Top, PATINVERT
    SelectObject lHDC, hOldBrush
    DeleteDC lHDC

End Sub

Private Sub SplitterFormMouseMove(ByVal x As Long, _
                                  ByVal Y As Long)

  Dim hdc As Long
  Dim tP  As POINTAPI

    ' Changed to private
    'Dim hPen As Long
    If (bDraw) Then
        m_bSizing = True
        ' Draw two rectangles in the screen DC to cause splitting:
        ' First get the Desktop DC:
        hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        ' Set the draw mode to XOR:
        SetROP2 hdc, R2_NOTXORPEN
        'hPen = CreatePen(2, 2, vbButtonFace)
        'SelectObject hDC, hPen
        ' Draw over and erase the old rectangle
        ' (if this is the first time, all the coords will be 0 and nothing will get drawn):
        'Rectangle hdc, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
        pDrawSplitter rcCurrent
        ' It is simpler to use the mouse cursor position than try to translate
        ' X,Y to screen coordinates!
        GetCursorPos tP
        ' Determine where to draw the splitter:
        ''debug.print rcWindow.Top, rcWindow.Bottom, x, Y
        If (m_eOrientation = espHorizontal) Then
            rcNew.Left = rcWindow.Left
            rcNew.Right = rcWindow.Right
            If (tP.Y >= rcWindow.Top) And (tP.Y < rcWindow.Bottom) Then
                rcNew.Top = tP.Y - 2
                rcNew.Bottom = tP.Y + 2
              Else
                If (tP.Y < rcWindow.Top) Then
                    rcNew.Top = rcWindow.Top - 2
                    rcNew.Bottom = rcNew.Top + 2
                  Else
                    rcNew.Top = rcWindow.Bottom - 2
                    rcNew.Bottom = rcNew.Top + 2
                End If
            End If
          Else
            ''debug.print "drawing vertical splitter"
            rcNew.Top = rcWindow.Top
            rcNew.Bottom = rcWindow.Bottom
            If (tP.x >= rcWindow.Left) And (tP.x <= rcWindow.Right) Then
                rcNew.Left = tP.x - 2
                rcNew.Right = tP.x + 2
              Else
                If (tP.x < rcWindow.Left) Then
                    rcNew.Left = rcWindow.Left + 4
                    rcNew.Right = rcNew.Left + 2
                  Else
                    rcNew.Left = rcWindow.Right - 2
                    rcNew.Right = rcNew.Left + 2
                End If
            End If
        End If
        ' Draw the new rectangle
        ' Rectangle hdc, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom
        pDrawSplitter rcNew
        ' Store this position so we can erase it next time:
        LSet rcCurrent = rcNew
        ' Free the reference to the Desktop DC we got (make sure you do this!)
        DeleteDC hdc
        'DeleteObject hPen
    End If

End Sub

Private Function SplitterFormMouseUp(ByVal x As Long, _
                                     ByVal Y As Long) As Boolean

  Dim hdc As Long
  Dim tP  As POINTAPI

    ' changed to private
    ' Release mouse capture:
    ReleaseCapture
    ' Don't leave orphaned rectangle on desktop; erase last rectangle.
    If (bDraw) Then
        bDraw = False
        ' Get the Desktop DC:
        hdc = CreateDCAsNull("DISPLAY", 0, 0, 0)
        ' Set to XOR drawing mode:
        SetROP2 hdc, R2_NOTXORPEN
        ' Erase the last rectangle:
        'Rectangle hdc, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
        pDrawSplitter rcCurrent
        ' Clear up the desktop DC:
        DeleteDC hdc
        ' Here we ensure the splitter is within bounds before releasing:
        GetCursorPos tP
        If (tP.x < rcWindow.Left) Then
            tP.x = rcWindow.Left
        End If
        If (tP.x > rcWindow.Right) Then
            tP.x = rcWindow.Right
        End If
        If (tP.Y < rcWindow.Top) Then
            tP.Y = rcWindow.Top
        End If
        If (tP.Y > rcWindow.Bottom) Then
            tP.Y = rcWindow.Bottom
        End If
        ' Move the splitter to the validated final position:
        If (m_eOrientation = espHorizontal) Then
            m_Offset = (tP.Y - m_Offset) * Screen.TwipsPerPixelY
          Else
            m_Offset = (tP.x - m_Offset) * Screen.TwipsPerPixelX
        End If
        ' Return true to tell the owner we have completed splitting:
        SplitterFormMouseUp = True
        m_bSizing = False
    End If

End Function

Friend Sub SplitterMouseDown(ByVal hwnd As Long, _
       RC As RECT, _
       ByVal x As Long, _
       ByVal Y As Long)

  Dim tP     As POINTAPI
  Dim tpPrev As POINTAPI

    m_hWnd = hwnd
    ' Send subsequent mouse messages to the owner window
    SetCapture m_hWnd
    ' get window rect
    rcWindow = RC
    'rcWindow.Left = rcWindow.Left + 23
    ' Get the Desktop DC:
    '    hDC = CreateDCAsNull("DISPLAY", 0, 0, 0)
    ' Set to XOR drawing mode:
    '    SetROP2 hDC, R2_NOTXORPEN
    'hPen = CreatePen(2, 1, vbButtonFace)
    'SelectObject hDC, hPen
    ' Erase the last rectangle:
    '    Rectangle hDC, Rc.Left, Rc.Top, Rc.Right, Rc.Bottom
    ' Clear up the desktop DC:
    '    DeleteDC hDC
    'DeleteObject hPen
    GetCursorPos tP
    ' Store the initial cursor position
    tpPrev.x = tP.x
    tpPrev.Y = tP.Y
    If (m_eOrientation = espHorizontal) Then
        m_Offset = tP.Y
      Else
        m_Offset = tP.x
    End If
    bDraw = True  ' start actual drawing from next move message
    rcCurrent.Left = 0
    rcCurrent.Top = 0
    rcCurrent.Right = 0
    rcCurrent.Bottom = 0
    SplitterFormMouseMove tP.x, tP.Y
    Do While GetKeyState(VK_LBUTTON) < 0
        GetCursorPos tP
        If tP.x <> tpPrev.x Or tP.Y <> tpPrev.Y Then
            tpPrev.x = tP.x
            tpPrev.Y = tP.Y
            SplitterFormMouseMove tP.x, tP.Y
            m_bSizing = True
        End If
        DoEvents
    Loop
    SplitterFormMouseUp tP.x, tP.Y

End Sub

