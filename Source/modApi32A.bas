Attribute VB_Name = "modApi32"
' ******************************************************************************
' Module      : modApi32.bas
' Created by  : Marclei V Silva
' Machine     : ZEUS
' Date-Time   : 09/05/20003:09:33
' Description : Several Api declares, constants and definitions
' ******************************************************************************
Option Explicit
Public GradClr1                           As OLE_COLOR
Public GradClr2                           As OLE_COLOR
Public Type RECT
    Left                                      As Long
    Top                                       As Long
    Right                                     As Long
    Bottom                                    As Long
End Type
Public Type POINTAPI
    x                                         As Long
    Y                                         As Long
End Type
Public Persist                            As Boolean
Public m_bSizing                          As Boolean
' System metrics constants
Public Const SM_CXMIN                     As Integer = 28
Public Const SM_CYMIN                     As Integer = 29
Public Const SM_CXSIZE                    As Integer = 30
Public Const SM_CXFRAME                   As Integer = 32
Public Const SM_CYFRAME                   As Integer = 33
Public Const SM_CYSIZE                    As Integer = 31
Public Const SM_CYCAPTION                 As Integer = 4
Public Const SM_CXBORDER                  As Integer = 5
Public Const SM_CYBORDER                  As Integer = 6
Public Const SM_CYMENU                    As Integer = 15
Public Const SM_CYSMCAPTION               As Integer = 51     'height of windows 95 small caption
' These constants define the style of border to draw.
Public Const BDR_RAISED                   As Long = &H5
Public Const BDR_RAISEDINNER              As Long = &H4
Public Const BDR_RAISEDOUTER              As Long = &H1
Public Const BDR_SUNKEN                   As Long = &HA
Public Const BDR_SUNKENINNER              As Long = &H8
Public Const BDR_SUNKENOUTER              As Long = &H2
Public Const BF_FLAT                      As Long = &H4000
Public Const BF_MONO                      As Long = &H8000
Public Const BF_SOFT                      As Long = &H1000    ' For softer buttons
Public Const EDGE_BUMP                    As Double = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_ETCHED                  As Double = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Public Const EDGE_RAISED                  As Double = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN                  As Double = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
' These constants define which sides to draw.
Public Const BF_BOTTOM                    As Long = &H8
Public Const BF_LEFT                      As Long = &H1
Public Const BF_RIGHT                     As Long = &H4
Public Const BF_TOP                       As Long = &H2
Public Const BF_RECT                      As Double = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const SWP_NOOWNERZORDER            As Long = &H200     ' Don"t do owner Z ordering
Public Const SWP_FRAMECHANGED             As Long = &H20
Public Const SWP_NOREPOSITION             As Long = SWP_NOOWNERZORDER
Public Const SWP_NOZORDER                 As Long = &H4
Public Const SWP_NOACTIVATE               As Long = &H10
Public Const SWP_SHOWWINDOW               As Long = &H40
Public Const SWP_NOSIZE                   As Long = &H1
Public Const SWP_NOMOVE                   As Long = &H2
Public Const TOPMOST_FLAGS                As Double = SWP_NOMOVE Or SWP_NOSIZE
Public Const GWL_STYLE                    As Long = (-16)
Public Const GWL_EXSTYLE                  As Long = (-20)
Public Const GWL_HWNDPARENT               As Long = (-8)
Public Const SW_SHOW                      As Integer = 5
Public Const SW_HIDE                      As Integer = 0
Public Const SW_SHOWNORMAL                As Integer = 1
' Window styles
Public Const WS_ACTIVECAPTION             As Long = &H1
Public Const WS_BORDER                    As Long = &H800000
Public Const WS_CAPTION                   As Long = &HC00000  'WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD                     As Long = &H40000000
Public Const WS_CHILDWINDOW               As Long = (WS_CHILD)
Public Const WS_CLIPCHILDREN              As Long = &H2000000
Public Const WS_CLIPSIBLINGS              As Long = &H4000000
Public Const WS_DISABLED                  As Long = &H8000000
Public Const WS_DLGFRAME                  As Long = &H400000
Public Const WS_GROUP                     As Long = &H20000
Public Const WS_TABSTOP                   As Long = &H10000
Public Const WS_GT                        As Double = WS_GROUP Or WS_TABSTOP
Public Const WS_HSCROLL                   As Long = &H100000
Public Const WS_MAXIMIZE                  As Long = &H1000000
Public Const WS_MINIMIZE                  As Long = &H20000000
Public Const WS_ICONIC                    As Long = WS_MINIMIZE
Public Const WS_MAXIMIZEBOX               As Long = &H10000
Public Const WS_MINIMIZEBOX               As Long = &H20000
Public Const WS_OVERLAPPED                As Long = &H0&
Public Const WS_SYSMENU                   As Long = &H80000
Public Const WS_THICKFRAME                As Long = &H40000
Public Const WS_OVERLAPPEDWINDOW          As Double = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
Public Const WS_POPUP                     As Long = &H80000000
Public Const WS_POPUPWINDOW               As Double = WS_POPUP Or WS_BORDER Or WS_SYSMENU
Public Const WS_SIZEBOX                   As Long = WS_THICKFRAME
Public Const WS_TILED                     As Long = WS_OVERLAPPED
Public Const WS_TILEDWINDOW               As Long = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE                   As Long = &H10000000
Public Const WS_VSCROLL                   As Long = &H200000
Public Const WM_NCRBUTTONUP               As Long = &HA5
Public Const WM_RBUTTONUP                 As Long = &H205
' Extended window styles
Public Const WS_EX_ACCEPTFILES            As Long = &H10&
Public Const WS_EX_APPWINDOW              As Long = &H40000
Public Const WS_EX_CLIENTEDGE             As Long = &H200&
Public Const WS_EX_CONTEXTHELP            As Long = &H400&
Public Const WS_EX_CONTROLPARENT          As Long = &H10000
Public Const WS_EX_DLGMODALFRAME          As Long = &H1&
Public Const WS_EX_LAYERED                As Long = &H80000
Public Const WS_EX_LAYOUTRTL              As Long = &H400000  ' Right to left mirroring
Public Const WS_EX_LEFT                   As Long = &H0&
Public Const WS_EX_LEFTSCROLLBAR          As Long = &H4000&
Public Const WS_EX_LTRREADING             As Long = &H0&
Public Const WS_EX_MDICHILD               As Long = &H40&
Public Const WS_EX_NOACTIVATE             As Long = &H8000000
Public Const WS_EX_NOINHERITLAYOUT        As Long = &H100000  ' Disable inheritence of mirroring by children
Public Const WS_EX_NOPARENTNOTIFY         As Long = &H4&
Public Const WS_EX_RIGHT                  As Long = &H1000&
Public Const WS_EX_RIGHTSCROLLBAR         As Long = &H0&
Public Const WS_EX_RTLREADING             As Long = &H2000&
Public Const WS_EX_STATICEDGE             As Long = &H20000
Public Const WS_EX_TOOLWINDOW             As Long = &H80&
Public Const WS_EX_TOPMOST                As Long = &H8&
Public Const WS_EX_TRANSPARENT            As Long = &H20&
Public Const WS_EX_WINDOWEDGE             As Long = &H100&
Public Const WS_EX_OVERLAPPEDWINDOW       As Double = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
Public Const WS_EX_PALETTEWINDOW          As Double = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
Public Const SC_CLOSE                     As Long = &HF060&
Public Const SC_MOVE                      As Long = &HF010&
Public Const SC_SIZE                      As Long = &HF000&
Public Const OPAQUE                       As Integer = 2
Public Const VK_LBUTTON                   As Long = &H1
Public Const PS_SOLID                     As Integer = 0
Public Const BLACK_PEN                    As Integer = 7
Public Const MOUSE_MOVE                   As Long = &HF012
Public Const TRANSPARENT                  As Integer = 1
Public Const BITSPIXEL                    As Integer = 12
' subclassing constants
Public Const WM_NCACTIVATE                As Long = &H86
Public Const WM_ACTIVATEAPP               As Long = &H1C
Public Const WM_NCLBUTTONDBLCLK           As Long = &HA3
Public Const WM_NCLBUTTONDOWN             As Long = &HA1
Public Const WM_NCRBUTTONDOWN             As Long = &HA4
Public Const WM_MOVE                      As Long = &H3
Public Const WM_EXITSIZEMOVE              As Long = &H232
Public Const WM_SIZE                      As Long = &H5
Public Const WM_USER                      As Long = &H400
Public Const WM_MOUSEMOVE                 As Long = &H200
Public Const WM_NCMOUSEMOVE               As Long = &HA0
Public Const WM_WININICHANGE              As Long = &H1A
Public Const HWND_BROADCAST               As Long = &HFFFF
Public Const WM_LBUTTONDOWN               As Long = &H201
Public Const WM_LBUTTONUP                 As Long = &H202
Public Const WM_SYSCOMMAND                As Long = &H112
Public Const WM_NULL                      As Long = &H0
Public Const WM_MOUSEACTIVATE             As Long = &H21
Public Const WM_WINDOWPOSCHANGING         As Long = &H46
Public Const WM_ACTIVATE                  As Long = &H6
Public Const WM_KILLFOCUS                 As Long = &H8
Public Const WM_PAINT                     As Long = &HF
Public Const WM_DESTROY                   As Long = &H2
Public Const WM_NCHITTEST                 As Long = &H84
Public Const WM_MDIMAXIMIZE               As Long = &H225
Public Const WM_COMMAND                   As Long = &H111
Public Const WM_CONTEXTMENU               As Long = &H7B
Public Const WM_ENTERMENULOOP             As Long = &H211
Public Const WM_EXITMENULOOP              As Long = &H212
Public Const WM_STYLECHANGED              As Long = &H7D&
'Public Const WM_DESTROY = &H2
Public Const WM_SIZING                    As Long = &H214
Public Const WM_MOVING                    As Long = &H216&
Public Const WM_ENTERSIZEMOVE             As Long = &H231&
'Public Const WM_EXITSIZEMOVE = &H232&
'Public Const WM_ACTIVATE = &H6
'Public Const WM_SIZE = &H5
Public Const WM_CLOSE                     As Long = &H10
Public Const WM_NCPAINT                   As Long = &H85

Public Const PS_INSIDEFRAME               As Integer = 6
' Region constants
Public Const RGN_OR                       As Integer = 2      ' RGN_OR creates the union of combined regions
Public Const RGN_DIFF                     As Integer = 4      ' RGN_DIFF creates the intersection of combined regions
Public Const RGN_AND                      As Integer = 1
Public Const RGN_XOR                      As Integer = 3
' SysCommand
'Public Const HTCAPTION = 2
Public Const HTCLOSE                      As Integer = 20
Public Const R2_BLACK                     As Integer = 1      '   0
Public Const R2_COPYPEN                   As Integer = 13     '  P
Public Const R2_LAST                      As Integer = 16
Public Const R2_MASKNOTPEN                As Integer = 3      '  DPna
Public Const R2_MASKPEN                   As Integer = 9      '  DPa
Public Const R2_MASKPENNOT                As Integer = 5      '  PDna
Public Const R2_MERGENOTPEN               As Integer = 12     '  DPno
Public Const R2_MERGEPEN                  As Integer = 15     '  DPo
Public Const R2_MERGEPENNOT               As Integer = 14     '  PDno
Public Const R2_NOP                       As Integer = 11     '  D
Public Const R2_NOT                       As Integer = 6      '  Dn
Public Const R2_NOTCOPYPEN                As Integer = 4      '  PN
Public Const R2_NOTMASKPEN                As Integer = 8      '  DPan
Public Const R2_NOTMERGEPEN               As Integer = 2      '  DPon
Public Const R2_NOTXORPEN                 As Integer = 10     '  DPxn
Public Const R2_WHITE                     As Integer = 16     '   1
Public Const R2_XORPEN                    As Integer = 7      '  DPx
Public Type BITMAP '24 bytes
    bmType                                    As Long
    bmWidth                                   As Long
    bmHeight                                  As Long
    bmWidthBytes                              As Long
    bmPlanes                                  As Integer
    bmBitsPixel                               As Integer
    bmBits                                    As Long
End Type
Public m_lPattern(0 To 3)                 As Long
Public Const PATINVERT                    As Long = &H5A0049  ' (DWORD) dest = pattern XOR dest
Public Const DSTINVERT                    As Long = &H550009  ' (DWORD) dest = (NOT dest)
'*********************************
Public Const DFC_CAPTION                  As Integer = 1
Public Const DFC_MENU                     As Integer = 2      'Menu
Public Const DFC_SCROLL                   As Integer = 3      'Scroll bar
Public Const DFC_BUTTON                   As Integer = 4      'Standard button
Public Const DFCS_CAPTIONCLOSE            As Long = &H0
Public Const DFCS_CAPTIONRESTORE          As Long = &H3
Public Const DFCS_FLAT                    As Long = &H4000
Public Const DFCS_PUSHED                  As Long = &H200
Public Const DFCS_INACTIVE                As Long = &H100
Public Const DFCS_MENUARROWRIGHT          As Long = &H4
Public Const DFCS_SCROLLUP                As Long = &H0
Public Const DFCS_SCROLLLEFT              As Long = &H2
Public Const DI_NORMAL                    As Long = &H3

Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                        lpDeviceName As Any, _
                        lpOutput As Any, _
                        lpInitData As Any) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                        lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
                        lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, _
                        ByVal Y As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                        lpRect As RECT) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Public Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                        ByVal lLeft As Long, _
                        ByVal lTop As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                        ByVal wMsg As Long, _
                        ByVal wParam As Long, _
                        lParam As Any) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
                        ByVal Y1 As Long, _
                        ByVal X2 As Long, _
                        ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                        ByVal hSrcRgn1 As Long, _
                        ByVal hSrcRgn2 As Long, _
                        ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                        ByVal hRgn As Long, _
                        ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, _
                        ByVal X1 As Long, _
                        ByVal Y1 As Long, _
                        ByVal X2 As Long, _
                        ByVal Y2 As Long) As Long
Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, _
                        ByVal nDrawMode As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                        ByVal hObject As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                        ByVal nWidth As Long, _
                        ByVal crColor As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
                        ByVal x As Long, _
                        ByVal Y As Long, _
                        ByVal nWidth As Long, _
                        ByVal nHeight As Long, _
                        ByVal bRepaint As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                        ByVal hdc As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                        lpRect As Any) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                        ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                        ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                        ByVal nIndex As Long, _
                        ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                        ByVal hWndInsertAfter As Long, _
                        ByVal x As Long, _
                        ByVal Y As Long, _
                        ByVal cx As Long, _
                        ByVal cy As Long, _
                        ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                        ByVal nIndex As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                        pSrc As Any, _
                        ByVal ByteLen As Long)
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
                        qrc As RECT, _
                        ByVal Edge As Long, _
                        ByVal grfFlags As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                        ByVal lpClassName As String, _
                        ByVal nMaxCount As Long) As Long
Public Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, _
                        ByVal x As Long, _
                        ByVal Y As Long, _
                        ByVal nWidth As Long, _
                        ByVal nHeight As Long, _
                        ByVal dwRop As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                        ByVal hWnd2 As Long, _
                        ByVal lpsz1 As String, _
                        lpsz2 As Any) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                        ByVal x As Long, _
                        ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                        ByVal x As Long, _
                        ByVal Y As Long, _
                        lpPoint As POINTAPI) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                        lpRect As RECT, _
                        ByVal hBrush As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, _
                        lpRect As RECT, _
                        ByVal un1 As Long, _
                        ByVal un2 As Long) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                        ByVal x As Long, _
                        ByVal Y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'***********************************
Private DummyToKeepDecCommentsInDEclarations As Boolean

Public Function BlendColor(ByVal oColorFrom As OLE_COLOR, _
                           ByVal oColorTo As OLE_COLOR, _
                           Optional ByVal alpha As Long = 128) As Long

  Dim lCFrom As Long
  Dim lCTo   As Long
  Dim lSrcR  As Long
  Dim lSrcG  As Long
  Dim lSrcB  As Long
  Dim lDstR  As Long
  Dim lDstG  As Long
  Dim lDstB  As Long

    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColor = RGB(((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255))

End Function

Public Function HiWord(ByVal dw As Long) As Integer

    If dw And &H80000000 Then
        HiWord = (dw \ 65535) - 1
      Else
        HiWord = dw \ 65535
    End If

End Function

Public Function LoWord(ByVal dw As Long) As Integer

    If dw And &H8000& Then
        LoWord = &H8000 Or (dw And &H7FFF&)
      Else
        LoWord = dw And &HFFFF&
    End If

End Function

Public Function ObjectFromPtr(ByVal lPtr As Long) As Object

  Dim oThis As Object

    '-- end code
    ' ******************************************************************************
    ' Routine       : ObjectFromPtr
    ' Created by    : Marclei V Silva
    ' Machine       : ZEUS
    ' Date-Time     : 28/08/005:17:24
    ' Inputs        : lPtr - pointer to the object
    ' Outputs       : An object
    ' Credits       : SP MacMahon (www.vbaccelerator.com articles)
    ' Modifications : None
    ' Description   : Get an object from the given pointer
    ' ******************************************************************************
    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory oThis, lPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oThis
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oThis, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but this will be your code rather than the uncounted reference!

End Function

Public Function PtrFromObject(ByRef oThis As Object) As Long

  ' ******************************************************************************
  ' Routine       : PtrFromObject
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:19:00
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : Get a pointer fro a object
  ' ******************************************************************************
  ' Return the pointer to this object:

    PtrFromObject = ObjPtr(oThis)

End Function

Public Sub RemoveTitleBar(frm As Form)

  Static OriginalStyle As Long
  Dim CurrentStyle As Long
  Dim x As Long

    OriginalStyle = 0
    CurrentStyle = GetWindowLong(frm.hwnd, GWL_STYLE)

    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_DLGFRAME)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_SYSMENU)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_MINIMIZEBOX)
    OriginalStyle = OriginalStyle Or (CurrentStyle And WS_MAXIMIZEBOX)

    CurrentStyle = CurrentStyle And Not WS_DLGFRAME
    CurrentStyle = CurrentStyle And Not WS_SYSMENU
    CurrentStyle = CurrentStyle And Not WS_MINIMIZEBOX
    CurrentStyle = CurrentStyle And Not WS_MAXIMIZEBOX

    x = SetWindowLong(frm.hwnd, GWL_STYLE, CurrentStyle)
    frm.Refresh

End Sub

Public Sub RestoreTitleBar(frm As Form)

  Static OriginalStyle As Long
  Dim CurrentStyle As Long
  Dim x As Long

    CurrentStyle = GetWindowLong(frm.hwnd, GWL_STYLE)
    CurrentStyle = CurrentStyle Or OriginalStyle
    x = SetWindowLong(frm.hwnd, GWL_STYLE, CurrentStyle)
    frm.Refresh

End Sub

Public Property Get VSNetBackgroundColor() As Long

    VSNetBackgroundColor = BlendColor(vbWindowBackground, vbButtonFace, 220)

End Property

Public Property Get VSNetBorderColor() As Long

    VSNetBorderColor = TranslateColor(vbHighlight)

End Property

Public Property Get VSNetCheckedColor() As Long

    VSNetCheckedColor = BlendColor(vbHighlight, vbWindowBackground, 30)

End Property

Public Property Get VSNetControlColor() As Long

    VSNetControlColor = BlendColor(vbButtonFace, VSNetBackgroundColor, 195)

End Property

Public Property Get VSNetPressedColor() As Long

    VSNetPressedColor = BlendColor(vbHighlight, VSNetSelectionColor, 70)

End Property

Public Property Get VSNetSelectionColor() As Long

    VSNetSelectionColor = BlendColor(vbHighlight, vbWindowBackground, 70)

End Property

