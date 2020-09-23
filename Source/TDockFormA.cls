VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDockForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************************************
' Project     : TabDock
' Module      : TDockForm
' Description : Save the form properties and information to be used
'               by the docking engine
' Created by  : Marclei V Silva
' Machine     : ZEUS
' Date-Time   : 11/05/2000 22:15:58
' Revisions   : See revisions.rtf shipped with this version
' ********************************************************************
Option Explicit
'Dim Isubclass As New Isubclass
Private m_wndprcNext                As Long
#If debugWINDOWPROC Then
Private m_SCHook                    As WindowProcHook
#End If
Private WithEvents MTimer           As CTimer
Attribute MTimer.VB_VarHelpID = -1
Private m_bOverMax                  As Boolean
Private m_bOverClose                As Boolean
Private m_hBrush                    As Long
Private WithEvents m_Sizer          As cSizer
Attribute m_Sizer.VB_VarHelpID = -1
Private CloseRec                    As RECT
Private MaxRec                      As RECT
' Keep up with the errors
Private Const g_ErrConstant         As Long = vbObjectError + 1000
Private Const m_constClassName      As String = "TDockForm"
Private PushDown                    As Boolean
Private StoreHWND                   As Long
Private m_lngErrNum                 As Long
Private m_strErrStr                 As String
Private m_strErrSource              As String
Private m_bActive                   As Boolean
Private m_bStyleChanged             As Boolean
Private m_bUnloaded                 As Boolean
' class properties
Private m_bDragging                 As Boolean        ' We are dragging?
Private m_SelectedItem              As Integer        ' Host that was active during moving
Private m_lFloatingWidth            As Long           ' Floating dimensions
Private m_lFloatingHeight           As Long
Private m_lFloatingLeft             As Long
Private m_lFloatingTop              As Long
Private m_State                     As tdDockedState  ' form state dock/undocked
Private m_Style                     As tdDockStyles   ' availlable style
Private m_PanelPtr                  As Long           ' pointer to panel object
Private m_Key                       As String         ' key associatate to the form
Private m_Percent                   As Integer        ' percent is how much of Panel form will occupy
Private m_hWnd                      As Long           ' handle for the form
Private m_FormPtr                   As Long           ' pointer to the form object
Private m_bRedraw                   As Boolean        ' flag that stop form painting or style change
Private m_bHasMaxButton             As Boolean
Private m_bHasCloseButton           As Boolean
Private m_bMaxDisabled              As Boolean
'*********************************
Public Enum EHitTestAreas
    HTERROR = (-2)
    HTTRANSPARENT = (-1)
    HTNOWHERE = 0
    HTCLIENT = 1
    HTCAPTION = 2
    HTSYSMENU = 3
    HTGROWBOX = 4
    HTMENU = 5
    HTHSCROLL = 6
    HTVSCROLL = 7
    HTMINBUTTON = 8
    HTMAXBUTTON = 9
    HTLEFT = 10
    HTRIGHT = 11
    HTTOP = 12
    HTTOPLEFT = 13
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTBORDER = 18
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private HTERROR, HTTRANSPARENT, HTNOWHERE, HTCLIENT, HTCAPTION, HTSYSMENU, HTGROWBOX, HTMENU, HTHSCROLL
Private HTVSCROLL, HTMINBUTTON, HTMAXBUTTON, HTLEFT, HTRIGHT, HTTOP, HTTOPLEFT, HTBOTTOM, HTBOTTOMLEFT
Private HTBOTTOMRIGHT, HTBORDER
#End If

' Array of objects to perform WM_NCHITTEST on:
Private m_lHTPtr                    As Long
' Array of hWnds to perform WM_NCHITTEST on:
Private m_hWndh                     As Long
' Number of objects to perform WM_NCHITTEST on:
Private m_iCount                    As Long
' Rectangles to perform hit testing on:
Private Type tHitTestRectangle
    hwnd                                As Long
    sKey                                As String
    lLeft                               As Long
    lTop                                As Long
    lRight                              As Long
    lBottom                             As Long
    eResponse                           As EHitTestAreas
End Type
Private m_tHTR                      As tHitTestRectangle
Private m_iHTRCount                 As Long
Private Const SPLITTER_HEIGHT       As Integer = 80
Private Const SPLITTER_WIDTH        As Integer = 80

Private m_SizingNow As Boolean

'--- Amend NR 10/04/2003
Private m_Tag                       As Variant        ' just for saving restoring from disk layout (MVS)
'--- end amend
Implements ISubclass

Public Function AddArea(ByVal hwnd As Long, _
                        Optional ByVal sKey As String = "") As Long

    m_iHTRCount = m_iHTRCount + 1
    'ReDim Preserve m_tHTR(1 To m_iHTRCount) As tHitTestRectangle
    AddArea = m_iHTRCount
    With m_tHTR
        .hwnd = hwnd
        If (sKey <> "") Then
            .sKey = sKey
          Else
            .sKey = CStr(m_iHTRCount)
        End If
    End With

End Function

Private Function AreaIndex(ByVal vKey As Variant) As Long

  Dim i As Long

    If (IsNumeric(vKey)) Then
        i = CLng(vKey)
        If (i > 0) And (i <= m_iHTRCount) Then
            AreaIndex = i
            Exit Function
        End If
      Else
        If (m_tHTR.sKey = vKey) Then
            AreaIndex = i
            Exit Function
        End If
    End If
    ' error

End Function

Friend Sub Attach(ByVal hWndA As Long)

  ' ******************************************************************************
  ' Routine       : Attach
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:19:32
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : attach window messages to a specific hWnd
  ' ******************************************************************************

    Detach
    m_hWnd = hWndA
    AttachMessage Me, m_hWnd, WM_NCACTIVATE
    AttachMessage Me, m_hWnd, WM_ACTIVATEAPP
    AttachMessage Me, m_hWnd, WM_SIZE
    AttachMessage Me, m_hWnd, WM_MOVE
    AttachMessage Me, m_hWnd, WM_SYSCOMMAND
    AttachMessage Me, m_hWnd, WM_STYLECHANGED
    AttachMessage Me, m_hWnd, WM_CLOSE
    AttachMessage Me, m_hWnd, WM_CONTEXTMENU
    AttachMessage Me, m_hWnd, WM_NCRBUTTONDOWN
    AttachMessage Me, m_hWnd, WM_PAINT
    AttachMessage Me, m_hWnd, WM_ACTIVATE
    AttachMessage Me, m_hWnd, WM_NCHITTEST
    AttachMessage Me, m_hWnd, WM_NCLBUTTONDBLCLK
    AttachMessage Me, m_hWnd, WM_NCPAINT
    Set m_Sizer = New cSizer
    If Panel.Parent.SmartSizing Then
        m_Sizer.NoFullDrag = True
      Else
        m_Sizer.NoFullDrag = False
    End If
    If Not m_Sizer Is Nothing Then
        m_Sizer.Attach m_hWnd
    End If
    'AttachMessage Me, m_hWnd, WM_MOUSEMOVE
    'AttachMessage Me, m_hWnd, WM_NCMOUSEMOVE

End Sub

Public Sub AttachHitTester(ByRef objThis As Object)

  Dim hwnd As Long

    hwnd = objThis.hwnd
    m_lHTPtr = ObjPtr(objThis)
    m_hWndh = hwnd
    AttachMessage Me, m_hWndh, WM_NCHITTEST
    'AttachMessage Me, m_hWndh, WM_SYSCOMMAND

End Sub

Private Sub Class_Initialize()

  Dim i             As Integer
  Const constSource As String = m_constClassName & ".Class_Initialize"

    ' ******************************************************************************
    ' Routine       : (Sub) Class_Initialize
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 25/06/2000 - 1:36:23
    ' Inputs        : N/A
    ' Outputs       : N/A
    ' Modifications : N/A
    ' Description   : Initialize class
    ' ******************************************************************************
    On Error GoTo Err_Class_Initialize
    Set MTimer = New CTimer
    MTimer.Interval = 100
    MTimer.PulseTimer
    m_bActive = False
    m_bStyleChanged = False
    m_SelectedItem = 0
    For i = 0 To 3
        m_lPattern(i) = &HAAAA5555
    Next i
    createBrush

Exit Sub

Err_Class_Initialize:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub Class_Terminate()

  Const constSource As String = m_constClassName & ".Class_Terminate"

    On Error Resume Next

        ' ******************************************************************************
        ' Routine       : (Sub) Class_Terminate
        ' Created by    : Marclei V Silva
        ' Company Name  : Spnorte Consultoria
        ' Machine       : ZEUS
        ' Date-Time     : 25/06/2000 - 1:36:03
        ' Inputs        : N/A
        ' Outputs       : N/A
        ' Modifications : N/A
        ' Description   : Terminate the class (clean it up)
        ' ******************************************************************************
    On Error GoTo Err_Class_Terminate
    MTimer.Interval = 0
    Set MTimer = Nothing
    ' detach all window messages
    Detach
    ' try to unload the extender
    Unload Extender

Exit Sub

Err_Class_Terminate:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
    On Error GoTo 0

End Sub

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

Public Sub DestroyHitTester()

    If (m_hWnd <> 0) Then
        DetachMessage Me, m_hWndh, WM_NCHITTEST
        DetachMessage Me, m_hWndh, WM_SYSCOMMAND
        m_hWndh = 0
    End If
    m_lHTPtr = 0

End Sub

Friend Sub Detach()

  ' ******************************************************************************
  ' Routine       : Detach
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:19:55
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : Detach window messages
  ' ******************************************************************************

    If Not m_hWnd = 0 Then
        DetachMessage Me, m_hWnd, WM_DESTROY
        DetachMessage Me, m_hWnd, WM_SIZE
        DetachMessage Me, m_hWnd, WM_ACTIVATE
        DetachMessage Me, m_hWnd, WM_MOVE
        DetachMessage Me, m_hWnd, WM_SYSCOMMAND
        DetachMessage Me, m_hWnd, WM_STYLECHANGED
        DetachMessage Me, m_hWnd, WM_CLOSE
        DetachMessage Me, m_hWnd, WM_CONTEXTMENU
        DetachMessage Me, m_hWnd, WM_NCRBUTTONDOWN
        DetachMessage Me, m_hWnd, WM_PAINT
        DetachMessage Me, m_hWnd, WM_NCACTIVATE
        DetachMessage Me, m_hWnd, WM_ACTIVATEAPP
        DetachMessage Me, m_hWnd, WM_NCHITTEST
        DetachMessage Me, m_hWnd, WM_NCLBUTTONDBLCLK
        DetachMessage Me, m_hWnd, WM_NCPAINT
        If Not m_Sizer Is Nothing Then
            m_Sizer.Detach
            Set m_Sizer = Nothing
        End If
        'DetachMessage Me, m_hWnd, WM_MOUSEMOVE
        'DetachMessage Me, m_hWnd, WM_NCMOUSEMOVE
    End If

End Sub

Public Function dfPaint()

    DockForm_Paint

End Function

Private Sub DockForm_Activate(ByVal bByMouse As Boolean)

    DockForm_Paint

End Sub

Private Sub DockForm_CaptionClick(ByVal Button As Integer, _
                                  ByVal x As Single, _
                                  ByVal Y As Single)

  ' trigger event

    Panel.Parent.TriggerEvent "CaptionClick", Object, Button, x * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    ' When I trigger this event the program locks up
    ' then I discovered that when we send the focus to the main window
    ' the problem is solved. Why? I really don't know
    ' I'm not so good at API the way you may think
    ' but for now it is working fine...
    Putfocus Panel.Parent.Parent.hwnd

End Sub

Private Sub DockForm_CaptionClose()

  Dim pt            As POINTAPI
  Dim ptPrev        As POINTAPI
  Dim objRect       As RECT
  Dim lObjWidth     As Long
  Dim lObjHeight    As Long
  Dim captionHeight As Long
  Dim hdc           As Long
  Dim frameRect     As RECT

    ReleaseCapture
    ''debug.print Me.Extender.Caption
    GetWindowRect Me.Extender.hwnd, objRect
    'objRect.Left = 0
    'objRect.Top = 0
    'objRect.Bottom = Extender.height / Screen.TwipsPerPixelY
    'objRect.Right = Extender.width / Screen.TwipsPerPixelY
    ' First get the window DC
    captionHeight = getCaptionButtonHeight
    lObjWidth = objRect.Right - objRect.Left
    lObjHeight = objRect.Bottom - objRect.Top
    GetCursorPos pt
    'Store the initial cursor position
    ptPrev.x = pt.x
    ptPrev.Y = pt.Y
    '****************************************************
    With frameRect
        .Top = objRect.Top + 5
        .Left = objRect.Right - captionHeight
        .Right = objRect.Right - 3
        .Bottom = .Top + captionHeight
    End With
    ''debug.print pt.x, pt.Y
    ''debug.print frameRect.Top, frameRect.Left, frameRect.Right, frameRect.Bottom
    '*****************************************************
    If PtInRect(frameRect, pt.x, pt.Y) Then
        hdc = GetWindowDC(Me.Extender.hwnd)
        GetClientRect Me.Extender.hwnd, objRect
        With frameRect
            .Top = objRect.Top + 5
            .Left = objRect.Right - captionHeight + 5
            .Right = objRect.Right + 2
            .Bottom = .Top + captionHeight - 3
        End With
        DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONCLOSE + DFCS_INACTIVE
        ''debug.print "redraw"
        'SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
        'Exit Sub
    End If
    '*****************************************************

End Sub

Private Sub DockForm_DblClick()

  ''debug.print "dblclick"

    If (Me.Style And tdDockFloat) <> tdDockFloat Then
        Exit Sub
    End If
    If State = tdDocked Then
        Panel.UnDock Me
      Else
        Panel.Dock Me
    End If

End Sub

Private Sub DockForm_Deactivate()

  ' ******************************************************************************
  ' Routine       : DockForm_Deactivate
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:29:52
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : Called when the form is deactivated
  ' ******************************************************************************

    DockForm_Paint

End Sub

Public Sub DockForm_Dock()

  ' ******************************************************************************
  ' Routine       : Dock
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:22:12
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : dock the form
  ' ******************************************************************************

    Panel.Dock Me
    'DockForm_Paint
    DoEvents
    If Not Panel.Expanded Then
        'Panel.Panel_Expand
        'Panel.DockCollapse
        ''debug.print "Panel Expanded " & Panel.Expanded
    End If

End Sub

Friend Sub DockForm_Hide()

  Const constSource As String = m_constClassName & ".Hide"

    ' ******************************************************************************
    ' Routine       : (Sub) Hide
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 25/06/2000 - 1:35:19
    ' Inputs        : N/A
    ' Outputs       : N/A
    ' Modifications : N/A
    ' Description   : hides the docked form no matter it is docked or undocked
    ' ******************************************************************************
    On Error GoTo Err_Hide
    Visible = False
    If Panel.IsVisible Then
        Panel.Repaint
      Else
        Panel.Visible = False
    End If

Exit Sub

Err_Hide:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub DockForm_MouseDown(FormLeft As Long, _
                               FormTop As Long, _
                               FormWidth As Long, _
                               FormHeight As Long, _
                               Cancel As Boolean)

  ' ******************************************************************************
  ' Routine       : DockForm_MouseDown
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:23:00
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : Do some actions before dragging a form
  ' ******************************************************************************
  ''debug.print "mousedown"

    m_bSizing = True
    If Style And tdDockFloat Then
        m_bDragging = True
        ' if form is undocked then save the floating position
        If State = tdUndocked Then
            DockForm_StoreFloatingPos FormLeft, FormTop, FormWidth, FormHeight
        End If
      Else
        Cancel = True
    End If

End Sub

Private Sub DockForm_MouseMove(ByRef FormLeft As Long, _
                               ByRef FormTop As Long, _
                               ByRef FormWidth As Long, _
                               ByRef FormHeight As Long, _
                               ByRef lBorderWidth As Long)

  Dim RC         As RECT
  Dim i          As Integer
  Dim Found      As Boolean
  Dim pt         As POINTAPI
  Dim lX1Offset  As Long
  Dim lX2Offset  As Long
  Dim lObjWidth  As Long
  Dim lObjHeight As Long

    ' ******************************************************************************
    ' Routine       : DockForm_MouseMove
    ' Created by    : Marclei V Silva
    ' Machine       : ZEUS
    ' Date-Time     : 28/08/005:25:03
    ' Inputs        :
    ' Outputs       :
    ' Credits       :
    ' Modifications :
    ' Description   : Called when the form is moving around
    ' ******************************************************************************
    ' flag that indicates we found a dock host
    Found = False
    ' get the cursor position
    GetCursorPos pt
    GetWindowRect Panel.hwnd, RC
    'resetVSNetRecs
    If PointInObject(Me, RC.Left, RC.Top) Then
        'debug.print "kkdk d d d "
    End If
    ' loop the panels to see we find a docking host
    ' for this form
    For i = 1 To 4
        ' get the host rect
        Panel.Parent.Panels(i).GetHostRect RC
        ' check if the form is within the panel's boundaries
        If PtInRect(RC, pt.x, pt.Y) Then
            If Panel.Parent.Panels(i).AllowDocking(Me) Then
                ' a host was found and it is different from the
                ' the current selected one
                If Not Panel.Parent.Panels(i).Index = m_SelectedItem Then
                    ' host selection changed here
                    m_SelectedItem = Panel.Parent.Panels(i).Index
                    ' we don't have a dragging boundary set yet
                    ' for this host then update with the host's rect
                    lObjWidth = RC.Right - RC.Left
                    lObjHeight = RC.Bottom - RC.Top
                    If Panel.Parent.Panels(i).Align = tdAlignLeft Or Panel.Parent.Panels(i).Align = tdAlignRight Then
                        lX1Offset = pt.x - FormLeft
                        lX2Offset = lObjWidth * lX1Offset / FormWidth 'pt.Y - rc.Top
                        FormLeft = pt.x - lX2Offset
                        FormWidth = lObjWidth
                        FormHeight = lObjHeight
                      Else
                        lX1Offset = pt.Y - FormTop
                        lX2Offset = lObjHeight * lX1Offset / FormHeight 'pt.Y - rc.Top
                        FormTop = pt.Y - lX2Offset
                        FormWidth = lObjWidth
                        FormHeight = lObjHeight
                    End If
                End If
                ' change border width (special effect)
                lBorderWidth = 2
                ' we found a host, set flag = true
                Found = True
            End If
        End If
        ' we have found a host so exit
        If Found Then
            Exit For
        End If
    Next '  I I
    ' if the form has not found any host to dock in
    ' then let the rectangle image the same as its
    ' floating original rectangle
    If Not Found Then
        ' if we are dragging and the form is docked
        If m_bDragging And m_State = tdDocked Then
            ' get active host rect
            GetWindowRect Panel.hwnd, RC
            ' restore form left property to its
            ' floating left position
            lObjWidth = RC.Right - RC.Left
            lObjHeight = RC.Bottom - RC.Top
            lX1Offset = pt.x - FormLeft
            lX2Offset = m_lFloatingWidth * lX1Offset / lObjWidth 'pt.Y - rc.Top
            FormLeft = pt.x - lX2Offset
            ' flag must be false
            m_bDragging = False
        End If
        ' change border width (just a special effect)
        lBorderWidth = 3
        ' update width and height with floating dimensions
        FormWidth = m_lFloatingWidth
        FormHeight = m_lFloatingHeight
        ' store floating position
        DockForm_StoreFloatingPos FormLeft, FormTop, FormWidth, FormHeight
        ' clear host selection
        m_SelectedItem = 0
    End If

End Sub

Private Sub DockForm_MouseUp(FormLeft As Long, _
                             FormTop As Long, _
                             FormWidth As Long, _
                             FormHeight As Long)

  ' ******************************************************************************
  ' Routine       : DockForm_MouseUp
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:23:29
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : Actions after a form is dropped somewhere
  ' ******************************************************************************
  ' If no host was found to dock the form then
  ' undock it if it is not undocked

    m_bSizing = False
    If m_SelectedItem = 0 Then
        If m_State = tdDocked Then
            ' undock the form only if it can float
            If m_Style And tdDockFloat Then
                Panel.UnDock Me
                ' repaint the host
                Panel.Repaint
            End If
        End If
        Exit Sub
    End If
    ' if form is undocked then dock it
    If m_State = tdUndocked Then
        ' dock the form
        If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
            Panel.Parent.Panels(m_SelectedItem).Dock Me
        End If
      ElseIf m_State = tdDocked Then
        ' if the selected host is the active host
        ' just repaint the host panel
        If m_SelectedItem = Panel.Index Then
            Panel.DockSwap Me, FormLeft, FormTop
            Panel.Repaint
            DockForm_Paint
          Else
            ' if selected host is different from
            ' active host, first undock the form
            If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
                Panel.UnDock Me
                ' dock the form into the selected host
                Panel.Parent.Panels(m_SelectedItem).Dock Me
            End If
        End If
    End If
    ' A docking site changed then we must set
    ' the new dock host for this form
    If Panel.Parent.Panels(m_SelectedItem).AllowDocking(Me) Then
        Set Panel = Panel.Parent.Panels(m_SelectedItem)
    End If

End Sub

Private Sub DockForm_Paint()

  Dim RC            As RECT
  Dim bdrStyle      As Long
  Dim bdrSides      As Long
  Dim BorderStyle   As tdBorderStyles
  Dim CaptionStyle  As tdCaptionStyles
  Dim hdc           As Long
  Dim captionHeight As Long
  Dim tDock         As TDockForm
  Dim visCount      As Integer
  Dim captionRect   As RECT
  Dim frameRect     As RECT
  Dim hbr           As Long

    RefreshVSNet

    If Not Me.Panel.Expanded Then
        Exit Sub
    End If

    If Me.State = tdUndocked Or m_bRedraw = False Or Extender.Visible = False Then
        Exit Sub
    End If
    '-- Get The Visible Panel Count
    For Each tDock In Panel.WindowList
        If tDock.Visible Then
            visCount = visCount + 1
        End If
    Next '  TDOCK TDOCK
    ' draw a custom border based on parante's color
    DrawBorder Extender, Panel.Parent.BackColor
    ' retrieve TabDock border style
    BorderStyle = Panel.Parent.BorderStyle
    CaptionStyle = Panel.Parent.CaptionStyle
    ' all sides must be updated
    bdrSides = BF_RECT
    ' update border styles
    If BorderStyle = bdrFlat Then
        bdrSides = bdrSides Or BF_FLAT
    End If
    If BorderStyle = bdrMono Then
        bdrSides = bdrSides Or BF_MONO
    End If
    If BorderStyle = bdrSoft Then
        bdrSides = bdrSides Or BF_SOFT
    End If
    Select Case BorderStyle
      Case bdrRaisedOuter
        bdrStyle = BDR_RAISEDOUTER
      Case bdrRaisedInner
        bdrStyle = BDR_RAISEDINNER
      Case bdrRaised
        bdrStyle = EDGE_RAISED
      Case bdrSunkenOuter
        bdrStyle = BDR_SUNKENOUTER
      Case bdrSunkenInner
        bdrStyle = BDR_SUNKENINNER
      Case bdrSunken
        bdrStyle = EDGE_SUNKEN
      Case bdrEtched
        bdrStyle = EDGE_ETCHED
      Case bdrBump
        bdrStyle = EDGE_BUMP
      Case bdrFlat
        bdrStyle = BDR_SUNKEN
      Case bdrMono
        bdrStyle = BDR_SUNKEN
      Case bdrSoft
        bdrStyle = BDR_RAISED
    End Select
    ' get a window rect by hand
    ' GetWindowRect will not work here!
    RC.Left = 0
    RC.Top = 0
    RC.Bottom = Extender.Height / Screen.TwipsPerPixelY
    RC.Right = Extender.width / Screen.TwipsPerPixelY
    ' First get the window DC
    hdc = GetWindowDC(hwnd)
    ' Simply call the API and draw the edge.
    DrawEdge hdc, RC, bdrStyle, bdrSides
    ' get a window rect by hand
    ' GetWindowRect will not work here!
    RC.Left = 0
    RC.Top = 0
    RC.Bottom = Panel.Height / Screen.TwipsPerPixelY + 1
    RC.Right = Panel.width / Screen.TwipsPerPixelX
    If Panel.Align = tdAlignLeft Then
        RC.Right = (Panel.width - SPLITTER_WIDTH) / Screen.TwipsPerPixelX + 1
        RC.Bottom = Panel.Height / Screen.TwipsPerPixelY
    End If
    If Panel.Align = tdAlignRight Then
        RC.Left = SPLITTER_WIDTH / Screen.TwipsPerPixelX
        RC.Bottom = Panel.Height / Screen.TwipsPerPixelY
    End If
    If Panel.Align = tdAlignTop Then
        RC.Bottom = (Panel.Height - SPLITTER_HEIGHT) / Screen.TwipsPerPixelY + 1
    End If
    If Panel.Align = tdAlignBottom Then
        RC.Top = SPLITTER_HEIGHT / Screen.TwipsPerPixelY
        RC.Bottom = Panel.Height / Screen.TwipsPerPixelY
    End If
    ' First get the window DC
    hdc = GetWindowDC(Panel.hwnd)
    ' Simply call the API and draw the edge.
    'removed as it causes extra lines to be drawn  yuk!!!
    'DrawEdge hDC, Rc, bdrStyle, bdrSides
    RC.Left = 0
    RC.Top = 0
    RC.Bottom = Extender.Height / Screen.TwipsPerPixelY
    RC.Right = Extender.width / Screen.TwipsPerPixelY
    hdc = GetWindowDC(hwnd)
    '************************************************
    If Me.Extender.BorderStyle = vbBSNone Or Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
        'draw custom caption here!!!!
        captionHeight = getCaptionHeight + 6
        hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
        With captionRect
            .Top = RC.Top + 3
            .Left = RC.Left + 3
            .Right = RC.Right - 3
            .Bottom = captionHeight - 4
            If Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
                .Top = .Top
                .Bottom = RC.Bottom - 3
                .Right = captionHeight - 4
            End If
        End With
        'blank out current caption
        FillRect hdc, captionRect, hbr
        DeleteObject hbr
        Select Case CaptionStyle
          Case tdCaptionNormal ' = 0
            drawGradient captionRect, hdc, Me.Extender.Caption, m_bActive, False, Panel.Align, Me.Extender
          Case tdCaptionEtched ' = 1
            drawGripper captionRect, hdc, EDGE_ETCHED, BF_RECT, False, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSoft ' = 2
            drawGripper captionRect, hdc, BDR_RAISED, BF_RECT Or BF_SOFT, False, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionRaised ' = 3
            drawGripper captionRect, hdc, EDGE_RAISED, BF_RECT, False, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionRaisedInner ' = 4
            drawGripper captionRect, hdc, BDR_RAISEDINNER, BF_RECT, False, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSunkenOuter ' = 5
            drawGripper captionRect, hdc, BDR_SUNKENOUTER, BF_RECT, False, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSunken ' = 6
            drawGripper captionRect, hdc, EDGE_SUNKEN, BF_RECT, False, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSingleRaisedBar ' = 7
            drawGripper captionRect, hdc, EDGE_RAISED, BF_RECT, True, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSingleRaisedInner ' = 9
            drawGripper captionRect, hdc, BDR_RAISEDINNER, BF_RECT, True, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSingleSoft ' = 10
            drawGripper captionRect, hdc, BDR_RAISED, BF_RECT Or BF_SOFT, True, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSingleEtched ' = 11
            drawGripper captionRect, hdc, EDGE_ETCHED, BF_RECT, True, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSingleSunken ' = 12
            drawGripper captionRect, hdc, EDGE_SUNKEN, BF_RECT, True, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionSingleSunkenOuter ' = 13
            drawGripper captionRect, hdc, BDR_SUNKENOUTER, BF_RECT, True, captionHeight, Panel.Align, m_bHasMaxButton, m_bHasCloseButton
          Case tdCaptionGradient ' = 8
            drawGradient captionRect, hdc, Me.Extender.Caption, m_bActive, True, Panel.Align, Me.Extender
          Case tdCaptionOfficeXP ' = 14
            drawOfficeXP captionRect, hdc, Me.Extender.Caption, m_bActive, True, Panel.Align, Me.Extender
          Case tdCaptionVSNet ' = 0
            drawVSNet captionRect, hdc, Me.Extender.Caption, m_bActive, True, Panel.Align, Me.Extender
            'drawGradient captionRect, hdc, Me.Extender.Caption, m_bActive, False, Panel.Align, Me.Extender
        End Select
        'draw close box on form
        If m_bHasCloseButton Then
            With frameRect
                .Top = RC.Top + 5
                .Left = RC.Right - captionHeight + 6
                .Right = RC.Right - 4
                .Bottom = captionHeight - 5
                If Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
                    .Right = RC.Left + captionHeight - 6
                    .Left = RC.Left + 4
                End If
            End With
            frameRect.Left = frameRect.Left + 1
            If CaptionStyle = tdCaptionVSNet Then
                'debug.print "Here"
                LSet CloseRec = frameRect
                CloseRec.Bottom = CloseRec.Bottom - 1
                ''-HoverCaption
                If m_bOverClose Then
                    hbr = GetSysColorBrush(VSNetSelectionColor And &H1F&)
                  Else
                    hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                End If
                m_bOverClose = False
                '                ' drawcloseButton hdc, hbr, frameRect
              Else
                DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONCLOSE 'Or DFCS_FLAT
            End If
            'm_bCloseOver = False
            frameRect.Left = frameRect.Left - 1
        End If
        If m_bHasMaxButton And m_bHasCloseButton Then
            If Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
                With frameRect
                    .Top = .Bottom + 2
                    .Bottom = .Top + captionHeight - 10
                    '.Right = .Right + 1
                    '.Left = .Left - 1
                End With
                'DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE
                frameRect.Left = frameRect.Left + 1
                '- Enable or Disable The Max Button
                If Panel.WindowList.Count > 1 And visCount > 1 Then
                    If CaptionStyle = tdCaptionVSNet Then
                        LSet MaxRec = frameRect
                        MaxRec.Bottom = MaxRec.Bottom - 1
                       ' '-HoverCaption
                        If m_bOverMax Then
                            hbr = GetSysColorBrush(VSNetSelectionColor And &H1F&)
                          Else
                            hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        End If
                        'hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        ' drawChevronButton hdc, hbr, frameRect
                      Else
                        DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE
                    End If
                    m_bMaxDisabled = False
                  Else
                    LSet MaxRec = frameRect
                    MaxRec.Bottom = MaxRec.Bottom - 1
                    '-HoverCaption
                    If CaptionStyle = tdCaptionVSNet Then
                        If m_bOverMax Then
                            hbr = GetSysColorBrush(VSNetSelectionColor And &H1F&)
                          Else
                            hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        End If
                        'hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        ' drawChevronButton hdc, hbr, frameRect
                      Else
                        DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_INACTIVE
                    End If
                    m_bMaxDisabled = True
                End If
                frameRect.Left = frameRect.Left - 1
              Else
                With frameRect
                    .Right = .Left - 2
                    .Left = .Right - captionHeight + 10
                End With
                'DrawFrameControl hDC, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE 'Or DFCS_FLAT
                If Panel.WindowList.Count > 1 And visCount > 1 Then
                    LSet MaxRec = frameRect
                    MaxRec.Bottom = MaxRec.Bottom - 1
                    '-HoverCaption
                    If CaptionStyle = tdCaptionVSNet Then
                        If m_bOverMax Then
                            hbr = GetSysColorBrush(VSNetSelectionColor And &H1F&)
                          Else
                            hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        End If
                        '                        hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        '     ' drawChevronButton hdc, hbr, frameRect
                      Else
                        DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE
                    End If
                    m_bMaxDisabled = False
                  Else
                    If CaptionStyle = tdCaptionVSNet Then
                        LSet MaxRec = frameRect
                        MaxRec.Bottom = MaxRec.Bottom - 1
                        '-HoverCaption
                        '                    hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        If m_bOverMax Then
                            hbr = GetSysColorBrush(VSNetSelectionColor And &H1F&)
                          Else
                            hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        End If
                        ' ' drawChevronButton hdc, hbr, frameRect
                      Else
                        DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_INACTIVE
                    End If
                    m_bMaxDisabled = True
                End If
            End If
          ElseIf m_bHasMaxButton And Not m_bHasCloseButton Then
            With frameRect
                .Top = RC.Top + 5
                .Left = RC.Right - captionHeight + 6
                .Right = RC.Right - 4
                .Bottom = captionHeight - 5
                If Panel.Align = tdAlignTop Or Panel.Align = tdAlignBottom Then
                    .Right = RC.Left + captionHeight - 6
                    .Left = RC.Left + 4
                End If
                .Left = .Left + 1
                '- Enable or Disable The Max Button
                LSet MaxRec = frameRect
                MaxRec.Bottom = MaxRec.Bottom - 1
                '-HoverCaption
                If Panel.WindowList.Count > 1 And visCount > 1 Then
                    If CaptionStyle = tdCaptionVSNet Then
                        If m_bOverMax Then
                            hbr = GetSysColorBrush(VSNetSelectionColor And &H1F&)
                          Else
                            hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        End If
                        '                        hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        ' ' drawChevronButton hdc, hbr, frameRect
                      Else
                        DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE
                    End If
                    m_bMaxDisabled = False
                  Else
                    If CaptionStyle = tdCaptionVSNet Then
                        LSet MaxRec = frameRect
                        MaxRec.Bottom = MaxRec.Bottom - 1
                        '-HoverCaption
                        If m_bOverMax Then
                            hbr = GetSysColorBrush(VSNetSelectionColor And &H1F&)
                          Else
                            hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        End If
                        '                        hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        '  ' drawChevronButton hdc, hbr, frameRect
                      Else
                        DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_INACTIVE
                    End If
                    '            LSet MaxRec = frameRect
                    m_bMaxDisabled = True
                End If
            End With
        End If
    End If
    '*********************************************************
    ' release it
    'debug.print "Repaint"
    DeleteObject hbr
    ReleaseDC hwnd, hdc

End Sub

Private Sub DockForm_ResetStyle()

  Dim New_Style As Long

    ' ******************************************************************************
    ' Routine       : DockForm_ResetStyle
    ' Created by    : Marclei V Silva
    ' Machine       : ZEUS
    ' Date-Time     : 28/08/005:27:39
    ' Inputs        :
    ' Outputs       :
    ' Credits       :
    ' Modifications :
    ' Description   : Resets the style to a tollwindow
    '                 This routine is necessay when the style of
    '                 the form changes to force it to be a toolwindow
    ' ******************************************************************************
    ' if redraw is disable do not set any style
    ' or else you may get a loop here
    If m_bRedraw = False Then
        Exit Sub
    End If
    ' set this flag to true so that the next set style event will
    ' not be handled (trick!)
    m_bStyleChanged = True
    ' get current window style
    New_Style = GetWindowLong(hwnd, GWL_STYLE)
    ' check form state
    If State = tdDocked Then
        New_Style = New_Style And Not WS_THICKFRAME
        SetWindowLong hwnd, GWL_STYLE, New_Style
      Else
        New_Style = New_Style And Not WS_DLGFRAME
        New_Style = New_Style Or WS_CAPTION
        New_Style = New_Style Or WS_THICKFRAME
    End If
    ' get extended style
    New_Style = GetWindowLong(hwnd, GWL_EXSTYLE)
    ' set it to toolwindow
    New_Style = New_Style Or WS_EX_TOOLWINDOW
    New_Style = New_Style And Not (WS_EX_APPWINDOW)
    SetWindowLong hwnd, GWL_EXSTYLE, New_Style
    ' reset flag here
    m_bStyleChanged = False
    DockForm_Paint

End Sub

Friend Sub DockForm_Show()

  Const constSource As String = m_constClassName & ".DockForm_Show"

    ' ******************************************************************************
    ' Routine       : (Sub) DockForm_Show
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 25/06/2000 - 1:35:36
    ' Inputs        : N/A
    ' Outputs       : N/A
    ' Modifications : N/A
    ' Description   : shows the docked form no matter it is docked or undocked
    ' ******************************************************************************
    On Error GoTo Err_DockForm_Show
    If State = tdDocked Then
        Panel.Dock Me
      Else
        Panel.UnDock Me
    End If
    Attach Me.hwnd
    m_bUnloaded = False
    DockForm_ResetStyle
    DockForm_Paint

Exit Sub

Err_DockForm_Show:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub DockForm_StoreFloatingPos(lLeft As Long, _
                                      lTop As Long, _
                                      lWidth As Long, _
                                      lHeight As Long)

  ' ******************************************************************************
  ' Routine       : (Sub) DockForm_StoreFloatingPos
  ' Created by    : Marclei V Silva
  ' Company Name  : Spnorte Consultoria
  ' Machine       : ZEUS
  ' Date-Time     : 09/06/2000 12:47:34
  ' Inputs        :
  ' Outputs       :
  ' Modifications :
  ' Description   : Store the floating position of the form
  ' ******************************************************************************

    m_lFloatingLeft = lLeft
    m_lFloatingTop = lTop
    m_lFloatingWidth = lWidth
    m_lFloatingHeight = lHeight

End Sub

Public Sub DockForm_UnDock()

  ' ******************************************************************************
  ' Routine       : UnDock
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:22:26
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : undock the form
  ' ******************************************************************************

    Panel.UnDock Me
    DockForm_Paint

End Sub

Private Sub DockForm_Unload()

  ' ******************************************************************************
  ' Routine       : DockForm_Unload
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:27:09
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : Hides the form
  ' ******************************************************************************

    DockForm_Hide

End Sub

Private Sub DragObject(ByVal hwnd As Long)

  Dim pt              As POINTAPI
  Dim ptPrev          As POINTAPI
  Dim objRect         As RECT
  Dim DragRect        As RECT
  Dim lBorderWidth    As Long
  Dim lObjTop         As Long
  Dim lObjLeft        As Long
  Dim lObjWidth       As Long
  Dim lObjHeight      As Long
  Dim lXOffset        As Long
  Dim lYOffset        As Long
  Dim captionHeight   As Long
  Dim bMoved          As Boolean
  Dim bCancel         As Boolean
  Dim hdc             As Long
  Dim tDock           As TDockForm
  Dim tDockItem       As Integer
  Dim frameRect       As RECT
  Dim frameRectShrink As RECT
  Dim hbr             As Long

    ' ******************************************************************************
    ' Routine       : (Sub) DragObject
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:38:15
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Credits       : This rotine was extract from the DockForm_Dragger project
    '               : by Steve
    '               : Steve81@ mediaone.net
    '               : http://www.vbtutor.com
    '               : Thanks Steve!
    '               :
    ' Modifications : Added the ability to change the dragging border width
    '               : Removed reposition
    '               :
    ' Description   : Procedure which simulates windows dragging of an object
    ' ******************************************************************************
    ReleaseCapture
    GetWindowRect hwnd, objRect
    captionHeight = getCaptionButtonHeight
    lObjWidth = objRect.Right - objRect.Left
    lObjHeight = objRect.Bottom - objRect.Top
    GetCursorPos pt
    'Store the initial cursor position
    ptPrev.x = pt.x
    ptPrev.Y = pt.Y
    'resetVSNetRecs
    HoverCaption
    '****************************************************
    With frameRect
        .Top = objRect.Top + 5
        .Left = objRect.Right - captionHeight
        .Right = objRect.Right - 3
        .Bottom = .Top + captionHeight
        If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
            frameRectShrink.Top = .Bottom + 2
            frameRectShrink.Bottom = frameRectShrink.Top + captionHeight - 7
          Else
            frameRectShrink.Top = .Top
            frameRectShrink.Bottom = .Bottom
        End If
        If Me.State = tdDocked Then
            If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
                .Right = objRect.Left + captionHeight
                .Left = objRect.Left + 4
                frameRectShrink.Right = .Right
                frameRectShrink.Left = .Left
              Else
                frameRectShrink.Right = .Left - 2
                frameRectShrink.Left = frameRectShrink.Right - captionHeight + 3
            End If
        End If
    End With
    
    
    '*****************************************************
    If PtInRect(frameRect, pt.x, pt.Y) Then
        hdc = GetWindowDC(Me.Extender.hwnd)
        GetClientRect Me.Extender.hwnd, objRect
        With frameRect
            .Top = objRect.Top + 5
            .Left = objRect.Right - captionHeight + 5 - 6
            .Right = objRect.Right + 2 - 6
            .Bottom = .Top + captionHeight - 3
            If Me.State = tdDocked Then
                If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
                    .Right = objRect.Left + captionHeight
                    .Left = objRect.Left + 4
                End If
              Else
                .Right = .Right + 5
                .Left = .Left + 5
                .Top = .Top + 1
            End If
        End With
        '-- Ammend to Take Care Of The Close and Max Buttons
        '-- By: Gary Noble (Phantom Man(PSC))
        If m_bHasCloseButton Then
            If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                hbr = CreateSolidBrush(VSNetCheckedColor)
                drawCloseButton hdc, hbr, CloseRec
                DeleteObject hbr
              Else
                DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONCLOSE Or DFCS_PUSHED
            End If
            Sleep (150)
            SendMessage hwnd, WM_SYSCOMMAND, SC_CLOSE, 0
           ' HoverCaption
          ElseIf Not m_bHasCloseButton And m_bHasMaxButton Then
            If Not m_bMaxDisabled Then
                If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                    hbr = CreateSolidBrush(VSNetCheckedColor)
                    frameRect.Left = frameRect.Left + 1
                    frameRect.Right = frameRect.Right + 1
                    If Not m_bMaxDisabled Then
                        drawChevronButton hdc, hbr, MaxRec
                    End If
                    DeleteObject hbr
                  Else
                    DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_PUSHED
                End If
            End If
            tDockItem = 1
            Sleep (150)
            If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                hbr = CreateSolidBrush(VSNetCheckedColor)
                If Not m_bMaxDisabled Then
                    drawChevronButton hdc, hbr, MaxRec
                End If
                DeleteObject hbr
              Else
                DrawFrameControl hdc, frameRect, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_PUSHED
            End If
            For Each tDock In Panel.WindowList
                If tDock Is Me Then
                    Panel.DockCollapse tDockItem
                  Else
                    tDockItem = tDockItem + 1
                End If
            Next '  TDOCK TDOCK
        End If
        DockForm_Paint
        DoEvents
        m_bChevronDown = False
        m_bChevronOver = False
        m_bCloseDown = False
        m_bCloseOver = False
       'HoverCaption'
        Exit Sub
    End If
    '-- OK UpWards
    If PtInRect(frameRectShrink, pt.x, pt.Y) And m_bHasMaxButton And m_bHasCloseButton Then
        hdc = GetWindowDC(Me.Extender.hwnd)
        GetClientRect Me.Extender.hwnd, objRect
        With frameRect
            .Top = objRect.Top + 5
            .Left = objRect.Right - captionHeight + 5
            .Right = objRect.Right + 2
            .Bottom = .Top + captionHeight - 3
            If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
                frameRectShrink.Top = .Bottom + 2
                frameRectShrink.Bottom = frameRectShrink.Top + captionHeight - 3
              Else
                frameRectShrink.Top = .Top
                frameRectShrink.Bottom = .Bottom
            End If
            If Me.State = tdDocked Then
                If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
                    .Right = objRect.Left + captionHeight
                    .Left = objRect.Left + 4
                    frameRectShrink.Right = .Right + 2
                    frameRectShrink.Left = .Left + 2
                  Else
                    frameRectShrink.Right = .Left - 2 - 6
                    frameRectShrink.Left = frameRectShrink.Right - captionHeight + 3
                End If
            End If
        End With
        '-----
        If Panel.Align = tdAlignBottom Or Panel.Align = tdAlignTop Then
            If Not m_bMaxDisabled Then
                If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                   ' m_bChevronOver = True
                    hbr = CreateSolidBrush(VSNetCheckedColor)
                    If Not m_bMaxDisabled Then
                        drawChevronButton hdc, hbr, MaxRec
                    End If
                    DeleteObject hbr
                  Else
                    DrawFrameControl hdc, frameRectShrink, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_PUSHED
                End If
                Sleep (150)
                DoEvents
                If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                    hbr = CreateSolidBrush(VSNetCheckedColor)
                    If Not m_bMaxDisabled Then
                        drawChevronButton hdc, hbr, MaxRec
                    End If
                    DeleteObject hbr
                  Else
                    DrawFrameControl hdc, frameRectShrink, DFC_CAPTION, DFCS_CAPTIONRESTORE
                End If
            End If
          Else
            If Not m_bMaxDisabled Then
                If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                   hbr = CreateSolidBrush(VSNetCheckedColor)
                    If Not m_bMaxDisabled Then
                        drawChevronButton hdc, hbr, MaxRec
                    End If
                    DeleteObject hbr
                  Else
                    DrawFrameControl hdc, frameRectShrink, DFC_CAPTION, DFCS_CAPTIONRESTORE Or DFCS_PUSHED
                End If
                Sleep (150)
                If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                    hbr = CreateSolidBrush(VSNetCheckedColor)
                    If Not m_bMaxDisabled Then
                        drawChevronButton hdc, hbr, MaxRec
                    End If
                    DeleteObject hbr
                  Else
                    DrawFrameControl hdc, frameRectShrink, DFC_CAPTION, DFCS_CAPTIONRESTORE
                End If
            End If
        End If
        tDockItem = 1
        For Each tDock In Panel.WindowList
            If tDock Is Me Then
                Panel.DockCollapse tDockItem
              Else
                tDockItem = tDockItem + 1
            End If
        Next '  TDOCK TDOCK
        Exit Sub
    End If
    '*****************************************************
    Select Case m_Style
      Case tdDockLeft, tdDockTop, tdDockBottom, tdDockRight
      Case Else
        'Set the initial rectangle, and draw it to show the user that
        'the object can be moved
        lXOffset = pt.x - objRect.Left
        lYOffset = pt.Y - objRect.Top
        With DragRect
            .Left = pt.x - lXOffset
            .Top = pt.Y - lYOffset
            .Right = .Left + lObjWidth
            .Bottom = .Top + lObjHeight
        End With
        ' use form border width highlighting
        lBorderWidth = 0.5
        DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
        ' Move the object
        DockForm_MouseDown pt.x - lXOffset, pt.Y - lYOffset, lObjWidth, lObjHeight, bCancel
        If bCancel = False Then
            Do While GetKeyState(VK_LBUTTON) < 0
                ''''debug.print "moving..."
                GetCursorPos pt
                m_bSizing = True
                
                Screen.MousePointer = 4 ' vbSizeAll
                If pt.x <> ptPrev.x Or pt.Y <> ptPrev.Y Then
                    ptPrev.x = pt.x
                    ptPrev.Y = pt.Y
                    ' erase the previous drag rectangle if any
                    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                    DoEvents
                    'Tell the user we've moved
                    lObjLeft = pt.x - lXOffset
                    lObjTop = pt.Y - lYOffset
                    DockForm_MouseMove lObjLeft, lObjTop, lObjWidth, lObjHeight, lBorderWidth
                    ' Adjust the height/width
                    With DragRect
                        .Left = lObjLeft
                        .Top = lObjTop
                        .Right = .Left + lObjWidth
                        .Bottom = .Top + lObjHeight
                    End With
                    DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
                    bMoved = True
                    lXOffset = pt.x - DragRect.Left
                    lYOffset = pt.Y - DragRect.Top
                End If
                DoEvents
            Loop
            Screen.MousePointer = vbDefault
        End If
        ' erase the previous drag rectangle if any
        DrawDragRectangle DragRect.Left, DragRect.Top, DragRect.Right, DragRect.Bottom, lBorderWidth
        ' move and repaint the window
    End Select
    If bMoved Then
        ' If m_RepositionForm Then
        MoveWindow hwnd, DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top, True
        ' End If
        ' tell the user we've dropped the form
        DockForm_MouseUp DragRect.Left, DragRect.Top, DragRect.Right - DragRect.Left, DragRect.Bottom - DragRect.Top
        DeleteDC 0
        'debug.print "We Have Moved"
        If Not Panel.Expanded Then
            If m_bSizing Then
                Panel.Panel_Expand
            End If
        End If
        DeleteDC 0
    End If

    m_bSizing = False
    
End Sub

Private Sub DrawDragRectangle(ByVal x As Long, _
                              ByVal Y As Long, _
                              ByVal X1 As Long, _
                              ByVal Y1 As Long, _
                              ByVal lWidth As Long)

  Dim DragBox(3) As RECT

    ' ******************************************************************************
    ' Routine       : (Sub) DrawDragRectangle
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:40:49
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Credits       : This rotine was extract from the DockForm_Dragger project
    '                 by Steve (Steve81@ mediaone.net)
    '                 http://www.vbtutor.com
    '                 Thanks Steve!
    ' Modifications : Changed the pen color to vbButtonFace (that's all)
    '
    ' Description   : Draw a rectangle using the Win32 API
    ' ******************************************************************************
    DragBox(0).Bottom = Y1
    DragBox(0).Top = Y
    DragBox(0).Left = x
    DragBox(0).Right = DragBox(0).Left + 5
    pDrawSplitter DragBox(0)
    DoEvents
    DragBox(1).Top = Y
    DragBox(1).Bottom = DragBox(1).Top + 5
    DragBox(1).Left = x + 5
    DragBox(1).Right = X1
    pDrawSplitter DragBox(1)
    DoEvents
    DragBox(2).Bottom = Y1
    DragBox(2).Top = DragBox(2).Bottom - 5
    DragBox(2).Left = x + 5
    DragBox(2).Right = X1
    pDrawSplitter DragBox(2)
    DoEvents
    DragBox(3).Bottom = Y1
    DragBox(3).Top = Y
    DragBox(3).Left = X1
    DragBox(3).Right = X1 + 5
    pDrawSplitter DragBox(3)
    DoEvents

End Sub

Friend Property Get Extender() As Form

    Set Extender = ObjectFromPtr(m_FormPtr)

End Property

Friend Property Set Extender(ByVal vData As Form)

    m_FormPtr = PtrFromObject(vData)
    Attach vData.hwnd

End Property

Public Property Get FloatingHeight() As Long
Attribute FloatingHeight.VB_Description = "Returns the floating height of the form"

    FloatingHeight = m_lFloatingHeight

End Property

Public Property Let FloatingHeight(ByVal vData As Long)

    m_lFloatingHeight = vData

End Property

Public Property Get FloatingLeft() As Long
Attribute FloatingLeft.VB_Description = "Returns the floating Left of the form"

    FloatingLeft = m_lFloatingLeft

End Property

Public Property Let FloatingLeft(ByVal vData As Long)

    m_lFloatingLeft = vData

End Property

Public Property Get FloatingTop() As Long
Attribute FloatingTop.VB_Description = "Returns the floating Top of the form"

    FloatingTop = m_lFloatingTop

End Property

Public Property Let FloatingTop(ByVal vData As Long)

    m_lFloatingTop = vData

End Property

Public Property Get FloatingWidth() As Long
Attribute FloatingWidth.VB_Description = "Returns the floating Width of the form"

    FloatingWidth = m_lFloatingWidth

End Property

Public Property Let FloatingWidth(ByVal vData As Long)

    m_lFloatingWidth = vData

End Property

Friend Property Get HasCloseButton() As Boolean

    HasCloseButton = m_bHasCloseButton

End Property

Friend Property Let HasCloseButton(ByVal vData As Boolean)

    m_bHasCloseButton = vData

End Property

Friend Property Get HasMaxButton() As Boolean

    HasMaxButton = m_bHasMaxButton

End Property

Friend Property Let HasMaxButton(ByVal vData As Boolean)

    m_bHasMaxButton = vData

End Property

Friend Property Get Height() As Variant

    Height = Extender.Height

End Property

Friend Property Let Height(ByVal vData As Variant)

    Extender.Height = vData

End Property

Public Function HoverCaption() As Boolean

  Dim RC      As RECT
  Dim hbr     As Long
  Dim pt      As POINTAPI
  Dim xLeft   As Long
  Dim xRight  As Long
  Dim xTop    As Long
  Dim xBottom As Long

    On Error Resume Next

        If Panel.Splitting Then Exit Function
        If Not Panel.Expanded Then Exit Function
        If m_bSizing Then Exit Function
        
        GetCursorPos pt
        GetWindowRect m_hWnd, RC
        xLeft = (pt.x) - (RC.Left + 1)
        xRight = (pt.x - 5) + 10
        xTop = ((pt.Y - 5))
        xBottom = (pt.Y - 5) + 16
        xTop = ((pt.Y)) - RC.Top
        m_bCloseDown = False
        m_bCloseOver = False
        m_bChevronDown = False
        m_bChevronOver = False
        m_bOverMax = False
        m_bOverClose = False
        If Not Panel.Parent.CaptionStyle = tdCaptionVSNet Then
            Exit Function
        End If

        'HoverCaption = False
        m_bCloseDown = True
        m_bCloseOver = True
        m_bChevronDown = False
        m_bChevronOver = True
      
        

        If Not m_bHasMaxButton And Not m_bHasCloseButton Then Exit Function

        If m_bHasMaxButton Then
            If xLeft >= MaxRec.Left And xLeft <= MaxRec.Right - 3 Then
                If xTop >= MaxRec.Top + 3 And xTop <= MaxRec.Bottom - 3 Then
                    'Debug.Print "Hit Max Button"
                    m_bOverMax = True
                    If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                        m_bChevronDown = False
                        m_bChevronOver = True
                        hbr = CreateSolidBrush(VSNetCheckedColor)
                        If m_bHasMaxButton Then
                            If Not m_bMaxDisabled Then
                                drawChevronButton Extender.hdc, hbr, MaxRec
                              Else
                                DeleteObject hbr
                                hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                                m_bChevronOver = False
                                m_bChevronDown = False

                                If m_bHasMaxButton Then
                                    drawChevronButton Extender.hdc, hbr, MaxRec, True, Me
                                End If
                            End If
                        End If
                        DeleteObject hbr

                        HoverCaption = True
                        m_bCloseDown = False
                        m_bCloseOver = False ''''
                        hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        '

                        If m_bHasCloseButton Then
                            drawCloseButton Extender.hdc, hbr, CloseRec
                        End If
                        m_bCloseDown = False
                        m_bCloseOver = False ''''
                        DeleteObject hbr
                    End If
                End If
              Else
                hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                If m_bHasMaxButton Then
                    If m_bHasCloseButton Then
                       ' drawCloseButton Extender.hdc, hbr, MaxRec
                        DeleteObject hbr
                    End If
                End If

                If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                    m_bChevronDown = False
                    m_bChevronOver = False
                    m_bCloseDown = False
                    m_bCloseOver = False
                    hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)

                    If Not m_bMaxDisabled Then
                        drawChevronButton Extender.hdc, hbr, MaxRec
                      Else
                        DeleteObject hbr
                        hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                        m_bChevronOver = False
                        m_bChevronDown = False
                        If m_bHasMaxButton Then
                            drawChevronButton Extender.hdc, hbr, MaxRec, True, Me
                        End If
                    End If

                    HoverCaption = True
                    If m_bHasCloseButton Then
                        drawCloseButton Extender.hdc, hbr, CloseRec
                        DeleteObject hbr
                    End If
                End If
            End If
        End If

        If m_bHasCloseButton Then

            If xLeft >= CloseRec.Left And xLeft <= CloseRec.Right - 3 Then
                If xTop >= CloseRec.Top + 1 And xTop <= CloseRec.Bottom - 2 Then
                    m_bOverClose = True
                    If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                        m_bCloseDown = False
                        m_bCloseOver = True
                        hbr = CreateSolidBrush(VSNetCheckedColor)
                        If m_bHasCloseButton Then
                            drawCloseButton Extender.hdc, hbr, CloseRec
                            DeleteObject hbr
                        End If

                        m_bChevronDown = False
                        m_bChevronOver = False

                        hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)

                        If m_bHasMaxButton Then
                            If Not m_bMaxDisabled Then
                                drawChevronButton Extender.hdc, hbr, MaxRec
                              Else
                                DeleteObject hbr
                                hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                                m_bChevronOver = False
                                m_bChevronDown = False
                                drawChevronButton Extender.hdc, hbr, MaxRec, True, Me
                            End If
                        End If
                        DeleteObject hbr
                    End If
                    HoverCaption = True
                    'Debug.Print "Hit close Button"
                End If
              Else
                If Panel.Parent.CaptionStyle = tdCaptionVSNet Then
                    m_bChevronDown = False
                    m_bChevronOver = False
                    m_bCloseDown = False
                    m_bCloseOver = False
                    hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                    HoverCaption = True
                    If m_bHasCloseButton Then
                        drawCloseButton Extender.hdc, hbr, CloseRec
                        DeleteObject hbr
                    End If
                End If
            End If

        End If

        m_bCloseDown = False
        m_bCloseOver = False
        m_bChevronDown = False
        m_bChevronOver = False

        If Not m_bOverMax And Not m_bOverClose Then
            hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)

            If m_bHasCloseButton Then
                drawCloseButton Extender.hdc, hbr, CloseRec
            End If

            If m_bHasMaxButton Then
                If Not m_bMaxDisabled Then
                    drawChevronButton Extender.hdc, hbr, MaxRec
                  Else
                    DeleteObject hbr
                    hbr = GetSysColorBrush(vbInactiveTitleBar And &H1F&)
                    m_bChevronOver = False
                    m_bChevronDown = False
                    If m_bHasMaxButton Then
                        If Not m_bMaxDisabled Then
                            drawChevronButton Extender.hdc, hbr, MaxRec, True, Me
                        End If
                    End If
                End If
            End If
            'Debug.Print "Not Over Panel Buttons"
            DeleteObject hbr
        End If
        DeleteObject hbr

    Exit Function

End Function

Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns the form handle. This may not be used in api callings once the docked form is already subclassed and another subclassing of this form may crash Vb IDE or unpredictable results may take place."

    hwnd = Extender.hwnd

End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse

  ' Process before windows:
  ''debug.print CurrentMessage

    Select Case CurrentMessage
      Case WM_SYSCOMMAND
        ISubClass_MsgResponse = emrConsume
      Case WM_CONTEXTMENU
        ISubClass_MsgResponse = emrConsume
      Case WM_NCRBUTTONDOWN
        ISubClass_MsgResponse = emrConsume
        '        Case WM_NCLBUTTONDOWN
        '            ISubclass_MsgResponse = emrPreprocess
        '        Case WM_ENTERSIZEMOVE
        '            ISubclass_MsgResponse = emrPostProcess
        '        Case WM_EXITSIZEMOVE
        '            ISubclass_MsgResponse = emrConsume
      Case WM_SIZE
        ISubClass_MsgResponse = emrPreprocess
      Case WM_MOVE
        ISubClass_MsgResponse = emrConsume
      Case WM_NCLBUTTONDBLCLK
        ISubClass_MsgResponse = emrConsume
      Case WM_PAINT
        ISubClass_MsgResponse = emrPreprocess
      Case WM_NCHITTEST
        ISubClass_MsgResponse = emrPreprocess
      Case Else
        ISubClass_MsgResponse = emrPreprocess
    End Select

End Property

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)

  '

End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

  Dim uCmdType     As Long
  Static iRefCount As Long
  Dim retVal       As Long

    ' ******************************************************************************
    ' Routine       : ISubClass_WindowProc
    ' Created by    : Marclei V Silva
    ' Machine       : ZEUS
    ' Date-Time     : 28/08/005:30:36
    ' Inputs        :
    ' Outputs       :
    ' Credits       :
    ' Modifications :
    ' Description   : Window messages handler for the form
    ' ******************************************************************************
    Select Case iMsg
      Case WM_NCHITTEST
        ''debug.print "nchittest", lParam, wParam
        ''debug.print iMsg
        'With m_tHTR
        '*******************
        'GetCursorPos tP
        'ScreenToClient hwnd, tP
        '            eType = HTCLIENT
        'tR.Left = .lLeft
        'tR.Top = .lTOp
        'tR.Right = .lRight
        'tR.Bottom = .lBottom
        'If (PtInRect(tR, tP.x, tP.y) <> 0) Then
        retVal = CallOldWindowProc(hwnd, iMsg, wParam, lParam) 'HTCAPTION
        ''debug.print "retval="; retVal
       ' HoverCaption
        If retVal = 1 Then
            ISubClass_WindowProc = HTCAPTION
          Else
            ISubClass_WindowProc = retVal
        End If
        'MsgBox SC_CLOSE & "|" & uCmdType & "|" & PushDown
        'End With
      Case WM_SYSCOMMAND
        'Debug.Print "SYSCOMMAND"
        uCmdType = wParam And &HFFF0
        If uCmdType = SC_MOVE Then
            ReleaseCapture
            DragObject hwnd
          ElseIf uCmdType = SC_CLOSE Then
            m_bUnloaded = True
            ''debug.print "close"
            DockForm_Unload
            If Not Panel.Expanded Then
                Panel.RefreshCollaped
            End If
          Else
            ISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
        End If
        'MsgBox SC_CLOSE & "|" & uCmdType & "|" & PushDown
      Case WM_NCACTIVATE
        ''debug.print "WM_NCACTIVATE"
        'resetVSNetRecs
        If wParam = 0 Then
            iRefCount = iRefCount + 1
            If iRefCount < 2 Then
                '-- Removed To Stop Flickering
                '-- Gary Noble (Phantom Man(PSC))
                '   LockWindowUpdate ParentHWND
                ISubClass_WindowProc = CallOldWindowProc(ParentHWND, iMsg, wParam, lParam)
                'CallOldWindowProc m_hWnd, WM_NCACTIVATE, 1, 0
                'LockWindowUpdate 0
                '''debug.print "Unlocking"
                DockForm_Paint
                ''resetVSNetRecs
               ' HoverCaption
              Else
                ''debug.print "ncact 2"
                ISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
                DockForm_Paint
            End If
          Else
            ''debug.print "ncact 3"
            ISubClass_WindowProc = CallOldWindowProc(ParentHWND, iMsg, wParam, lParam)
        End If
      Case WM_NCLBUTTONDBLCLK
        ''debug.print "NCLBUTTONDBLCLK"
        If wParam = HTCAPTION Then
            DockForm_DblClick
        End If
      Case WM_STYLECHANGED
        If m_bStyleChanged = False Then
            DockForm_ResetStyle
        End If
      Case WM_ACTIVATE
        ' Form is activated/deactivated:
        '''debug.print "WM_ACTIVATE"
        If wParam = 0 Then
            '''debug.print "DEACT"
            iRefCount = 0
            m_bActive = False
            m_bRedraw = True
            DockForm_Deactivate
          Else
            '''debug.print "ACT"
            m_bActive = True
            DockForm_Activate wParam = 2
        End If
      Case WM_ACTIVATEAPP
        ''debug.print "ACTIVATEAPP"
        If (wParam = 0) Then
            iRefCount = 0
            ' app being deactivated
            CallOldWindowProc m_hWnd, WM_NCACTIVATE, 0, 0

            DockForm_Paint
          Else
            ' app being activated
            ' if not the active form then we should repaint
            ' the title bar
            CallOldWindowProc m_hWnd, WM_NCACTIVATE, 1, 0
            DockForm_Paint
        End If
      Case WM_NCLBUTTONDOWN
        ' 'debug.print "ggggg"
        ' DockForm_CaptionClose
      Case WM_NCRBUTTONDOWN
        '''debug.print "NCRBUTTONDOWN"
        If (wParam = HTCAPTION) Then
            DockForm_CaptionClick vbRightButton, LoWord(lParam), HiWord(lParam)
        End If
      Case WM_CONTEXTMENU
        '''debug.print "WM_CONTEXTMENU"
      Case WM_CLOSE
        '''debug.print "WM_CLOSE"
      Case WM_DESTROY
        '''debug.print "WM_DESTROY"
        Detach
      Case WM_PAINT
        '''debug.print "WM_PAINT"
        RefreshVSNet
        DockForm_Paint
      Case WM_NCPAINT
        If Not m_SizingNow And m_State = tdUndocked Then
            RefreshVSNet
            DockForm_Paint
        End If
      Case WM_MOVE
        ''debug.print "WM_MOVE"
        'DockForm_Paint
        'resetVSNetRecs
      '  HoverCaption
      Case WM_SIZE
        ''debug.print "WM_SIZE"
        RefreshVSNet

        If Me.State = tdUndocked And m_bDragging = False Then
            Me.FloatingWidth = (Me.width \ Screen.TwipsPerPixelX)
            Me.FloatingHeight = (Me.Height \ Screen.TwipsPerPixelY)
        End If
        'DockForm_Paint
      Case WM_NCMOUSEMOVE
        'debug.print "mousemove", Now
      Case Else
        ''debug.print "imsg"; iMsg
        'resetVSNetRecs
    End Select

End Function

Public Property Get Key() As String
Attribute Key.VB_Description = "Returns the key of the form to be used with TabDock functions like FormShow() or FormHide. Instead of the key you may use the Form index to retrieve form reference"

    Key = m_Key

End Property

Public Property Let Key(ByVal vData As String)

    m_Key = vData

End Property

Friend Property Get Left() As Variant

    Left = Extender.Left

End Property

Friend Property Let Left(ByVal vData As Variant)

    Extender.Left = vData

End Property

Private Sub m_Sizer_EnterSizeMove()

    m_SizingNow = True

End Sub

Private Sub m_Sizer_ExitSizeMove()

    m_SizingNow = False

End Sub

Friend Function Move(Left, _
       Optional Top, _
       Optional width, _
       Optional Height)

    On Error Resume Next
        Extender.Move Left, Top, width, Height
    On Error GoTo 0

End Function

Private Sub mTimer_ThatTime()

    On Error Resume Next

        HoverCaption

    On Error GoTo 0

End Sub

Public Property Get Object() As TDockForm

    Set Object = Me

End Property

Public Property Get Panel() As TTabDockHost
Attribute Panel.VB_Description = "Returns the panel that is responsible for the form. This property changes when you dock the form into a different panel."

    Set Panel = ObjectFromPtr(m_PanelPtr)

End Property

Friend Property Set Panel(ByVal vData As TTabDockHost)

    m_PanelPtr = PtrFromObject(vData)

End Property

Public Property Get ParentHWND() As Long

    ParentHWND = StoreHWND

End Property

Public Property Let ParentHWND(ByVal hwnd As Long)

    StoreHWND = hwnd

End Property

Private Sub pDrawSplitter(oRect As RECT)

  Dim lHDC      As Long
  Dim hOldBrush As Long

    lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    hOldBrush = SelectObject(lHDC, m_hBrush)
    PatBlt lHDC, oRect.Left, oRect.Top, oRect.Right - oRect.Left, oRect.Bottom - oRect.Top, PATINVERT
    SelectObject lHDC, hOldBrush
    DeleteDC lHDC

End Sub

Public Property Get Percent() As Integer

    Percent = m_Percent

End Property

Public Property Let Percent(ByVal vData As Integer)

    m_Percent = vData

End Property

Public Property Get PointInObject(ByRef oThis As Object, _
                                  ByVal xPixels As Long, _
                                  ByVal yPixels As Long) As Boolean

  Dim tR As RECT

    On Error Resume Next
        ' Get object position:
        tR.Left = oThis.Left \ Screen.TwipsPerPixelX
        tR.Right = tR.Left + oThis.width \ Screen.TwipsPerPixelX
        tR.Top = oThis.Top \ Screen.TwipsPerPixelX
        tR.Bottom = tR.Top + oThis.Height \ Screen.TwipsPerPixelY
        If (Err.Number = 0) Then
            ' Check for hit test:
            PointInObject = (PtInRect(tR, xPixels, yPixels) <> 0)
        End If
    On Error GoTo 0

End Property

Friend Property Get Redraw() As Boolean

    Redraw = m_bRedraw

End Property

Friend Property Let Redraw(ByVal vData As Boolean)

    m_bRedraw = vData

End Property

Public Function RefreshVSNet()

  Dim tR As RECT

    Call GetWindowRect(Me.hwnd, tR)
  Dim hbr As Long

    If Panel.Parent.CaptionStyle = tdCaptionVSNet Then

        If State = tdUndocked Then

            If m_bActive Then
                DrawBorderUnDocked Extender, TranslateColor(vbActiveTitleBar)
              Else
                DrawBorderUnDocked Extender, TranslateColor(vbInactiveTitleBar)
            End If

        End If

    End If

End Function

Public Function resetVSNetRecs()

    With MaxRec
        .Left = 0
        .Right = 0
        .Top = 0
        .Bottom = 0
    End With
    With CloseRec
        .Left = 0
        .Right = 0
        .Top = 0
        .Bottom = 0
    End With
    DockForm_Paint

End Function

Public Sub SetTestArea(ByVal vKey As Variant, _
                       ByVal lLeftPixels As Long, _
                       ByVal lTopPixels As Long, _
                       ByVal lWidthPixels As Long, _
                       ByVal lHeightPixels As Long, _
                       ByVal eResponse As EHitTestAreas)

  Dim iItem As Long

    '**************************************
    iItem = AreaIndex(vKey)
    If (iItem > 0) Then
        With m_tHTR
            .lLeft = lLeftPixels
            .lTop = lTopPixels
            .lRight = lLeftPixels + lWidthPixels
            .lBottom = lTopPixels + lHeightPixels
            .eResponse = eResponse
        End With
    End If

End Sub

Public Sub SetTestAreaFromObject(ByVal vKey As Variant, _
                                 ByRef objThis As Object, _
                                 ByVal eResponse As EHitTestAreas)

    SetTestArea vKey, objThis.Left \ Screen.TwipsPerPixelX, objThis.Top \ Screen.TwipsPerPixelY, objThis.width \ Screen.TwipsPerPixelY, objThis.Height \ Screen.TwipsPerPixelY, eResponse

End Sub

Public Property Let State(ByVal vData As tdDockedState)
Attribute State.VB_Description = "Returns the state of the form. Docked or Undocked."

    m_State = vData

End Property

Public Property Get State() As tdDockedState

    State = m_State

End Property

Public Property Let Style(ByVal vData As tdDockStyles)
Attribute Style.VB_Description = "Returns or sets the form style. The style may allow or avoid some features of a docked window. Check all the possible styles to choose the apropriate one."

    m_Style = vData

End Property

Public Property Get Style() As tdDockStyles

    Style = m_Style

End Property

Friend Property Let Tag(ByVal Value As Boolean)

    m_Tag = Value

End Property

Friend Property Get Tag() As Boolean

    Tag = m_Tag

End Property

Friend Property Let Top(ByVal vData As Variant)

    Extender.Top = vData

End Property

Friend Property Get Top() As Variant

    Top = Extender.Top

End Property

Friend Property Let Visible(ByVal vData As Boolean)

    Extender.Visible = vData
    If vData Then
        ' trigger TabDock.FormShow() Event
        Panel.Parent.TriggerEvent "ShowForm", Object
      Else
        ' trigger TabDock.Formhide() Event
        Panel.Parent.TriggerEvent "HideForm", Object
    End If

End Property

Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Returns wether the form is visible or not"

    Visible = Extender.Visible

End Property

Friend Property Let width(ByVal vData As Variant)

    Extender.width = vData

End Property

Friend Property Get width() As Variant

    width = Extender.width

End Property

