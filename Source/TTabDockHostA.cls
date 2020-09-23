VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TTabDockHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************************************
' Project     :  TabDock
' Module      :  TTabDockHost.cls
' Description :  Host for Docked Forms
' Created by  :  Marclei V Silva
' Machine     :  ZEUS
' Date-Time   :  12/05/2000 0:40:08
' ********************************************************************
Option Explicit
' Keep up with the errors
Private Const g_ErrConstant               As Long = vbObjectError + 1000
Private Const m_constClassName            As String = "TTabDockHost"
Private m_lngErrNum                       As Long
Private m_strErrStr                       As String
Private m_strErrSource                    As String

Private m_iFormIndex                      As Integer

' cursor enumeration from the resource file
Private Enum EnumCursors
    tdSizeWE = 103
    tdSizeNS = 102
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private tdSizeWE, tdSizeNS
#End If
' Splitters Size/Height
Private Const SPLITTER_HEIGHT             As Integer = 80
Private Const SPLITTER_WIDTH              As Integer = 80
' Minimum allowed sizes for any PictureBox.
Private Const MIN_HORIZONTAL              As Integer = 50
Private Const MIN_VERTICAL                As Integer = 50
' Drag directions.
Private Const DRAG_NONE                   As Integer = 0
Private Const DRAG_HORIZONTAL             As Integer = 1
Private Const DRAG_VERTICAL               As Integer = 2
Private Const DRAG_BOTH                   As Integer = 3
Private m_bSplitting                      As Boolean        ' Flag that indicates the
Private m_bDirty                          As Boolean        ' Flag that indicates the
' docked forms must be resized
Private DragDir                           As Integer        ' dragging direction
Private cHS                               As cSplitDDC      ' horizontal and
Private cVS                               As cSplitDDC      ' vertical splitters
Private deltaY                            As Integer        ' X offset when splitting
Private deltaX                            As Integer        ' Y offset when splitting
Private m_FormIndex                       As Integer        ' Index of the form over form drag area
Private m_WindowList                      As TDockForms     ' forms collection
Private m_Index                           As Integer
Private m_Resizable                       As Integer
Private m_ParentPtr                       As Long
Private m_Align                           As tdAlignProperty
Private m_AutoShowCaptionOnCollapse       As Boolean
' --- Amended By NR 10/04/2003
Private m_Expanded                        As Boolean
Private m_Size                            As Long
Private m_SplitWidth                      As Long
Private m_SplitHeight                     As Long
Private m_Redraw                          As Boolean
' ---- Amend end NR 10/04/2003
'-- Ammended
'-- Added Disable Close Button If The Dock Host Is NoCloseButton
'-- By: Gary Noble (Phantom Man(PSC))
Private Const xSC_CLOSE                   As Long = -10&
Private Const WS_EX_APPWINDOW             As Long = &H40000
Private Const MIIM_STATE                  As Long = &H1&
Private Const MIIM_ID                     As Long = &H2&
Private Const MFS_GRAYED                  As Long = &H3&
Private Const WM_NCACTIVATE               As Long = &H86
Private WithEvents Extender               As VB.PictureBox  ' Used as the host
Attribute Extender.VB_VarHelpID = -1
Private Type MENUITEMINFO
    cbSize                                    As Long
    fMask                                     As Long
    fType                                     As Long
    fState                                    As Long
    wID                                       As Long
    hSubMenu                                  As Long
    hbmpChecked                               As Long
    hbmpUnchecked                             As Long
    dwItemData                                As Long
    dwTypeData                                As String
    cch                                       As Long
End Type
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, _
                          ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
                          ByVal un As Long, _
                          ByVal b As Boolean, _
                          lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, _
                          ByVal un As Long, _
                          ByVal bool As Boolean, _
                          lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                          ByVal wMsg As Long, _
                          ByVal wParam As Long, _
                          lParam As Any) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                          ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                          ByVal nIndex As Long, _
                          ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                          ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                          ByVal hWndInsertAfter As Long, _
                          ByVal x As Long, _
                          ByVal Y As Long, _
                          ByVal cx As Long, _
                          ByVal cy As Long, _
                          ByVal wFlags As Long) As Long
'-- End Ammend
Private DummyToKeepDecCommentsInDEclarations As Boolean

Private m_bDockedFormSizing As Boolean
Private m_bPanelSizing As Boolean

Public Property Let DockedFormSizing(bln As Boolean)

    m_bDockedFormSizing = bln

End Property

Public Property Let PanelSizing(bln As Boolean)

    m_bPanelSizing = bln

End Property


Public Property Let Align(New_Align As tdAlignProperty)

    Extender.Align = New_Align
    m_Align = New_Align

End Property

Public Property Get Align() As tdAlignProperty

    Align = m_Align

End Property

Friend Function AllowDocking(ByVal Df As TDockForm) As Boolean

  Const constSource As String = m_constClassName & ".AllowDocking"

    ' ******************************************************************************
    ' Routine       : (Function) AllowDocking
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 6:50:30
    ' Inputs        : Byval df: TDockForm
    ' Outputs       :
    ' Modifications :
    ' Description   : Returns true if the specified form is allowed to dock
    ' ******************************************************************************
    On Error GoTo Err_AllowDocking
    Select Case Align
      Case tdAlignLeft
        AllowDocking = Df.Style And tdDockLeft
      Case tdAlignRight
        AllowDocking = Df.Style And tdDockRight
      Case tdAlignTop
        AllowDocking = Df.Style And tdDockTop
      Case tdAlignBottom
        AllowDocking = Df.Style And tdDockBottom
    End Select

Exit Function

Err_AllowDocking:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Function

Public Property Let AutoShowCaptionOnCollapse(Auto As Boolean)

    m_AutoShowCaptionOnCollapse = Auto

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = Extender.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    Extender.BackColor() = New_BackColor

End Property

Public Property Get BorderStyle() As Integer

    BorderStyle = Extender.BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)

    Extender.BorderStyle() = New_BorderStyle

End Property

Private Sub Class_Initialize()

  Const constSource As String = m_constClassName & ".Class_Initialize"

    ' ******************************************************************************
    ' Routine       : (Sub) Class_Initialize
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:35:31
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : Initialize the class
    ' ******************************************************************************
    On Error GoTo Err_Class_Initialize
    ' initially it is resizable
    m_Resizable = True
    ' configure splitters
    Set cHS = New cSplitDDC
    cHS.Orientation = espVertical
    Set cVS = New cSplitDDC
    cVS.Orientation = espHorizontal
    Set m_WindowList = New TDockForms
    m_Expanded = True
    m_Redraw = True

Exit Sub

Err_Class_Initialize:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub Class_Terminate()

  Const constSource As String = m_constClassName & ".Class_Terminate"

    ' ******************************************************************************
    ' Routine       : (Sub) Class_Terminate
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:35:21
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : Terminate the class
    ' ******************************************************************************
    On Error GoTo Err_Class_Terminate
    Set cHS = Nothing
    Set cVS = Nothing
    Set m_WindowList = Nothing

Exit Sub

Err_Class_Terminate:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Function ClassName(ByVal lhWnd As Long) As String

  Dim lLen As Long
  Dim sBuf As String

    ' ******************************************************************************
    ' Routine       : (Function) ClassName
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:14:30
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : Returns the Classname
    ' ******************************************************************************
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lhWnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left(sBuf, lLen)
    End If

End Function

Private Sub ClipHostRect(ByRef RC As RECT)

  Dim rcPanel As RECT

    ' restrict resizing over other Panels
    If Align = tdAlignRight Then
        Parent.Panels(tdAlignLeft).GetHostRect rcPanel
        RC.Left = rcPanel.Right + MIN_HORIZONTAL
      ElseIf Align = tdAlignLeft Then
        Parent.Panels(tdAlignRight).GetHostRect rcPanel
        RC.Right = rcPanel.Left - MIN_HORIZONTAL
      ElseIf Align = tdAlignTop Then
        Parent.Panels(tdAlignBottom).GetHostRect rcPanel
        RC.Bottom = rcPanel.Top - MIN_VERTICAL
      ElseIf Align = tdAlignBottom Then
        Parent.Panels(tdAlignTop).GetHostRect rcPanel
        RC.Top = rcPanel.Bottom + MIN_VERTICAL
    End If

End Sub

Public Property Set Container(NewValue As Object)

    Set Extender = NewValue

End Property

Public Property Get Container() As Object

    Set Container = Extender

End Property

Public Sub Dock(ByVal Df As TDockForm)

  Const constSource As String = m_constClassName & ".Dock"
  Dim hWndA         As Long
  Dim Style         As Long
  Dim idx           As Integer
    
    m_iFormIndex = -1
    ' ******************************************************************************
    ' Routine       : (Sub) Dock
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 09/06/2000 - 14:57:02
    ' Inputs        : df: TDockForm
    ' Outputs       :
    ' Modifications :
    ' Credits       : All credits for this code to VbWeb resources:
    '
    ' Description   : Dock a form within host boundaries
    ' ******************************************************************************
    On Error GoTo Err_Dock
    ''debug.print "Dock"

    Extender.AutoRedraw = True
    Extender.Cls
    ' LockWindowUpdate Extender.Parent.hwnd
    ' change the dock form state
    Df.State = tdDocked
    Df.Redraw = True
    ' get the form handle
    hWndA = Df.hwnd
    ' check if the form may dock here
    If AllowDocking(Df) = False Then
        ' if not just show the form and good bye
        Call ShowWindow(hWndA, SW_SHOW)
        Exit Sub
    End If
    ' We have to rebuild window list order
    If m_WindowList.HandleExists(hWndA) Then
        m_WindowList.RemoveByHandle hWndA
    End If
    ' get a new index for this window
    idx = GetFormIndex(Df.FloatingLeft, Df.FloatingTop)
    ' add this form to the collection
    m_WindowList.Append Df, idx
    ' This is necessary in order to make top and bottom
    ' panels to align after/before any
    ' aligned controls on your form
    If Extender.Align = tdAlignTop Then
        Top = 1000
    End If
    If Extender.Align = tdAlignBottom Then
        Top = 0
    End If
    ' hide window
    Call ShowWindow(hWndA, SW_HIDE)
    Df.Left = Df.Left + 500
    ' get the form style
    Style = GetWindowLong(hWndA, GWL_STYLE)
    ' set new style for the form
    Style = Style Or WS_DLGFRAME            ' have dialog frame
    Style = Style Or WS_SYSMENU             ' have sysmenu close button
    Style = Style Or WS_OVERLAPPED
    Style = Style Or WS_CLIPCHILDREN
    Style = Style Or WS_CLIPSIBLINGS
    Style = Style Or WS_OVERLAPPED
    Style = Style And Not WS_MAXIMIZE       ' No maximize here
    Style = Style And Not WS_MINIMIZE       ' No minimize too
    Style = Style Xor WS_MAXIMIZEBOX        ' No maximize box
    Style = Style Xor WS_MINIMIZEBOX        ' No minimize box
    Style = Style And Not WS_THICKFRAME     ' No sizing frame
    Style = Style And Not WS_SYSMENU        ' No sizing frame
    If Extender.Align = tdAlignTop Or Extender.Align = tdAlignBottom Then
        Style = Style And Not WS_CAPTION    ' no caption
        Df.Extender.BorderStyle = 0
      Else
        Style = Style And Not WS_CAPTION    ' no caption
        Df.Extender.BorderStyle = 0
        'Df.Extender.Refresh
    End If
    ' set the style
    SetWindowLong hWndA, GWL_STYLE, Style
    ' this must be a toolwindow
    Style = GetWindowLong(hWndA, GWL_EXSTYLE)
    Style = Style Or WS_EX_TOOLWINDOW
    Style = Style And Not (WS_EX_APPWINDOW)
    Df.Left = Df.Left '+ 5000
    SetWindowLong hWndA, GWL_EXSTYLE, Style
    SetWindowPos hWndA, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
    ' set the parent to the dock host
    Call SetParent(hWndA, hwnd)
    ' show the container once we have a new form
    Extender.Visible = True
    ' show docked window
    If m_Expanded Then
        Call ShowWindow(hWndA, SW_SHOW)
        ' activate it please!
        ''debug.print "dock ncact"
        'Call SendMessage(hWndA, WM_NCACTIVATE, 1, 0)
        ' change the dock form state
        'Df.State = tdDocked
        ' chage visibility
        Df.Visible = True
        ' Call this procedure to raise a TabDock event
        'Parent.TriggerEvent "Dock", Df
        ' needs to re-calc heights then set this var to True
        m_bDirty = True
        ' repaint the control
        Extender_Paint
        ' style and update may take place now
        Df.Redraw = True
        '    LockWindowUpdate ByVal 0&
        ''debug.print "end docking"
        'SendMessage Df.hwnd, WM_SIZE, 0, 130
        Df.Tag = True
        Df.Visible = True
        Df.Extender.Height = Df.Extender.Height + 10
        'Df.Extender.height = Df.Extender.height - 10
        Df.Extender.width = Df.Extender.width + 10
        'Df.Extender.width = Df.Extender.width - 10
        ' Call this procedure to raise a TabDock event
        Parent.TriggerEvent "Dock", Df
       
        
      Else
        RefreshCollaped
    End If

Exit Sub

Err_Dock:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Sub DockArrange(Optional formIndex As Integer = -1, _
                       Optional deltaOffset As Integer = 0)

  Const constSource As String = m_constClassName & ".DockArrange"
  Dim lHeight       As Single
  Dim lWidth        As Single
  Dim lTop          As Single
  Dim lLeft         As Single
  Dim i             As Integer
  Dim CurrentH      As Long
  Dim NextH         As Long
  Dim Count         As Integer

    ' ******************************************************************************
    ' Routine       : (Sub) DockArrange
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:15:59
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : Reorder the visible forms within this host
    ' ******************************************************************************
    On Error GoTo Err_DockArrange
    
    m_iFormIndex = formIndex
    
    If formIndex > -1 Then
        m_FormIndex = formIndex
    End If
    If deltaOffset <> 0 Then
        deltaY = deltaOffset
    End If
    Count = m_WindowList.VisibleCount
    ' no controls (forms) then exit
    If Count = 0 Then
        Exit Sub
    End If
    ' reset position
    lLeft = 0
    lTop = 0
    ' based on align property, re-ordering will take different actions
    If Align = tdAlignLeft Or Align = tdAlignRight Then
        lWidth = Extender.ScaleWidth - SPLITTER_WIDTH
        If Align = tdAlignRight Then
            lLeft = lLeft + SPLITTER_WIDTH
        End If
      ElseIf Align = tdAlignTop Or Align = tdAlignBottom Then
        lHeight = Extender.ScaleHeight - SPLITTER_HEIGHT
        If Align = tdAlignBottom Then
            lTop = lTop + SPLITTER_HEIGHT
        End If
    End If
    ' loop the form array to draw each control in its proper
    ' position
    For i = 1 To m_WindowList.Count
        If m_WindowList(i).Visible Then

            If Align = tdAlignLeft Or Align = tdAlignRight Then
                ' if this form is a splitform then resize it
                If i = m_FormIndex Then
                    ' current height
                    CurrentH = (m_WindowList(i).Height + deltaY)
                    ' height of the next form
                    NextH = (m_WindowList(i + 1).Height - deltaY)
                    ' clear split form
                    m_FormIndex = 0
                  ElseIf NextH > 0 Then
                    ' update next form height
                    CurrentH = NextH
                    ' reset next height
                    NextH = 0
                  Else
                    ' update current height
                    CurrentH = m_WindowList(i).Height
                End If
                ' update properties for moving the form
                lHeight = CurrentH

                m_WindowList(i).Move lLeft, lTop, lWidth, lHeight
                m_WindowList(i).Extender.Refresh

                ' set next top position
                lTop = lTop + lHeight + SPLITTER_HEIGHT
              Else
                ' if this form is a splitform then resize it
                If i = m_FormIndex Then
                    ' current height
                    ''debug.print deltaY
                    CurrentH = (m_WindowList(i).width + deltaY)
                    If CurrentH < (getCaptionHeight + 3) * Screen.TwipsPerPixelX Then
                        CurrentH = (getCaptionHeight + 3) * Screen.TwipsPerPixelX
                    End If
                    ' height of the next form
                    NextH = (m_WindowList(i + 1).width - deltaY)
                    ' clear split form
                    m_FormIndex = 0
                  ElseIf NextH > 0 Then
                    ' update next form height
                    CurrentH = NextH
                    ' reset next height
                    NextH = 0
                  Else
                    ' update current height
                    CurrentH = m_WindowList(i).width
                End If
                ' update properties for moving the form
                lWidth = CurrentH
                ''debug.print lWidth
                m_WindowList(i).Move lLeft, lTop, lWidth, lHeight
                m_WindowList(i).Extender.Refresh
                ' set next top position
                lLeft = lLeft + lWidth + SPLITTER_HEIGHT
            End If
        End If
    Next '  I I

Exit Sub

Err_DockArrange:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Sub DockCollapse(formIndex As Integer)

  Const constSource       As String = m_constClassName & ".DockArrange"
  Dim lHeight             As Single
  Dim lWidth              As Single
  Dim lTop                As Single
  Dim lLeft               As Single
  Dim i                   As Integer
  Dim CurrentH            As Long
  Dim Count               As Integer
  Dim lMaxWidth           As Integer
  Dim lMaxHeight          As Integer
  Dim lCaptionHeightTwips As Integer

  
    On Error GoTo Err_DockArrange
    
    If m_iFormIndex = formIndex Then Exit Sub
    m_iFormIndex = formIndex
    
    Count = m_WindowList.VisibleCount
    ' no controls (forms) then exit
    If Count <= 1 Then
        Exit Sub
    End If
    ' reset position
    If Align = tdAlignBottom Or Align = tdAlignTop Then
        lCaptionHeightTwips = (getCaptionHeight + 5) * Screen.TwipsPerPixelX ' GetSystemMetrics(SM_CYMIN) * Screen.TwipsPerPixelX '(getCaptionHeight + 3) * Screen.TwipsPerPixelX
      Else
        lCaptionHeightTwips = (getCaptionHeight + 5) * Screen.TwipsPerPixelX 'GetSystemMetrics(SM_CYMIN) * Screen.TwipsPerPixelX
        'lCaptionHeightTwips = lCaptionHeightTwips + 0#
        'debug.print "Collapse Top"
    End If
    lLeft = 0
    lTop = 0
    ' based on align property, re-ordering will take different actions
    If Align = tdAlignLeft Or Align = tdAlignRight Then
        lWidth = Extender.ScaleWidth - SPLITTER_WIDTH
        If Align = tdAlignRight Then
            lLeft = lLeft + SPLITTER_WIDTH
        End If
        ''debug.print GetSystemMetrics(SM_CXMIN), GetSystemMetrics(SM_CYMIN)
        lMaxHeight = Extender.ScaleHeight - ((lCaptionHeightTwips + SPLITTER_WIDTH) * (Count - 1))
      ElseIf Align = tdAlignTop Or Align = tdAlignBottom Then
        lHeight = Extender.ScaleHeight - SPLITTER_HEIGHT
        If Align = tdAlignBottom Then
            lTop = lTop + SPLITTER_HEIGHT
        End If
        ''debug.print "dockcoll"; Count; Me.width
        lMaxWidth = Extender.ScaleWidth - ((lCaptionHeightTwips + SPLITTER_WIDTH) * (Count - 1))
    End If
    ' loop the form array to draw each control in its proper
    ' position
    For i = 1 To m_WindowList.Count
        If m_WindowList(i).Visible Then
            If Align = tdAlignLeft Or Align = tdAlignRight Then
                ' if this form is a splitform then resize it
                If i = formIndex Then
                    CurrentH = lMaxHeight
                    ' clear split form
                    formIndex = 0
                  Else
                    ' update current height
                    CurrentH = lCaptionHeightTwips
                End If
                ' update properties for moving the form
                lHeight = CurrentH
                m_WindowList(i).Move lLeft, lTop, lWidth, lHeight

                m_WindowList(i).Extender.Refresh
                ' set next top position
                lTop = lTop + lHeight + SPLITTER_HEIGHT
              Else
                ' if this form is a splitform then resize it
                If i = formIndex Then
                    CurrentH = lMaxWidth '(m_WindowList(i).width + deltaY)
                    ' clear split form
                    formIndex = 0
                  Else
                    ' update current height
                    CurrentH = lCaptionHeightTwips
                End If
                ' update properties for moving the form
                lWidth = CurrentH
                m_WindowList(i).Move lLeft, lTop, lWidth, lHeight
                m_WindowList(i).Extender.Refresh
                ' set next top position
                lLeft = lLeft + lWidth + SPLITTER_HEIGHT
            End If
        End If
    Next '  I I

Exit Sub

Err_DockArrange:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub DockHide(hWndA As Long)

  Const constSource As String = m_constClassName & ".DockHide"

    ' ******************************************************************************
    ' Routine       : (Sub) DockHide
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:08:24
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    '
    ' Description   : Simply hide a docked form and re-DockArrange the host
    ' ******************************************************************************
    On Error GoTo Err_DockHide
    ' hide window
    Call ShowWindow(hWndA, SW_HIDE)
    Extender_Paint

Exit Sub

Err_DockHide:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Sub DockResize()

  Const constSource As String = m_constClassName & ".DockResize"
  Dim lHeight       As Single
  Dim lWidth        As Single
  Dim lTop          As Single
  Dim lLeft         As Single
  Dim i             As Integer
  Dim Count         As Integer
  Dim lTotalPercent As Long
  Dim lMaxHeight    As Long
  Dim lMaxWidth     As Long

    ' ******************************************************************************
    ' Routine       : (Sub) DockResize
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:21:50
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    '
    ' Description   : Used to DockArrange all the docked forms when a resize
    '                 method is performed
    ' ******************************************************************************
    On Error GoTo Err_DockResize
    If Not m_Expanded Then
        Exit Sub
    End If
    Extender.AutoRedraw = True
    Count = m_WindowList.VisibleCount
    ' no controls then exit
    If Count = 0 Then
        Exit Sub
    End If
    ' update coordinates based on align method of the host
    lLeft = 0
    lTop = 0
    lHeight = Extender.ScaleHeight
    ''debug.print "Align "; Align
    ' determine maximum percentage for panel
    For i = 1 To WindowList.Count
        If m_WindowList(i).Visible Then
            lTotalPercent = lTotalPercent + m_WindowList(i).Percent
        End If
    Next i
    If Align = tdAlignLeft Or Align = tdAlignRight Then
        lWidth = Extender.ScaleWidth - SPLITTER_WIDTH
        lMaxHeight = ((lHeight - ((Count - 1) * SPLITTER_HEIGHT)))
        If Align = tdAlignRight Then
            lLeft = lLeft + SPLITTER_WIDTH
        End If
      ElseIf Align = tdAlignTop Or Align = tdAlignBottom Then
        lWidth = Extender.ScaleWidth
        lMaxHeight = (lHeight - SPLITTER_HEIGHT)
        lMaxWidth = (lWidth - ((Count - 1) * SPLITTER_WIDTH)) / Count
        If Align = tdAlignBottom Then
            lTop = lTop + SPLITTER_HEIGHT
        End If
    End If
    ' this is necessary because i don't wanna check if the parent
    ' form is minimize or not
    On Error Resume Next
        For i = 1 To WindowList.Count
            If m_WindowList(i).Visible Then
                If Align = tdAlignTop Or Align = tdAlignBottom Then
                    ' height of window = percent / total percent * maxheight
                    lHeight = lMaxHeight
                    m_WindowList(i).Move lLeft, lTop, lMaxWidth, lHeight
                   'debug.Print lLeft
                    lLeft = lLeft + lMaxWidth + SPLITTER_WIDTH
                  Else
                    ' height of window = percent / total percent * maxheight
                    lHeight = Int((m_WindowList(i).Percent / lTotalPercent) * lMaxHeight)
                    m_WindowList(i).Move lLeft, lTop, lWidth, lHeight
                    lTop = lTop + lHeight + SPLITTER_HEIGHT
                End If
            End If
        Next '  I I
        Extender.AutoRedraw = False

    Exit Sub

Err_DockResize:
        Extender.AutoRedraw = False
        Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
    On Error GoTo 0

End Sub

Friend Sub DockSwap(Df As TDockForm, _
       x As Long, _
       Y As Long)

  Const constSource As String = m_constClassName & ".DockSwap"
  Dim Index         As Integer

    On Error GoTo Err_DockSwap
    m_WindowList.RemoveByHandle Df.hwnd
    Index = GetFormIndex(x, Y)
    m_WindowList.Append Df, Index

Exit Sub

Err_DockSwap:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Function EnableCloseButton(ByVal hwnd As Long, _
                                  Enable As Boolean) As Integer

    EnableSystemMenuItem hwnd, SC_CLOSE, xSC_CLOSE, Enable, "EnableCloseButton"

End Function

Private Sub EnableSystemMenuItem(hwnd As Long, _
                                 Item As Long, _
                                 Dummy As Long, _
                                 Enable As Boolean, _
                                 FuncName As String)

  Dim hMenu     As Long
  Dim MII       As MENUITEMINFO
  Dim lngMenuID As Long

    '-- end code
    On Error Resume Next
        ' Retrieve a handle to the window's system menu
        hMenu = GetSystemMenu(hwnd, 0)
        ' Retrieve the menu item information for the Max menu item/button
        MII.cbSize = Len(MII)
        MII.dwTypeData = String$(80, 0)
        MII.cch = Len(MII.dwTypeData)
        MII.fMask = MIIM_STATE
        If Enable Then
            MII.wID = Dummy
          Else
            MII.wID = Item
        End If
        If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then
            Err.Raise vbObjectError, "modCloseBtn::" & FuncName, "modCloseBtn::" & FuncName & "() - Menu Item Not Found"
            Exit Sub
        End If
        ' Switch the ID of the menu item so that VB can not undo the action itself
        lngMenuID = MII.wID
        If Enable Then
            MII.wID = Item
          Else
            MII.wID = Dummy
        End If
        MII.fMask = MIIM_ID
        If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then
            Err.Raise vbObjectError, "modCloseBtn::" & FuncName, "modCloseBtn::" & FuncName & "() - Error encountered " & "changing ID"
            Exit Sub
        End If
        ' Set the enabled / disabled state of the menu item
        If Enable Then
            MII.fState = MII.fState And Not MFS_GRAYED
          Else
            MII.fState = MII.fState Or MFS_GRAYED
        End If
        MII.fMask = MIIM_STATE
        If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then
            Err.Raise vbObjectError, "modCloseBtn::" & FuncName, "modCloseBtn::" & FuncName & "() - Error encountered " & "changing state"
            Exit Sub
        End If
        ' Activate the non-client area of the window to update the titlebar, and
        ' draw the Max button in its new state.
        SendMessage hwnd, WM_NCACTIVATE, True, 0
    On Error GoTo 0

End Sub

Public Property Get Expanded() As Boolean

    Expanded = m_Expanded

End Property

Public Property Let Expanded(bln As Boolean)

    m_Expanded = bln
    Extender_Paint

End Property

Private Sub Extender_DblClick()

    If m_Expanded Then
        Panel_Collapse
      Else
        Panel_Expand
    End If

End Sub

Private Sub Extender_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)

  Const constSource As String = m_constClassName & ".Extender_MouseDown"
  Dim RC            As RECT
  Dim hWndA         As Long
  Dim XX As Long
  Dim YY As Long

    ' ******************************************************************************
    ' Routine       : (Sub) Extender_MouseDown
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:34:16
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    '
    ' Description   : The Extender (picturebox) mouse down event is used
    '                 to perform splittings
    ' ******************************************************************************
    On Error GoTo Err_Extender_MouseDown
    ' if it is not resizable then exit
    If Not m_Expanded Then
        Exit Sub
    End If
    If m_Resizable = False Then
        Exit Sub
    End If
    ' get the form handle that will receive mouse events
    hWndA = Parent.Parent.hwnd
    ' activate flag for sizing purposes
    m_bSplitting = True

    YY = Y
    XX = x

    ' If we have vertical splitting then
    ' get the SplitForm rect else get the main form rect
    If DragDir = DRAG_VERTICAL Then
        ''debug.print "mdown dragvert"
        If Align = tdAlignTop Or Align = tdAlignBottom Then
            OverFormDragArea x, Y, m_FormIndex
            GetFormRect m_FormIndex, RC
            '            GetClipRect Rc
          Else
            OverFormDragArea x, Y, m_FormIndex
            GetFormRect m_FormIndex, RC
        End If
      Else
        GetClipRect RC
    End If
    ' Let's see what to split based on the dragdir var
    If DragDir = DRAG_HORIZONTAL Then
        If m_bPanelSizing Then
            cHS.SplitterMouseDown hWndA, RC, x, Y
        End If
        ElseIf DragDir = DRAG_VERTICAL Then
        
        ''debug.print "extender mousedown vertical Split"
        If Align = tdAlignBottom Or Align = tdAlignTop Then
            If m_bDockedFormSizing Then
                cVS.SplitterMouseDown hWndA, RC, x, Y
            End If
            deltaY = cVS.Offset
          Else
           If m_bDockedFormSizing Then
            cVS.SplitterMouseDown hWndA, RC, x, Y
           End If
            deltaY = cVS.Offset
        End If
    End If

    ' - No Need To Repaint If We Have Not Moved
    If DragDir = DRAG_HORIZONTAL Then
        m_bSplitting = False
        If cHS.Offset = 0 Then Exit Sub
      Else
        m_bSplitting = False
        If cVS.Offset = 0 Then Exit Sub
    End If

    m_iFormIndex = -1
    
    ' lock updates
    LockWindowUpdate Extender.Parent.hwnd

    ' check drag direction
    If DragDir = DRAG_HORIZONTAL Then
        ''debug.print "extender mousedown"
        If Extender.Align = tdAlignLeft Or Extender.Align = tdAlignRight Then
            If Extender.Align = tdAlignLeft Then
                Extender.width = Extender.width + cHS.Offset
              Else
                Extender.width = Extender.width - cHS.Offset
            End If
          ElseIf Extender.Align = tdAlignTop Or Extender.Align = tdAlignBottom Then
            Extender.Visible = False
            If Extender.Align = tdAlignTop Then
                Extender.Top = 1000
                Extender.Height = Extender.Height + cHS.Offset
              ElseIf Extender.Align = tdAlignBottom Then
                Extender.Top = 0
                Extender.Height = Extender.Height - cHS.Offset
            End If
            Extender.Visible = True
            'm_bDirty = True
        End If
    End If

    ' repaint the host
    Extender_Paint
    ' activate flag for sizing purposes
    m_bSplitting = False
    ' reset updating
    LockWindowUpdate ByVal 0&
    ' raise sizing event
    Parent.TriggerEvent "ResizePanel", Object

Exit Sub

Err_Extender_MouseDown:
    ' reset updating
    'LockWindowUpdate ByVal 0&
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub Extender_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)

  Dim over As Integer
  Static bOver As Boolean
  
    ' ******************************************************************************
    ' Routine       : (Sub) Extender_MouseMove
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:33:18
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    '
    ' Description   : While moving over the Extender check for mouse cursor
    '                 and direction
    ' ******************************************************************************
    On Error GoTo Err_Extender_MouseMove
    ' if it is not resisable exit

    If m_Resizable = False Then
        Exit Sub
    End If
    If Not Me.Parent.MainAppActive Then
        Exit Sub
    End If
    If m_bSizing And Not Me.Parent.AutoExpand And m_AutoShowCaptionOnCollapse Then
        If Not Expanded Then
            Panel_Expand
        End If
        'debug.print "Expaning On Sizing"
    End If
    ' Set the correct arrow type.
     over = OverDragArea(x, Y)
     
   If Me.Expanded Then
      Extender.MousePointer = vbDefault
   End If
   
    Select Case over
      Case DRAG_HORIZONTAL
        bOver = False
        If Align = tdAlignLeft Or Align = tdAlignRight Then
            If m_bPanelSizing Then
           'debug.Print "kk"
            Extender.MousePointer = vbCustom
            Extender.MouseIcon = LoadResPicture(tdSizeWE, vbResCursor)
            cHS.Orientation = espVertical
            End If
            ''debug.print "ext mdown leftright draghoriz chsorivert"
          ElseIf Align = tdAlignTop Or Align = tdAlignBottom Then
          If m_bPanelSizing Then
         'debug.Print "lll"
                Extender.MousePointer = vbCustom
                Extender.MouseIcon = LoadResPicture(tdSizeNS, vbResCursor)
                cHS.Orientation = espHorizontal
            ''debug.print "ext mdown topbottom draghoriz chsorihoriz"
           End If
        End If
      Case DRAG_VERTICAL
        bOver = False
        ''debug.print "vert cur"
        If Align = tdAlignLeft Or Align = tdAlignRight Then
            If m_bDockedFormSizing Then
               'debug.Print "hehehee"
                Extender.MousePointer = vbCustom
                Extender.MouseIcon = LoadResPicture(tdSizeNS, vbResCursor)
                cVS.Orientation = espHorizontal
            End If
            ''debug.print "ext mdown leftright dragvert cvsorihoriz"
          ElseIf Align = tdAlignTop Or Align = tdAlignBottom Then
          If m_bDockedFormSizing Then
         'debug.Print "AAA"
            Extender.MousePointer = vbCustom
            Extender.MouseIcon = LoadResPicture(tdSizeWE, vbResCursor)
            cVS.Orientation = espVertical
        End If
            ''debug.print "ext mdown topbottom dragvert cvsorivert"
        End If
      Case 99
        
        bOver = True
        If Align = tdAlignLeft Then
            Extender.MousePointer = vbCustom
            Extender.MouseIcon = LoadResPicture(123, vbResCursor)
        End If
        If Align = tdAlignRight Then
            Extender.MousePointer = vbCustom
            Extender.MouseIcon = LoadResPicture(124, vbResCursor)
        End If
        If Align = tdAlignTop Then
            Extender.MousePointer = vbCustom
            Extender.MouseIcon = LoadResPicture(121, vbResCursor)
        End If
        If Align = tdAlignBottom Then
            Extender.MousePointer = vbCustom
            Extender.MouseIcon = LoadResPicture(122, vbResCursor)
        End If
        If Me.Parent.AutoExpand Then
            Panel_Expand
        End If
        If Not Me.Expanded Then
            RefreshCollaped
        End If
      
      Case Else
        'Extender.MousePointer = vbDefault
        bOver = False
    End Select
    DragDir = over
   
    On Error GoTo 0

Exit Sub

Err_Extender_MouseMove:
    '  Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
    Resume Next

End Sub

Private Sub Extender_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    m_bSplitting = False
     Extender.MousePointer = vbDefault

End Sub

Private Sub Extender_Paint()

  Const constSource As String = m_constClassName & ".Extender_Paint"

    ' ******************************************************************************
    ' Routine       : (Sub) Extender_Paint
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:33:52
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : Repaint the host
    ' ******************************************************************************
    On Error GoTo Err_Extender_Paint
    On Error Resume Next
        If m_Redraw Then
            If m_bDirty Then
                Extender.Cls
                DockResize
                m_bDirty = False
            End If
            DockArrange
        End If
    Exit Sub

Err_Extender_Paint:
        Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
    On Error GoTo 0

End Sub

Private Sub Extender_Resize()

    Dim collapsePanelWidth  As Long
    Dim collapsePanelHeight As Long
    Dim VSNetOffset As Integer

    On Error Resume Next
    ' ******************************************************************************
    ' Routine       : (Sub) Extender_Resize
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:34:02
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : Resize the host
    ' ******************************************************************************
    m_bSplitting = False
    
    
    If Me.Parent.CaptionStyle = tdCaptionVSNet Then
        VSNetOffset = 2
    Else
        VSNetOffset = 0
    End If
    
    
    collapsePanelWidth = ((getCaptionHeight + VSNetOffset) * Screen.TwipsPerPixelX) + SPLITTER_WIDTH
    collapsePanelHeight = (getCaptionHeight * Screen.TwipsPerPixelY) + SPLITTER_HEIGHT
    'LockWindowUpdate Extender.hwnd
    Extender.AutoRedraw = False
    If m_bSplitting = False And m_Expanded Then
        m_bDirty = True
    End If
    Extender_Paint
    If Extender.width <= collapsePanelWidth Or Extender.Height <= collapsePanelHeight Then
        Expanded = False
      Else
        Expanded = True
    End If
    If Not Expanded Then
        If m_Align = tdAlignLeft Or m_Align = tdAlignRight Then
            If m_AutoShowCaptionOnCollapse Then
                If Extender.width <= collapsePanelWidth Then
                    RefreshCollaped
                    Extender.AutoRedraw = True
                End If
            End If
          Else
            If m_AutoShowCaptionOnCollapse Then
                If Extender.Height <= collapsePanelHeight Then
                    RefreshCollaped
                    Extender.AutoRedraw = False
                End If
            End If
        End If
    End If
    Extender.AutoRedraw = False
   'debug.Print "panel ID: " & Me.Index
    

End Sub

Private Sub GetClipRect(RC As RECT)

  Dim rcPanel As RECT

    GetWindowRect Extender.hwnd, RC
    ' restrict resizing over other Panels
    ''debug.print m_Align
    If m_Align = tdAlignRight Then
        Parent.Panels(tdAlignLeft).GetHostRect rcPanel
        RC.Left = rcPanel.Right + MIN_HORIZONTAL
        RC.Right = RC.Right - MIN_HORIZONTAL
      ElseIf m_Align = tdAlignLeft Then
        Parent.Panels(tdAlignRight).GetHostRect rcPanel
        RC.Right = rcPanel.Left - MIN_HORIZONTAL
        RC.Left = RC.Left + MIN_HORIZONTAL
      ElseIf m_Align = tdAlignTop Then
        Parent.Panels(tdAlignBottom).GetHostRect rcPanel
        RC.Bottom = rcPanel.Top - MIN_VERTICAL
        RC.Top = RC.Top + MIN_VERTICAL
      ElseIf m_Align = tdAlignBottom Then
        Parent.Panels(tdAlignTop).GetHostRect rcPanel
        RC.Top = rcPanel.Bottom + MIN_VERTICAL
        RC.Bottom = RC.Bottom - MIN_VERTICAL
    End If

End Sub

Private Function GetFormIndex(x As Long, _
                              Y As Long) As Integer

  Const constSource As String = m_constClassName & ".GetFormIndex"
  Dim i             As Integer
  Dim RC            As RECT
  Dim Count         As Integer

    On Error GoTo Err_GetFormIndex
    GetFormIndex = 1
    Count = m_WindowList.VisibleCount
    If Count = 0 Then
        Exit Function
    End If
    For i = 1 To m_WindowList.Count
        If m_WindowList(i).Visible Then
            GetWindowRect m_WindowList(i).hwnd, RC
            If PtInRect(RC, x, Y) > 0 Then
                GetFormIndex = i + 1
                Exit Function
            End If
        End If
    Next '  I I
    GetWindowRect Extender.hwnd, RC
    If Y < RC.Top Then
        GetFormIndex = 1
      Else
        GetFormIndex = m_WindowList.Count + 1
    End If

Exit Function

Err_GetFormIndex:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Function

Private Sub GetFormRect(ByVal Index As Integer, _
                        RC As RECT)

  Const constSource As String = m_constClassName & ".GetFormRect"
  Dim lTop          As Long
  Dim lBottom       As Long
  Dim Rw            As RECT

    ' ******************************************************************************
    ' Routine       : (Sub) GetFormRect
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:15:07
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    '
    ' Description   : Get the form rect in screen coordinates based on th index
    ' ******************************************************************************
    On Error GoTo Err_GetFormRect
    If m_WindowList.Count = 0 Then
        Exit Sub
    End If
    GetWindowRect Extender.hwnd, Rw
    GetWindowRect m_WindowList(Index).hwnd, RC
    If Align = tdAlignLeft Or Align = tdAlignRight Then
        lTop = RC.Top + (GetSystemMetrics(SM_CYCAPTION) + (10 * GetSystemMetrics(SM_CYBORDER)))
        'lTOp = Rw.Top
    End If
    GetWindowRect m_WindowList(Index + 1).hwnd, RC
    If RC.Bottom > Rw.Bottom Then
        RC.Bottom = Rw.Bottom
    End If
    If Align = tdAlignLeft Or Align = tdAlignRight Then
        lBottom = RC.Bottom - (GetSystemMetrics(SM_CYCAPTION) + (10 * GetSystemMetrics(SM_CYBORDER)))
        RC.Top = lTop
        RC.Bottom = lBottom
      Else
        RC.Left = Rw.Left
    End If

Exit Sub

Err_GetFormRect:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Friend Sub GetHostRect(ByRef RC As RECT)

  ' ******************************************************************************
  ' Routine       : GetHostRect
  ' Created by    : Marclei V Silva
  ' Machine       : ZEUS
  ' Date-Time     : 28/08/005:54:02
  ' Inputs        :
  ' Outputs       :
  ' Credits       :
  ' Modifications :
  ' Description   : Get the host rect based on its align property
  '                 I use this function to figure out the position of the
  '                 picture box object, once when it is invisible it is not
  '                 possible to say where it really is
  ' ******************************************************************************
  ' if the Extender is visible no problem,
  ' just grab the visible rect

    If Extender.Visible Then
        GetWindowRect Extender.hwnd, RC
      Else
        ' here the problem is that we don't know exactly where the
        ' Extender is so get container rect
        GetWindowRect Parent.Parent.hwnd, RC
        ' get a rect based ona supposed position
        If Extender.Align = tdAlignRight Then
            RC.Left = RC.Right - (Parent.PanelWidth / Screen.TwipsPerPixelY)
          ElseIf Extender.Align = tdAlignBottom Then
            RC.Top = RC.Bottom - (Parent.PanelHeight / Screen.TwipsPerPixelY)
          ElseIf Extender.Align = tdAlignTop Then
            RC.Bottom = RC.Top + (Parent.PanelHeight / Screen.TwipsPerPixelY)
          ElseIf Extender.Align = tdAlignLeft Then
            RC.Right = RC.Left + (Parent.PanelWidth / Screen.TwipsPerPixelY)
        End If
    End If

End Sub

Public Property Get Height() As Variant

    Height = Extender.Height

End Property

Public Property Let Height(ByVal vData As Variant)

    Extender.Height = vData

End Property

Public Property Get hwnd() As Long

    hwnd = Extender.hwnd

End Property

Public Property Get Index() As Integer

    Index = m_Index

End Property

Friend Property Let Index(New_Index As Integer)

    m_Index = New_Index

End Property

Public Function IsDocked(Df As TDockForm) As Boolean

  Const constSource As String = m_constClassName & ".IsDocked"
  Dim idx           As Integer

    ' ******************************************************************************
    ' Routine       : (Function) IsDocked
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 09/06/2000 - 14:56:37
    ' Inputs        : df : TDockform
    ' Outputs       :
    ' Modifications :
    '
    ' Description   : check if the given docked form is docked in this host panel
    ' ******************************************************************************
    On Error GoTo Err_IsDocked
    IsDocked = False
    If m_WindowList.Count = 0 Then
        Exit Function
    End If
    For idx = 1 To m_WindowList.Count
        If m_WindowList(idx).hwnd = Df.hwnd Then
            IsDocked = True
            Exit Function
        End If
    Next '  IDX IDX

Exit Function

Err_IsDocked:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Function

Friend Function IsVisible() As Boolean

  Const constSource As String = m_constClassName & ".IsVisible"
  Dim idx           As Integer
  Dim Count         As Integer

    ' ******************************************************************************
    ' Routine       : (Function) IsVisible
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 09/06/2000 - 14:18:33
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : check if the given form is docked in here
    '                 this is a friend function the end-user has no access
    '                 to it
    ' ******************************************************************************
    On Error GoTo Err_IsVisible
    IsVisible = False
    If m_WindowList.Count = 0 Then
        Exit Function
    End If
    For idx = 1 To m_WindowList.Count
        If m_WindowList(idx).Visible Then
            Count = Count + 1
        End If
    Next '  IDX IDX
    If m_Expanded Then
        IsVisible = Count > 0
      Else
        IsVisible = True
    End If

Exit Function

Err_IsVisible:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Function

Public Property Get Left() As Variant

    Left = Extender.Left

End Property

Public Property Let Left(ByVal vData As Variant)

    Extender.Left = vData

End Property

Public Property Get Object() As TTabDockHost

    Set Object = Me

End Property

Private Function OverDragArea(ByVal x As Single, _
                              ByVal Y As Single)

  Const constSource As String = m_constClassName & ".OverDragArea"
  Dim over          As Integer

    ' ******************************************************************************
    ' Routine       : (Function) OverDragArea
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:24:22
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    '
    ' Description   : Return drag flags indicating the area at this point
    ' ******************************************************************************
    On Error GoTo Err_OverDragArea
    ' initially, dragging is none
    over = DRAG_NONE
    If Align = tdAlignLeft And x > Extender.width - (SPLITTER_WIDTH) And x < Extender.width Then
        over = DRAG_HORIZONTAL
        ''debug.print "overdrag left horiz"
      ElseIf Align = tdAlignRight And x > 0 And x < (5 * SPLITTER_WIDTH) Then
        over = DRAG_HORIZONTAL
        ''debug.print "overdrag right horiz"
      ElseIf Align = tdAlignTop And Y > Extender.Height - (SPLITTER_HEIGHT) And Y < Extender.Height Then
        over = DRAG_HORIZONTAL
        ''debug.print "overdrag top horiz"
      ElseIf OverFormDragArea(x, Y) Then
        ''debug.print "overdrag bottom vert"
        over = DRAG_VERTICAL
      ElseIf Align = tdAlignBottom Then
        If Y > 0 And Y < (SPLITTER_HEIGHT) Then
            over = DRAG_HORIZONTAL
            ''debug.print "overdrag bottom horiz"
          ElseIf OverFormDragArea(x, Y) Then
            ''debug.print "overdrag bottom vert"
            over = DRAG_VERTICAL
        End If
      Else
        If OverFormDragArea(x, Y) Then
            over = DRAG_VERTICAL
        End If
    End If
    If Not m_Expanded Then
        over = 99
    End If
    OverDragArea = over

Exit Function

Err_OverDragArea:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Function

Private Function OverFormDragArea(x As Single, _
                                  Y As Single, _
                                  Optional wIndex As Integer) As Boolean

  Const constSource As String = m_constClassName & ".OverFormDragArea"
  Dim i             As Integer
  Dim lTop          As Long
  Dim Count         As Integer
  Dim lLeft         As Long

    ' ******************************************************************************
    ' Routine       : (Function) OverFormDragArea
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 10/06/2000 - 7:06:34
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    '
    ' Description   : this will check if the cursor position is over
    '                 a draggable are between forms
    ' ******************************************************************************
    On Error GoTo Err_OverFormDragArea
    OverFormDragArea = False
    Count = m_WindowList.VisibleCount
    If Count = 0 Then
        Exit Function
    End If
    lTop = 0
    wIndex = 0
    If Align = tdAlignTop Then
        Count = Count - 1
      ElseIf Align = tdAlignBottom Then
        lTop = lTop + SPLITTER_HEIGHT
    End If
    For i = 1 To m_WindowList.Count
        If m_WindowList(i).Visible Then
            If Align = tdAlignLeft Or Align = tdAlignRight Then
                lTop = lTop + m_WindowList(i).Height
                If Y >= lTop And Y <= lTop + SPLITTER_HEIGHT Then
                    OverFormDragArea = True
                    wIndex = i
                    Exit Function
                End If
                lTop = lTop + SPLITTER_HEIGHT
              Else
                lLeft = lLeft + m_WindowList(i).width
                If x >= lLeft And x <= lLeft + SPLITTER_WIDTH Then
                    OverFormDragArea = True
                    wIndex = i
                    Exit Function
                End If
                lLeft = lLeft + SPLITTER_WIDTH
            End If
        End If
    Next '  I I

Exit Function

Err_OverFormDragArea:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Function

Public Sub Panel_Collapse()

    Dim f As TDockForm
    Dim VSNetOffset As Integer

    On Error GoTo Panel_Collapse_Err
    If Not m_Expanded Then
        Exit Sub
    End If
    If Not Me.Parent.MainAppActive Then
        Exit Sub
    End If
    Putfocus Extender.Parent.hwnd
    m_Expanded = False
    ' select host original size after we collapse it
    Select Case m_Align
      Case tdAlignLeft, tdAlignRight
        m_Size = Extender.width
      Case tdAlignTop, tdAlignBottom
        m_Size = Extender.Height
    End Select
    ' disable window refresh
    Extender.AutoRedraw = True
    LockWindowUpdate Extender.Parent.hwnd
    ' hide every window in the list
    For Each f In m_WindowList
        f.Tag = f.Visible       ' save visibility for future use
        f.Visible = False
    Next '
    ' clear host
    Extender.Cls
    Extender.AutoRedraw = True
    If Parent.CaptionStyle = tdCaptionVSNet Then
        Extender.BackColor = BlendColor(vb3DHighlight, vb3DLight)
    End If
    ' set collapsed size
    
    If Me.Parent.CaptionStyle = tdCaptionVSNet Then
        VSNetOffset = 2
    Else
        VSNetOffset = 0
    End If
    
    If m_Align = tdAlignLeft Or m_Align = tdAlignRight Then
        If m_AutoShowCaptionOnCollapse Then
            Extender.width = SPLITTER_WIDTH + ((getCaptionHeight + VSNetOffset) * Screen.TwipsPerPixelX) ' * 4
          Else
            Extender.width = SPLITTER_WIDTH
        End If
      Else
        If m_AutoShowCaptionOnCollapse Then
            Extender.Height = SPLITTER_HEIGHT + (getCaptionHeight * Screen.TwipsPerPixelY) '* 4
          Else
            Extender.Height = SPLITTER_HEIGHT
        End If
        If Extender.Align = tdAlignTop Then
            Top = 1000
        End If
        If Extender.Align = tdAlignBottom Then
            Top = 0
        End If
    End If
    '-- Draw The Collapsed Panel Captions
    If m_AutoShowCaptionOnCollapse Then
        RefreshCollaped
    End If
    Extender.AutoRedraw = True
    ' set host as collapsed
    m_Expanded = False
    ' restore window refresh
    LockWindowUpdate 0
    '<EhFooter>

Exit Sub

Panel_Collapse_Err:
    ' restore window refresh
    LockWindowUpdate 0
    Err.Raise vbObjectError + 100, "TabDock.TTabDockHost.Panel_Collapse", "TTabDockHost component failure::" & Err.Description
    '</EhFooter>

End Sub

Public Sub Panel_Expand()

  Dim f As TDockForm

    On Error GoTo Panel_Expand_Err
    If Not Me.Parent.MainAppActive Then
        Exit Sub
    End If
    Extender.AutoRedraw = True
    Extender.Cls
    Extender.AutoRedraw = False
    'For Each f In m_WindowList
    '    f.Panel.Repaint
    'Next
    m_Expanded = True
    ' stop window refresh
    LockWindowUpdate Extender.Parent.hwnd
    ' show all visible forms
    For Each f In m_WindowList
        ''debug.print f.Extender.Caption, f.Visible
        f.Visible = f.Tag
        ''debug.print f.Extender.Caption, f.Visible
    Next '  F F
    ' restore host size based on the alignment
    Extender.BackColor = vbButtonFace
    If m_Align = tdAlignLeft Or m_Align = tdAlignRight Then
        Extender.width = m_Size
      Else
        m_Redraw = False
        Extender.Height = m_Size
        m_Redraw = True
        '  trick, just to re align the host
        If Extender.Align = tdAlignTop Then
            Extender.Top = 1000
        End If
        If Extender.Align = tdAlignBottom Then
            Extender.Top = 0
        End If
    End If
    ' set flag as expanded
    m_Expanded = True
    ' restore window update
    LockWindowUpdate ByVal 0&
    '<EhFooter>
    Me.Parent.resetTimer Extender.Align

Exit Sub

Panel_Expand_Err:
    LockWindowUpdate 0&
    Err.Raise vbObjectError + 100, "TabDock.TTabDockHost.Panel_Expand", "TTabDockHost component failure::" & Err.Description
    '</EhFooter>

End Sub

Public Property Get Parent() As TTabDock

    Set Parent = ObjectFromPtr(m_ParentPtr)

End Property

Public Property Set Parent(NewValue As TTabDock)

    m_ParentPtr = PtrFromObject(NewValue)

End Property

Public Sub Refresh()

    Extender.Refresh

End Sub

Public Function RefreshCollaped()

  Dim f        As TDockForm
  Dim LastTop  As Long
  Dim LastLeft As Long
  Dim sCaption As String
  Dim frameRec As RECT
  Dim hbr      As Long
  Dim VsNet    As Boolean

    'Extender.AutoRedraw = True
    On Error GoTo Panel_Collapse_Err
    ' disable window refresh
    Extender.AutoRedraw = True
    Extender.Cls
    LockWindowUpdate Extender.Parent.hwnd
    ' hide every window in the list
    'For Each f In m_WindowList
    '    f.Tag = f.Visible       ' save visibility for future use
    '    f.Visible = False
    '    'f.Tag = True
    'Next
    ' clear host
            m_bCloseDown = False
        m_bCloseOver = False
        m_bChevronDown = False
        m_bChevronOver = False
    
    
    If Me.Parent.CaptionStyle = tdCaptionVSNet Then
        VsNet = True
        Extender.BackColor = BlendColor(vb3DHighlight, vb3DLight)
    Else
        VsNet = False
        Extender.BackColor = vbButtonFace
    End If
     
    Extender.Cls
    
    'Extender.AutoRedraw = True
    ''debug.print m_WindowList.Count
    '-- Draw The Collapsed Panel Captions
    If m_AutoShowCaptionOnCollapse Then
        If m_Align = tdAlignLeft Or m_Align = tdAlignRight Then
            LastTop = 1
            For Each f In m_WindowList
                If f.Tag Then
                    sCaption = f.Extender.Caption
                    With frameRec
                        .Top = LastTop + 5
                        If VsNet Then
                            .Bottom = LastTop + 100
                        Else
                            .Bottom = Len(sCaption) * Screen.TwipsPerPixelY
                        End If
                        .Left = 2
                        .Right = Extender.width / Screen.TwipsPerPixelY
                    End With
                    LastTop = frameRec.Bottom + 3
                    
                    If Not VsNet Then
                        drawGripper frameRec, Extender.hdc, EDGE_ETCHED, BF_RECT, False, 16, f.Panel.Align
                        If (f.HasMaxButton And f.HasCloseButton) Then
                            frameRec.Bottom = frameRec.Bottom + (getCaptionHeight * 15) / 2 ' (Len(sCaption) / 1.7) * Screen.TwipsPerPixelY
                            frameRec.Bottom = frameRec.Bottom + 16
                            'LastTop = LastTop + 12
                            LastTop = frameRec.Bottom + 3
                          ElseIf f.HasMaxButton Then
                            'LastTop = LastTop - 6
                            frameRec.Bottom = frameRec.Bottom + (Len(sCaption) / 3) * Screen.TwipsPerPixelY
                            frameRec.Bottom = frameRec.Bottom + (Len(sCaption) * 1.5)
                            LastTop = frameRec.Bottom + 3
                          ElseIf f.HasCloseButton Then
                            frameRec.Bottom = frameRec.Bottom + (Len(sCaption) / 3) * Screen.TwipsPerPixelY
                            frameRec.Bottom = frameRec.Bottom + (Len(sCaption) * 1.5)
                            LastTop = frameRec.Bottom + 3 ' LastTop + (Len(sCaption) * 3)
                          Else
                            frameRec.Bottom = frameRec.Bottom + (Len(sCaption) / 1.3) * Screen.TwipsPerPixelY
                            frameRec.Bottom = frameRec.Bottom + 16
                            LastTop = frameRec.Bottom + 3
                        End If
                    Else
                        With frameRec
                            .Right = .Right - 2
                        End With
                        LastTop = frameRec.Top + getCaptionHeight - 2
                    End If
                                        
                    drawGradientx frameRec, Extender.hdc, sCaption, False, True, 1, f.Extender, VsNet

                    If Not VsNet Then
                        frameRec.Right = frameRec.Right - 2
                        frameRec.Left = frameRec.Left + 2
                        
                        If f.HasCloseButton Then
                            frameRec.Top = frameRec.Top + 2
                            frameRec.Bottom = frameRec.Bottom + (getCaptionHeight) - 6 '10
                            If Parent.CaptionStyle = tdCaptionVSNet Then
                                ' hbr = GetSysColorBrush(vbButtonFace And &H1F&)
                                drawCloseButton Extender.hdc, hbr, frameRec, True, f
                              Else
                                DrawFrameControl Extender.hdc, frameRec, DFC_CAPTION, DFCS_CAPTIONCLOSE + DFCS_INACTIVE
                            End If
                        End If
                        
                        If f.HasMaxButton Then
                            If f.HasCloseButton Then
                                frameRec.Top = frameRec.Bottom + 2
                                frameRec.Bottom = frameRec.Bottom + (getCaptionHeight) - 2 '14
                            Else
                                frameRec.Top = frameRec.Top + 2
                                frameRec.Bottom = frameRec.Bottom + (getCaptionHeight) - 4 '10
                            End If
                            
                            frameRec.Right = frameRec.Right
                            frameRec.Left = frameRec.Left
                            
                            If Parent.CaptionStyle = tdCaptionVSNet Then
                                ' hbr = GetSysColorBrush(vbButtonFace And &H1F&)
                                m_bChevronOver = False
                                m_bChevronDown = False
                                drawChevronButton Extender.hdc, hbr, frameRec, True, f
                                
                              Else
                                DrawFrameControl Extender.hdc, frameRec, DFC_CAPTION, DFCS_CAPTIONRESTORE + DFCS_INACTIVE
                            End If
                        End If
                    End If
                End If

            Next '  F F
          Else
            LastLeft = 5
            For Each f In m_WindowList
                If f.Tag Then
                    sCaption = f.Extender.Caption
                    With frameRec
                        .Top = 0
                        .Bottom = Extender.ScaleHeight / Screen.TwipsPerPixelY
                        .Bottom = .Bottom - 1
                        .Left = LastLeft
                        If VsNet Then
                            .Right = LastLeft + 100
                        Else
                            .Right = Len(f.Extender.Caption) * Screen.TwipsPerPixelX
                        End If
                    End With
                    
                    If VsNet Then
                        LastLeft = frameRec.Right + 1
                    Else
                        LastLeft = LastLeft + frameRec.Right + 1
                    End If

                    If Not VsNet Then
                        drawGripper frameRec, Extender.hdc, EDGE_ETCHED, BF_RECT, False, 16, f.Panel.Align
                    Else
                        With frameRec
                            .Top = .Top + 2
                        End With
                    End If
                    
                    frameRec.Right = LastLeft
                    
                    If VsNet Then
                        drawGradientx frameRec, Extender.hdc, sCaption, False, True, 0, f.Extender, VsNet
                        LastLeft = frameRec.Left + getCaptionHeight + 2
                    Else
                        If f.HasMaxButton And f.HasCloseButton Then
                            frameRec.Right = frameRec.Right + 26
                            drawGradientx frameRec, Extender.hdc, sCaption, False, True, 0, f.Extender, VsNet
                            LastLeft = LastLeft + 18
                            frameRec.Right = frameRec.Right - getCaptionHeight '10
                            'frameRec.Right = frameRec.Right - 6
                            frameRec.Top = frameRec.Top + 2
                            frameRec.Bottom = frameRec.Bottom - 2
                            frameRec.Left = frameRec.Right - getCaptionHeight + 4 '12
                            If Parent.CaptionStyle = tdCaptionVSNet Then
                                ' hbr = GetSysColorBrush(vbButtonFace And &H1F&)
                                drawCloseButton Extender.hdc, hbr, frameRec, True, f
                              Else
                                DrawFrameControl Extender.hdc, frameRec, DFC_CAPTION, DFCS_CAPTIONCLOSE + DFCS_INACTIVE
                            End If
                            frameRec.Right = frameRec.Right - getCaptionHeight + 2 '10
                            'frameRec.Right = frameRec.Right - 4
                            frameRec.Left = frameRec.Right - getCaptionHeight + 4 '15
                            If Parent.CaptionStyle = tdCaptionVSNet Then
                                ' hbr = GetSysColorBrush(vbButtonFace And &H1F&)
                                drawChevronButton Extender.hdc, hbr, frameRec, True, f
                                'debug.print "DrawClose"
                              Else
                                DrawFrameControl Extender.hdc, frameRec, DFC_CAPTION, DFCS_CAPTIONRESTORE + DFCS_INACTIVE
                            End If
                          ElseIf f.HasMaxButton Then
                            LastLeft = LastLeft + 10
                            frameRec.Right = frameRec.Right + 20
                            drawGradientx frameRec, Extender.hdc, sCaption, False, True, 0, f.Extender, VsNet
                            'frameRec.Right = frameRec.Right + 10
                            frameRec.Right = frameRec.Right - getCaptionHeight '6
                            frameRec.Top = frameRec.Top + 2
                            frameRec.Bottom = frameRec.Bottom - 2
                            frameRec.Left = frameRec.Right - getCaptionHeight + 4
                            If Parent.CaptionStyle = tdCaptionVSNet Then
                                ' hbr = GetSysColorBrush(vbButtonFace And &H1F&)
                                drawChevronButton Extender.hdc, hbr, frameRec, True, f
                              Else
                                DrawFrameControl Extender.hdc, frameRec, DFC_CAPTION, DFCS_CAPTIONRESTORE + DFCS_INACTIVE
                            End If
                          ElseIf f.HasCloseButton Then
                            LastLeft = LastLeft + 10
                            frameRec.Right = frameRec.Right + 20
                            drawGradientx frameRec, Extender.hdc, sCaption, False, True, 0, f.Extender, VsNet
                            frameRec.Right = frameRec.Right - getCaptionHeight '6
                            frameRec.Top = frameRec.Top + 2
                            frameRec.Bottom = frameRec.Bottom - 2
                            frameRec.Left = frameRec.Right - getCaptionHeight + 4
                            If Parent.CaptionStyle = tdCaptionVSNet Then
                                ' hbr = GetSysColorBrush(vbButtonFace And &H1F&)
                                drawCloseButton Extender.hdc, hbr, frameRec, True, f
                              Else
                                DrawFrameControl Extender.hdc, frameRec, DFC_CAPTION, DFCS_CAPTIONCLOSE + DFCS_INACTIVE
                            End If
                          Else
                            drawGradientx frameRec, Extender.hdc, sCaption, False, True, 0, f.Extender, VsNet
                            frameRec.Right = frameRec.Right + 6
                        End If
                    End If
                End If
            Next '  F F
        End If
    End If
    Extender.AutoRedraw = False
    DeleteObject hbr
    ' restore window refresh
    LockWindowUpdate 0
    '<EhFooter>
    On Error GoTo 0

Exit Function

Panel_Collapse_Err:
    ' restore window refresh
    LockWindowUpdate 0
    On Error Resume Next
        'Err.Raise vbObjectError + 100, "TabDock.TTabDockHost.Panel_Collapse", "TTabDockHost component failure::" & Err.Description
        '</EhFooter>

End Function

Friend Sub Repaint()

  Const constSource As String = m_constClassName & ".Repaint"

    ' ******************************************************************************
    ' Routine       : (Sub) Repaint
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 09/06/2000 - 14:17:09
    ' Inputs        : N/A
    ' Outputs       : N/A
    ' Modifications :
    ' Description   : Repaint current host panel
    ' ******************************************************************************
    On Error GoTo Err_Repaint
    m_bDirty = True
    Extender_Paint

Exit Sub

Err_Repaint:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Property Let Resizable(New_Resizable As Boolean)

    m_Resizable = New_Resizable

End Property

Public Property Get Resizable() As Boolean

    Resizable = m_Resizable

End Property

Public Property Get Splitting() As Boolean

    Splitting = m_bSplitting

End Property

Public Property Get Top() As Variant

    Top = Extender.Top

End Property

Public Property Let Top(ByVal vData As Variant)

    On Error Resume Next
        Extender.Top = vData
    On Error GoTo 0

End Property

Public Sub UnDock(ByVal Df As TDockForm)

  Const constSource As String = m_constClassName & ".UnDock"
  Dim hWndA         As Long
  Dim Style         As Long
  Dim hWndDesktop   As Long
  
  
    ' ******************************************************************************
    ' Routine       : (Sub) UnDock
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 09/06/2000 - 14:59:00
    ' Inputs        : Byval df: TDockForm
    ' Outputs       :
    ' Modifications :
    ' Credits       : All credits for part of this code to vbWeb:
    '
    ' Description   : Undock and show the form
    ' ******************************************************************************
    On Error GoTo Err_UnDock
    ' LockWindowUpdate Extender.Parent.hWnd
    ' change the dock form state
    Df.State = tdUndocked
    Df.Redraw = False
    m_iFormIndex = -1
    
    ' get form handle
    hWndA = Df.hwnd
    ' if this form can not float then dock it

    If (Df.Style And tdDockFloat) = False Then
        ' if it can not float and it is undocked...
        If Df.State = tdUndocked Then
            ' hide the form
            Call ShowWindow(hWndA, SW_HIDE)
        End If
        Exit Sub
    End If
    ' remove the form from the control list
    m_WindowList.RemoveByHandle Df.hwnd
    ' This is necessary in order to make top and bottom
    ' panels to align after/before any existing
    ' aligned controls on your form (Statusbars, Toolbars, etc...)
    If Extender.Align = tdAlignTop Then
        Top = 1000
    End If
    If Extender.Align = tdAlignBottom Then
        Top = 0
    End If
    ' first of all let's hide this window
    ' change form's style
    Style = GetWindowLong(hWndA, GWL_STYLE)
    Style = Style And Not WS_DLGFRAME ' dialog frame
    Style = Style Or WS_CAPTION       ' has a caption
    Style = Style Or WS_THICKFRAME    ' has sizing border
    Style = Style Or WS_SYSMENU       ' add system menu
    Style = Style Or WS_EX_TOOLWINDOW
    ' set new style
    SetWindowLong hWndA, GWL_STYLE, Style
    ' This must be a toolwindow
    Style = GetWindowLong(hWndA, GWL_EXSTYLE)
    Style = Style ' Or WS_EX_TOOLWINDOW
    Style = Style And Not (WS_EX_APPWINDOW)
    If Df.HasCloseButton Then
        SetWindowLong hWndA, GWL_EXSTYLE, Style
        '        EnableCloseButton hWndA, True
      Else
        EnableCloseButton hWndA, False
    End If
    '    ' good to use it after SetWindowLong() function
    SetWindowPos hWndA, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
    ' set the parent to 0->Desktop (tricky!)

    'Call SetParent(pbUndock.hwnd, hWndDesktop)
    '    pbUndock.Visible = True

    'MoveWindow Df.hwnd, Df.FloatingLeft, Df.FloatingTop, Df.FloatingWidth, Df.FloatingHeight, 1
    Call SetParent(hWndA, hWndDesktop)
    ' move window to its floating position
    'debug.print "F Top= : " & Df.FloatingTop
    If Df.FloatingTop = -1 Then
        Df.FloatingTop = ((Screen.Height / Screen.TwipsPerPixelY) / 2) - ((Df.Height / Screen.TwipsPerPixelY)) '/ 2
        Df.FloatingLeft = ((Screen.width / Screen.TwipsPerPixelX) / 2) - ((Df.width / Screen.TwipsPerPixelX)) / 3
        MoveWindow Df.hwnd, Df.FloatingLeft, Df.FloatingTop, Df.FloatingWidth, Df.FloatingHeight, 1
        'debug.print "F Top"
      Else
        MoveWindow Df.hwnd, Df.FloatingLeft, Df.FloatingTop, Df.FloatingWidth, Df.FloatingHeight, 1
    End If
    ' show the window
    ShowWindow hWndA, SW_SHOW
    ' make the form 'float' above the MDI form
    SetWindowLong hWndA, GWL_HWNDPARENT, Parent.Parent.hwnd
    ' update window pos
    SetWindowPos hWndA, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
    ' extender is visible if there is more forms docked...
    If Not IsVisible Then
        Extender.Visible = False
    End If
    ' chage visiblity
    Df.Visible = True
    ' Call this function to raise TabDock Event FormUnDocked
    Parent.TriggerEvent "UnDock", Df
    ' needs to re-calc heights
    m_bDirty = True
    ' repaint the control
    Extender_Paint
    Df.Redraw = True
    ' free windows
    '    LockWindowUpdate ByVal 0&

    If Not m_Expanded Then
        RefreshCollaped
    End If

Exit Sub

Err_UnDock:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Property Get Visible() As Boolean

    Visible = Extender.Visible

End Property

Public Property Let Visible(ByVal vData As Boolean)

    On Error Resume Next
        '    LockWindowUpdate Parent.Parent.hWnd
        Extender.Visible = vData
        If vData Then
            If IsVisible = False Then
                Extender.Visible = False
            End If
        End If
        '    LockWindowUpdate ByVal 0&
    On Error GoTo 0

End Property

Public Property Get width() As Variant

    width = Extender.width

End Property

Public Property Let width(ByVal vData As Variant)

    Extender.width = vData

End Property

Public Property Get WindowList() As TDockForms

    Set WindowList = m_WindowList

End Property

