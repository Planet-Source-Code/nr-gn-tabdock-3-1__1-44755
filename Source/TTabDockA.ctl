VERSION 5.00
Begin VB.UserControl TTabDock 
   CanGetFocus     =   0   'False
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2940
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "TTabDockA.ctx":0000
   ScaleHeight     =   195.152
   ScaleMode       =   0  'User
   ScaleWidth      =   184.471
   ToolboxBitmap   =   "TTabDockA.ctx":08CA
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   3000
      Left            =   1680
      Top             =   660
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   3000
      Left            =   1200
      Top             =   660
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   3000
      Left            =   720
      Top             =   660
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   3000
      Left            =   240
      Top             =   660
   End
End
Attribute VB_Name = "TTabDock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' ******************************************************************************
' Control    : TabDock.ctl
' Created by : Marclei V Silva
' Machine    : ZEUS
' Date-Time  : 09/05/2000 3:13:22
' Description: Docking system engine
' ******************************************************************************
Option Explicit
Option Base 1
' Keep up with the errors
Private Const g_ErrConstant                       As Long = vbObjectError + 1000
Private Const m_constClassName                    As String = "TTabDock"
Private Const m_Grad1                             As Long = vbRed
Private Const m_Grad2                             As Long = vbBlack
Private m_lngErrNum                               As Long
Private m_strErrStr                               As String
Private m_strErrSource                            As String
Private m_Panels                                  As TTabDockHosts
Private m_DockedForms                             As TDockForms
Private Const m_PersistantDef                     As Boolean = False
Private NewHWND                                   As Long
Private m_Persistant                              As Boolean
Private m_AutoShowCaptionOnCollapse               As Boolean
Private m_MainApp                                 As Boolean
' Events Held by this control
Public Event FormDocked(ByVal DockedForm As TDockForm)
Attribute FormDocked.VB_Description = "Occurs when the user drag and dock a form at a specific panel on the screen"
Public Event FormUnDocked(ByVal DockedForm As TDockForm)
Attribute FormUnDocked.VB_Description = "Occurs when the user undocks a form from a specific panel"
Public Event FormShow(ByVal DockedForm As TDockForm)
Public Event FormHide(ByVal DockedForm As TDockForm)
Public Event MenuClick(ByVal ItemIndex As Long)
Public Event PanelResize(ByVal Panel As TTabDockHost)
Attribute PanelResize.VB_Description = "Occurs when a specific panel is resized. This is useful when you want to set a specific Height or width for a panel in the screen or avoid user to resize a panel to a not desired size."
Public Event PanelClick(ByVal Panel As TTabDockHost)
Public Event CaptionClick(ByVal DockedForm As TDockForm, ByVal Button As Integer, ByVal x As Single, ByVal Y As Single)
Attribute CaptionClick.VB_Description = "Occurs when the user clicks on the caption bar of a form. This is very useful when we want to show a popup menu for that form like Dockable or Hide."
' Default Property Values:
Private Const m_def_BackColor                     As Long = &H8000000F
Private Const m_def_BorderStyle                   As Integer = 0  ' flat
Private Const m_def_CaptionStyle                  As Integer = 0  ' etched
Private Const m_def_PanelHeight                   As Integer = 1300
Private Const m_def_PanelWidth                    As Integer = 2500
Private Const m_def_Visible                       As Integer = 0
' Property Variables:
Private m_BackColor                               As OLE_COLOR
Private m_BorderStyle                             As tdBorderStyles
Private m_CaptionStyle                            As tdCaptionStyles
Private m_MaximizeButton                          As Boolean
Private m_AutoExpand                              As Boolean
Private m_AutoCollapseTop                         As Boolean
Private m_AutoCollapseLeft                        As Boolean
Private m_AutoCollapseRight                       As Boolean
Private m_AutoCollapseBottom                      As Boolean
Private m_CollapseInterval                        As Long
Private m_Parent                                  As Object
Private m_PanelHeight                             As Long
Private m_PanelWidth                              As Long
Private m_Visible                                 As Boolean
Private m_bLoaded                                 As Boolean
Private m_Gradient1                               As OLE_COLOR
Private m_Gradient2                               As OLE_COLOR
Private WithEvents m_Size                         As cSizer
Attribute m_Size.VB_VarHelpID = -1
Private m_bSmartSizing                            As Boolean
Private m_PanelResizeTop                          As Boolean
Private m_PanelResizeLeft                         As Boolean
Private m_PanelResizeRight                        As Boolean
Private m_PanelResizeBottom                       As Boolean
Private m_PanelTopDockFormResize                  As Boolean
Private m_PanelRightDockFormResize                As Boolean
Private m_PanelLeftDockFormResize                 As Boolean
Private m_PanelBottomDockFormResize               As Boolean

Implements ISubclass
' ******************************************************************************
' Routine       : AddForm
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 28/08/006:00:45
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Description   : Adds forms to the main engine
' ******************************************************************************
Private DummyToKeepDecCommentsInDEclarations As Boolean

Public Function AddForm(ByVal Item As Object, _
                        Optional State As tdDockedState = tdUndocked, _
                        Optional Align As tdAlignProperty = tdAlignLeft, _
                        Optional Key As String, _
                        Optional Style As tdDockStyles, _
                        Optional Percent As Integer, _
                        Optional bHasMaxButton As Boolean = False, _
                        Optional bHasCloseButton As Boolean = True) As TDockForm
Attribute AddForm.VB_Description = "Add a form reference to the dock system and updates its initial properties"

  Const constSource As String = m_constClassName & ".AddForm"

    On Error Resume Next
    On Error GoTo Err_AddForm
    If IsFormLoaded(Item.hwnd) Then
        m_strErrStr = "Form is already loaded"
        m_strErrSource = constSource
        m_lngErrNum = 0
        m_lngErrNum = m_lngErrNum + g_ErrConstant
        Err.Raise Description:="Unexpected Error: " & m_strErrStr, Number:=m_lngErrNum, Source:=constSource
    End If
    ' if we are initializing (panels were not created) then create panels
    If m_bLoaded = False Then
        LoadPanels
    End If
    ' loads the form if it wasn't loaded yet!
    Load Item
    ' if the form style was not furnished then set
    ' all styles available to the form
    If IsMissing(Style) Or IsEmpty(Style) Or Style = 0 Or Style = tdShowInvisible Then
        Style = Style Or tdDockFloat
        Style = Style Or tdDockLeft
        Style = Style Or tdDockRight
        Style = Style Or tdDockTop
        Style = Style Or tdDockBottom
    End If
    If Persistant Then
        Align = GetSetting(App.Title, "Docking", Key & "Align", Align)
    End If
    ' add the form to the list
    Set AddForm = m_DockedForms.Add(Item, Panels(Align), Style, State, Key, NewHWND, Percent, bHasMaxButton, bHasCloseButton)

Exit Function

Err_AddForm:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
    On Error GoTo 0

End Function

Public Property Let PanelTopDockFormResize(bln As Boolean)

    m_PanelTopDockFormResize = bln
    PropertyChanged "PanelTopDockFormResize"
    ChangePanelDockFormRezise vbAlignTop, bln

End Property

Public Property Get PanelTopDockFormResize() As Boolean

    PanelTopDockFormResize = m_PanelTopDockFormResize

End Property
Public Property Let PanelLeftDockFormResize(bln As Boolean)

    m_PanelLeftDockFormResize = bln
    PropertyChanged "PanelLeftDockFormResize"
    ChangePanelDockFormRezise vbAlignLeft, bln

End Property

Public Property Get PanelLeftDockFormResize() As Boolean

    PanelLeftDockFormResize = m_PanelLeftDockFormResize

End Property
Public Property Let PanelRightDockFormResize(bln As Boolean)

    m_PanelRightDockFormResize = bln
    PropertyChanged "PanelRightDockFormResize"
    ChangePanelDockFormRezise vbAlignRight, bln

End Property

Public Property Get PanelRightDockFormResize() As Boolean

    PanelRightDockFormResize = m_PanelRightDockFormResize

End Property

Public Property Let PanelBottomDockFormResize(bln As Boolean)

    m_PanelBottomDockFormResize = bln
    PropertyChanged "PanelBottomDockFormResize"
    ChangePanelDockFormRezise vbAlignBottom, bln

End Property

Public Property Get PanelBottomDockFormResize() As Boolean

    PanelBottomDockFormResize = m_PanelBottomDockFormResize

End Property


Public Property Get PanelResizeRight() As Boolean

    PanelResizeRight = m_PanelResizeRight

End Property

Public Property Let PanelResizeRight(bln As Boolean)

    m_PanelResizeRight = bln
    PropertyChanged "PanelResizeRight"
    ChangePanelRezise vbAlignRight, bln

End Property

Public Property Get PanelResizeLeft() As Boolean

    PanelResizeLeft = m_PanelResizeLeft

End Property

Public Property Let PanelResizeLeft(bln As Boolean)

    m_PanelResizeLeft = bln
    PropertyChanged "PanelResizeLeft"
    ChangePanelRezise vbAlignLeft, bln

End Property


Public Property Get PanelResizeTop() As Boolean

    PanelResizeTop = m_PanelResizeTop

End Property

Public Property Let PanelResizeTop(bln As Boolean)

    m_PanelResizeTop = bln
    PropertyChanged "PanelResizeTop"
    ChangePanelRezise vbAlignTop, bln

End Property

Public Property Get PanelResizeBottom() As Boolean

    PanelResizeBottom = m_PanelResizeBottom

End Property

Public Property Let PanelResizeBottom(bln As Boolean)
    
    m_PanelResizeBottom = bln
    PropertyChanged "PanelResizeBottom"
    ChangePanelRezise vbAlignBottom, bln
    
End Property


Public Property Get AutoCollapseBottom() As Boolean

    AutoCollapseBottom = m_AutoCollapseBottom

End Property

Public Property Let AutoCollapseBottom(Auto As Boolean)

    m_AutoCollapseBottom = Auto
    PropertyChanged "AutoCollapseBottom"

End Property

Public Property Get AutoCollapseLeft() As Boolean

    AutoCollapseLeft = m_AutoCollapseLeft

End Property

Public Property Let AutoCollapseLeft(Auto As Boolean)

    m_AutoCollapseLeft = Auto
    PropertyChanged "AutoCollapseLeft"

End Property

Public Property Get AutoCollapseRight() As Boolean

    AutoCollapseRight = m_AutoCollapseRight

End Property

Public Property Let AutoCollapseRight(Auto As Boolean)

    m_AutoCollapseRight = Auto
    PropertyChanged "AutoCollapseRight"

End Property

Public Property Get AutoCollapseTop() As Boolean

    AutoCollapseTop = m_AutoCollapseTop

End Property

Public Property Let AutoCollapseTop(Auto As Boolean)

    m_AutoCollapseTop = Auto
    PropertyChanged "AutoCollapseTop"

End Property

Public Property Get AutoExpand() As Boolean

    AutoExpand = m_AutoExpand

End Property

Public Property Let AutoExpand(Auto As Boolean)

    m_AutoExpand = Auto
    PropertyChanged "AutoExpand"

End Property

Public Property Get AutoShowCaptionOnCollapse() As Boolean

    AutoShowCaptionOnCollapse = m_AutoShowCaptionOnCollapse

End Property

Public Property Let AutoShowCaptionOnCollapse(Auto As Boolean)

  Dim i As Integer

    m_AutoShowCaptionOnCollapse = Auto
    PropertyChanged "AutoShowCaptionOnCollapse"
    On Error Resume Next
        For i = 0 To m_Panels.Count - 1
            m_Panels(i).AutoShowCaptionOnCollapse = Auto
        Next '  I I
    On Error GoTo 0

End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns or sets the back color of the docking frame"

    BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

  Dim i As Integer

    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    LockWindowUpdate Extender.Parent.hwnd
    For i = 1 To Panels.Count
        Panels(i).BackColor = New_BackColor
    Next '  I I
    LockWindowUpdate ByVal 0&

End Property

Public Property Get BorderStyle() As tdBorderStyles
Attribute BorderStyle.VB_Description = "Returns or set the border style of the docked forms."

  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MemberInfo=21,0,0,0

    BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As tdBorderStyles)

  Dim i As Integer

    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    LockWindowUpdate Extender.Parent.hwnd
    For i = 1 To Panels.Count
        Panels(i).DockArrange
    Next '  I I
    LockWindowUpdate ByVal 0&

End Property

Public Property Get CaptionStyle() As tdCaptionStyles

    CaptionStyle = m_CaptionStyle

End Property

Public Property Let CaptionStyle(ByVal New_CaptionStyle As tdCaptionStyles)

  Dim i As Integer

    m_CaptionStyle = New_CaptionStyle
    PropertyChanged "CaptionStyle"
    LockWindowUpdate Extender.Parent.hwnd
    For i = 1 To Panels.Count
        Panels(i).DockArrange
        If Not Panels(i).Expanded Then
            Panels(i).RefreshCollaped
        End If
    Next '  I I
    LockWindowUpdate ByVal 0&

End Property

Public Property Get CollapseInterval() As Long

    CollapseInterval = m_CollapseInterval

End Property

Public Property Let CollapseInterval(ByVal intval As Long)

    m_CollapseInterval = intval
    PropertyChanged "CollapseInterval"

End Property

Public Sub DockChange(formName As String, _
                      newPanel As tdAlignProperty)

  Dim formIndex As Integer
  Dim newDockStyle As tdDockStyles
  Dim visCount As Integer
  Dim visCountCol As Integer

    formIndex = DockedFormIndex(formName)
    If formIndex Then
        LockWindowUpdate Extender.Parent.hwnd

        visCount = Me.DockedForms(formIndex).Panel.WindowList.VisibleCount
        visCountCol = Me.DockedForms(formIndex).Panel.WindowList.VisibleCountCollapsed
        Me.DockedForms(formIndex).DockForm_Hide
        Me.DockedForms(formIndex).Panel.WindowList.RemoveByHandle (Me.DockedForms(formIndex).hwnd)

        If Not Me.DockedForms(formIndex).Panel.Expanded Then
            Me.DockedForms(formIndex).Panel.RefreshCollaped
            If visCountCol < 2 Then
                Me.DockedForms(formIndex).Panel.Visible = False
            End If
        End If

        Set Me.DockedForms(formIndex).Panel = Panels(newPanel)

        Select Case newPanel
          Case tdAlignTop
            newDockStyle = tdDockTop
          Case tdAlignBottom
            newDockStyle = tdDockBottom
          Case tdAlignLeft
            newDockStyle = tdDockLeft
          Case tdAlignRight
            newDockStyle = tdDockRight
        End Select

        Me.DockedForms(formIndex).Style = Me.DockedForms(formIndex).Style Or newDockStyle
        Me.DockedForms(formIndex).DockForm_Dock
        Me.DockedForms(formIndex).DockForm_Show
        LockWindowUpdate 0&
    End If

End Sub

Public Function DockedFormCaptionHeight()

    DockedFormCaptionHeight = getCaptionHeight

End Function

Public Function DockedFormCaptionOffset(DockedFormName As String) As Integer

    If IsFormDockedTopBottom(DockedFormName) Then
        DockedFormCaptionOffset = (getCaptionHeight + 4) * Screen.TwipsPerPixelX
      Else
        DockedFormCaptionOffset = 0
    End If

End Function

Public Function DockedFormCaptionOffsetBottom(DockedFormName As String) As Integer

    If IsFormDocked(DockedFormName) Then
        If IsFormDockedTopBottom(DockedFormName) Then
            DockedFormCaptionOffsetBottom = 8 * Screen.TwipsPerPixelY
          Else
            DockedFormCaptionOffsetBottom = (getCaptionHeight + 15) * Screen.TwipsPerPixelY
        End If
      Else
        DockedFormCaptionOffsetBottom = 0
    End If

End Function

Public Function DockedFormCaptionOffsetLeft(DockedFormName As String) As Integer

    If IsFormDocked(DockedFormName) Then
        If IsFormDockedTopBottom(DockedFormName) Then
            DockedFormCaptionOffsetLeft = (getCaptionHeight + 4) * Screen.TwipsPerPixelX
          Else
            DockedFormCaptionOffsetLeft = 4 * Screen.TwipsPerPixelX
        End If
      Else
        DockedFormCaptionOffsetLeft = 0
    End If

End Function

Public Function DockedFormCaptionOffsetRight(DockedFormName As String) As Integer

    If IsFormDocked(DockedFormName) Then
        If IsFormDockedTopBottom(DockedFormName) Then
            DockedFormCaptionOffsetRight = (getCaptionHeight + 8) * Screen.TwipsPerPixelX
          Else
            DockedFormCaptionOffsetRight = 8 * Screen.TwipsPerPixelX
        End If
      Else
        DockedFormCaptionOffsetRight = 0
    End If

End Function

Public Function DockedFormCaptionOffsetTop(DockedFormName As String) As Integer

    If IsFormDocked(DockedFormName) Then
        If IsFormDockedTopBottom(DockedFormName) Then
            DockedFormCaptionOffsetTop = 4 * Screen.TwipsPerPixelY
          Else
            DockedFormCaptionOffsetTop = (getCaptionHeight + 11) * Screen.TwipsPerPixelY
        End If
      Else
        DockedFormCaptionOffsetTop = 0
    End If

End Function

Public Function DockedFormIndex(DockedFormName As String) As Integer

  Dim formItem  As Integer
  Dim formFound As Boolean

    formItem = 1
    formFound = False
    For formItem = 1 To Me.DockedForms.Count
        If Me.DockedForms.Item(formItem).Key = DockedFormName Then
            formFound = True
            Exit For
        End If
    Next formItem
    If formFound Then
        DockedFormIndex = formItem
      Else
        DockedFormIndex = -1
    End If

End Function

Public Property Get DockedForms() As TDockForms
Attribute DockedForms.VB_Description = "Collection of docked forms"

    Set DockedForms = m_DockedForms

End Property

Public Sub FormDock(Index As Variant)
Attribute FormDock.VB_Description = "Docks a form in its panel host"

  Const constSource As String = m_constClassName & ".FormDock"

    On Error GoTo Err_FormDock
    If IsObject(Index) Then
        m_DockedForms.ItemByHandle(Index.hwnd).DockForm_Dock
      Else
        m_DockedForms(Index).DockForm_Dock
    End If

Exit Sub

Err_FormDock:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Sub FormHide(Index As Variant)

  Const constSource As String = m_constClassName & ".FormHide"

    On Error GoTo Err_FormHide
    If IsObject(Index) Then
        m_DockedForms.ItemByHandle(Index.hwnd).DockForm_Hide
      Else
        m_DockedForms(Index).DockForm_Hide
    End If

Exit Sub

Err_FormHide:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Sub FormShow(Index As Variant)
Attribute FormShow.VB_Description = "Shows a docked form"

  Const constSource As String = m_constClassName & ".FormShow"

    On Error GoTo Err_FormShow
    If IsObject(Index) Then
        m_DockedForms.ItemByHandle(Index.hwnd).DockForm_Show
        If Not m_DockedForms.ItemByHandle(Index.hwnd).Panel.Expanded Then
            m_DockedForms.ItemByHandle(Index.hwnd).Panel.RefreshCollaped
        End If
      Else
        m_DockedForms(Index).DockForm_Show
        If Not m_DockedForms(Index).Panel.Expanded Then
            m_DockedForms(Index).Panel.RefreshCollaped
        End If
    End If

Exit Sub

Err_FormShow:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Sub FormUndock(Index As Variant)
Attribute FormUndock.VB_Description = "Undocks a form from its panel host"

  Const constSource As String = m_constClassName & ".FormUndock"

    On Error GoTo Err_FormUndock
    If IsObject(Index) Then
        m_DockedForms.ItemByHandle(Index.hwnd).DockForm_UnDock
      Else
        m_DockedForms(Index).DockForm_UnDock
    End If

Exit Sub

Err_FormUndock:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Sub GrabMain(MainFormHwnd As Long)

  Dim hWndMdi As Long

    NewHWND = MainFormHwnd
    AttachMessage Me, NewHWND, WM_ACTIVATEAPP
    hWndMdi = FindWindowEx(GetParent(MainFormHwnd), UserControl.hwnd, "MDIClient", ByVal 0&)
    Set m_Size = New cSizer
    If (hWndMdi = 0) Then
        If SmartSizing Then
            m_Size.NoFullDrag = True
          Else
            m_Size.NoFullDrag = False
        End If
        m_Size.Attach NewHWND
        'debug.print "Sizable"
        Exit Sub
      ElseIf (Parent.BorderStyle = vbSizable Or Parent.BorderStyle = vbSizableToolWindow) Then
        If SmartSizing Then
            m_Size.NoFullDrag = True
          Else
            m_Size.NoFullDrag = False
        End If
        m_Size.Attach NewHWND
        'debug.print "Sizable"
      Else
        m_Size.NoFullDrag = False
        'debug.print "not Sizable"
    End If

End Sub

Public Property Let Gradient1(ByVal Grad As OLE_COLOR)

    m_Gradient1 = Grad
    PropertyChanged "Grad1"
    GradClr1 = Grad

End Property

Public Property Get Gradient1() As OLE_COLOR

    Gradient1 = m_Gradient1

End Property

Public Property Get Gradient2() As OLE_COLOR

    Gradient2 = m_Gradient2

End Property

Public Property Let Gradient2(ByVal Grad As OLE_COLOR)

    m_Gradient2 = Grad
    PropertyChanged "Grad2"
    GradClr2 = Grad

End Property

Public Function IsFormDocked(DockedFormName As String) As Boolean

  Dim formItem  As Integer
  Dim formFound As Boolean

    formItem = 1
    formFound = False
    For formItem = 1 To Me.DockedForms.Count
        If Me.DockedForms.Item(formItem).Key = DockedFormName Then
            If Me.DockedForms.Item(formItem).State = tdDocked Then
                formFound = True
                Exit For
            End If
        End If
    Next formItem
    IsFormDocked = formFound

End Function

Public Function IsFormDockedTopBottom(DockedFormName As String) As Boolean

  Dim formItem  As Integer
  Dim formFound As Boolean

    formItem = 1
    formFound = False
    For formItem = 1 To Me.DockedForms.Count
        If Me.DockedForms.Item(formItem).Key = DockedFormName Then
            If Me.DockedForms.Item(formItem).State = tdDocked Then
                If Me.DockedForms.Item(formItem).Panel.Align = tdAlignTop Or Me.DockedForms.Item(formItem).Panel.Align = tdAlignBottom Then
                    formFound = True
                End If
                Exit For
            End If
        End If
    Next formItem
    IsFormDockedTopBottom = formFound

End Function

Private Function IsFormLoaded(hWndA As Long) As Boolean

  Const constSource As String = m_constClassName & ".IsFormLoaded"
  Dim i             As Integer

    On Error GoTo Err_IsFormLoaded
    For i = 1 To m_DockedForms.Count
        If m_DockedForms(i).hwnd = hWndA Then
            IsFormLoaded = True
            Exit Function
        End If
    Next '  I I
    IsFormLoaded = False

Exit Function

Err_IsFormLoaded:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Function

Private Property Get ISubClass_MsgResponse() As EMsgResponse

    Select Case CurrentMessage
      Case WM_ACTIVATEAPP
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

    Select Case iMsg
      Case WM_ACTIVATEAPP
        ' Form is activated/deactivated:
        If wParam = 1 Then
            m_MainApp = True
            Timer1(0).Enabled = True
            Timer1(1).Enabled = True
            Timer1(2).Enabled = True
            Timer1(3).Enabled = True
          Else
            m_MainApp = False
            Timer1(0).Enabled = False
            Timer1(1).Enabled = False
            Timer1(2).Enabled = False
            Timer1(3).Enabled = False
        End If
      Case Else
        'debug.print "imsg"; iMsg
    End Select

End Function

Private Function LoadControl(oForm As Object, _
                             CtlType As String, _
                             ctlName As String, _
                             Optional CtlContainer) As Object

  Dim oCtl As Object

    ' ******************************************************************************
    ' Routine       : (Function) LoadControl
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 12/06/2000 - 22:22:42
    ' Inputs        :
    ' Outputs       :
    ' Credits       : This code was extract from
    '                 FreeVBCode.com (http://www.freevbcode.com)
    ' Modifications :
    ' Description   : Load a form control at run-time
    ' ******************************************************************************
    On Error Resume Next
        If IsObject(oForm.Controls) Then
            If IsMissing(CtlContainer) Then
                Set oCtl = oForm.Controls.Add(CtlType, ctlName)
              Else
                Set oCtl = oForm.Controls.Add(CtlType, ctlName, CtlContainer)
            End If
            If Not oCtl Is Nothing Then
                Set LoadControl = oCtl
                Set oCtl = Nothing
            End If
        End If
    On Error GoTo 0

End Function

Private Sub LoadPanels()

  Const constSource As String = m_constClassName & ".LoadPanels"
  Dim i             As Integer
  Dim pict          As VB.PictureBox
  Dim NewWidth      As Long
  Dim NewHeight     As Long

    ' ******************************************************************************
    ' Routine       : (Sub) LoadPanels
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 12/06/2000 - 22:11:07
    ' Inputs        : N/A
    ' Outputs       : N/A
    ' Modifications :
    ' Description   : Load the panels for the docking system
    ' ******************************************************************************
    On Error GoTo Err_LoadPanels
    ' only to avoid panels re-loading
    If m_bLoaded Then
        Exit Sub
    End If
    ' loop to create the 4 panels (left, top, right, bottom panels)
    For i = 1 To 4
        ' add a picture box at run-time to the extender (form)
        Set pict = LoadControl(Extender.Parent, "VB.PictureBox", "Host" & CStr(i))
        pict.BackColor = m_BackColor
        ' add a new panel to the list, the container
        If Persistant Then
            NewHeight = GetSetting(App.Title, "Panels", i & "Height", m_PanelHeight)
            NewWidth = GetSetting(App.Title, "Panels", i & "Width", m_PanelWidth)
          Else 'NOT PERSISTANT...
            NewHeight = m_PanelHeight
            NewWidth = m_PanelWidth
        End If
        ' will be our picture box
        m_Panels.Add i, NewHeight, NewWidth, False, Me, pict, "Host" & CStr(i)
        m_Panels(m_Panels.Count).AutoShowCaptionOnCollapse = m_AutoShowCaptionOnCollapse
                
        Select Case m_Panels(i).Align
            Case tdAlignTop
                m_Panels(i).PanelSizing = m_PanelResizeTop
                m_Panels(i).DockedFormSizing = m_PanelTopDockFormResize
            Case tdAlignBottom
                m_Panels(i).PanelSizing = m_PanelResizeBottom
                m_Panels(i).DockedFormSizing = m_PanelBottomDockFormResize
            Case tdAlignLeft
                m_Panels(i).PanelSizing = m_PanelResizeLeft
                m_Panels(i).DockedFormSizing = m_PanelLeftDockFormResize
            Case tdAlignRight
                m_Panels(i).PanelSizing = m_PanelResizeRight
                m_Panels(i).DockedFormSizing = m_PanelRightDockFormResize
        End Select
        
        Timer1(i - 1).Enabled = True
        Timer1(i - 1).Interval = m_CollapseInterval
    Next '  I I
    m_bLoaded = True

Exit Sub

Err_LoadPanels:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub m_Size_EnterSizeMove()

    m_bSizing = True

End Sub

Private Sub m_Size_ExitSizeMove()

  Dim i As Integer

    On Error Resume Next
        m_bSizing = False
        For i = 1 To m_Panels.Count
            m_Panels(i).Repaint
            If Not m_Panels(i).Expanded Then
                m_Panels(i).RefreshCollaped
            End If
        Next '  I I
    On Error GoTo 0

End Sub

Public Property Get MainAppActive() As Boolean

    MainAppActive = m_MainApp

End Property

Public Property Let MaximizeButton(maxButton As Boolean)

    m_MaximizeButton = maxButton
    PropertyChanged "MaximizeButton"

End Property

Public Property Get MaximizeButton() As Boolean

    MaximizeButton = m_MaximizeButton

End Property

Public Property Let PanelHeight(ByVal New_PanelHeight As Long)
Attribute PanelHeight.VB_Description = "Returns or sets the initial height of top and bottom panels"

    If Ambient.UserMode Then
        Err.Raise 382
    End If
    m_PanelHeight = New_PanelHeight
    PropertyChanged "PanelHeight"

End Property

Public Property Get PanelHeight() As Long

  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MemberInfo=8,1,0,2100

    PanelHeight = m_PanelHeight

End Property

Public Property Get Panels() As TTabDockHosts
Attribute Panels.VB_Description = "Panels of the docking system"

    Set Panels = m_Panels

End Property

Public Property Get PanelWidth() As Long
Attribute PanelWidth.VB_Description = "Returns or sets a initial Width for the left and right panels"

  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MemberInfo=8,1,0,1000

    PanelWidth = m_PanelWidth

End Property

Public Property Let PanelWidth(ByVal New_PanelWidth As Long)

    If Ambient.UserMode Then
        Err.Raise 382
    End If
    m_PanelWidth = New_PanelWidth
    PropertyChanged "PanelWidth"

End Property

Public Property Get Parent() As Object
Attribute Parent.VB_Description = "Generally this is the MDI form the control was dropped in"

    Set Parent = Extender.Parent

End Property

Public Property Get Persistant() As Boolean

    Persistant = m_Persistant

End Property

Public Property Let Persistant(ByVal Persist As Boolean)

    m_Persistant = Persist
    PropertyChanged "Persistant"

End Property

Private Function Repaint_Panels()

  Dim x As Integer

    On Error Resume Next
        For x = 1 To m_Panels.Count
            m_Panels(x).Repaint
            If Not m_Panels(x).Expanded Then
                m_Panels(x).Panel_Expand
            End If
        Next '  X X
    On Error GoTo 0

End Function

Public Function resetTimer(Index As Integer)

    Timer1(Index - 1).Enabled = True

End Function

Public Sub Show()
Attribute Show.VB_Description = "Show the host panels and update docked forms"

  Const constSource As String = m_constClassName & ".Show"
  Dim i             As Integer

    ' ******************************************************************************
    ' Routine       : (Sub) Show
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 12/06/2000 - 22:22:13
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : Show panels and forms docked/undocked
    ' ******************************************************************************
    On Error GoTo Err_Show
    ' let's avoid some flickering...
    LockWindowUpdate Extender.Parent.hwnd
    ' dock/undock the forms
    For i = 1 To m_DockedForms.Count
        If (m_DockedForms(i).Style And tdShowInvisible) = False Then
            ' it it it is docked then dock it
            If m_DockedForms(i).State = tdDocked Then
                m_DockedForms(i).Panel.Dock m_DockedForms(i)
              Else
                ' just show
                m_DockedForms(i).Panel.UnDock m_DockedForms(i)
            End If
        End If
    Next '  I I
    ' free willy! (I mean windows!)
    LockWindowUpdate 0

Exit Sub

Err_Show:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Public Property Get SmartSizing() As Boolean

    SmartSizing = m_bSmartSizing

End Property

Public Property Let SmartSizing(bln As Boolean)

    m_bSmartSizing = bln
    PropertyChanged "SmartSizing"

End Property

Private Sub Timer1_Timer(Index As Integer)

  Dim lpPosition As POINTAPI
  Dim panelX1    As Integer
  Dim panelY1    As Integer
  Dim panelX2    As Integer
  Dim panelY2    As Integer

    ' get the current position of the pointer
    GetCursorPos lpPosition
    'convert relative to the main form
    ScreenToClient NewHWND, lpPosition
    If m_bSizing Then
        'debug.print "We Are Sizing"
        Exit Sub
    End If
    If m_Panels.Count > 0 Then
        panelY1 = m_Panels(Index + 1).Top / Screen.TwipsPerPixelY
        panelX1 = m_Panels(Index + 1).Left / Screen.TwipsPerPixelX
        panelY2 = panelY1 + (m_Panels(Index + 1).Height / Screen.TwipsPerPixelY)
        panelX2 = panelX1 + (m_Panels(Index + 1).width / Screen.TwipsPerPixelX)
        If lpPosition.x >= panelX1 And lpPosition.x <= panelX2 And lpPosition.Y >= panelY1 And lpPosition.Y <= panelY2 Then
            Exit Sub
        End If
        '   m_Panels(Index + 1).Expanded = False
        Select Case (Index + 1)
          Case tdAlignTop
            If m_AutoCollapseTop Then
                m_Panels(Index + 1).Panel_Collapse
                UserControl.Parent.SetFocus
                Timer1(Index).Enabled = False
            End If
          Case tdAlignLeft
            If m_AutoCollapseLeft Then
                m_Panels(Index + 1).Panel_Collapse
                Timer1(Index).Enabled = False
                UserControl.Parent.SetFocus
            End If
          Case tdAlignBottom
            If m_AutoCollapseBottom Then
                m_Panels(Index + 1).Panel_Collapse
                Timer1(Index).Enabled = False
                UserControl.Parent.SetFocus
            End If
          Case tdAlignRight
            If m_AutoCollapseRight Then
                m_Panels(Index + 1).Panel_Collapse
                Timer1(Index).Enabled = False
                UserControl.Parent.SetFocus
            End If
        End Select
    End If

End Sub

Friend Sub TriggerEvent(ByVal RaisedEvent As String, _
       ParamArray aParams() As Variant)

  Const constSource As String = m_constClassName & ".TriggerEvent"

    ' ******************************************************************************
    ' Routine       : (Sub) TriggerEvent
    ' Created by    : Marclei V Silva
    ' Company Name  : Spnorte Consultoria
    ' Machine       : ZEUS
    ' Date-Time     : 12/06/2000 - 22:20:56
    ' Inputs        :
    ' Outputs       :
    ' Modifications :
    ' Description   : Used to raise events to the form user
    ' ******************************************************************************
    On Error GoTo Err_TriggerEvent
    Select Case RaisedEvent
      Case "Dock"
        RaiseEvent FormDocked(aParams(0))
      Case "UnDock"
        RaiseEvent FormUnDocked(aParams(0))
      Case "ShowForm"
        RaiseEvent FormShow(aParams(0))
      Case "HideForm"
        RaiseEvent FormHide(aParams(0))
      Case "ResizePanel"
        RaiseEvent PanelResize(aParams(0))
      Case "MenuClick"
        RaiseEvent MenuClick(aParams(0))
      Case "PanelClick"
        RaiseEvent PanelClick(aParams(0))
      Case "CaptionClick"
        RaiseEvent CaptionClick(aParams(0), aParams(1), aParams(2), aParams(3))
      Case Else
        ''debug.print "Event no handled: " & RaisedEvent
    End Select

Exit Sub

Err_TriggerEvent:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub UserControl_Initialize()

  Const constSource As String = m_constClassName & ".UserControl_Initialize"

    On Error GoTo Err_UserControl_Initialize
    Set m_DockedForms = New TDockForms
    Set m_Panels = New TTabDockHosts

Exit Sub

Err_UserControl_Initialize:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub UserControl_InitProperties()

  'Initialize Properties for User Control

    m_BackColor = m_def_BackColor
    m_BorderStyle = m_def_BorderStyle
    m_CaptionStyle = m_def_CaptionStyle
    m_PanelHeight = m_def_PanelHeight
    m_PanelWidth = m_def_PanelWidth
    m_Visible = m_def_Visible
    m_Persistant = m_PersistantDef
    Persist = m_PersistantDef
    Gradient1 = m_Grad1
    Gradient2 = m_Grad2

End Sub

Private Sub UserControl_Paint()

  Const constSource As String = m_constClassName & ".UserControl_Paint"
  Dim Edge          As RECT                                         ' Rectangle edge of control

    On Error GoTo Err_UserControl_Paint
    Edge.Left = 0                                   ' Set rect edges to outer
    Edge.Top = 0                                    ' most position in pixels
    Edge.Bottom = 32 'ScaleHeight
    Edge.Right = 32 ' ScaleWidth
    DrawEdge hdc, Edge, BDR_RAISEDOUTER, BF_RECT Or BF_SOFT ' Draw Edge...

Exit Sub

Err_UserControl_Paint:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  'Load property values from storage

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_CaptionStyle = PropBag.ReadProperty("CaptionStyle", m_def_CaptionStyle)
    m_MaximizeButton = PropBag.ReadProperty("MaximizeButton", False)
    m_AutoExpand = PropBag.ReadProperty("AutoExpand", True)
    m_AutoCollapseTop = PropBag.ReadProperty("AutoCollapseTop", False)
    m_AutoCollapseLeft = PropBag.ReadProperty("AutoCollapseLeft", False)
    m_AutoCollapseBottom = PropBag.ReadProperty("AutoCollapseBottom", False)
    m_AutoCollapseRight = PropBag.ReadProperty("AutoCollapseRight", False)
    m_CollapseInterval = PropBag.ReadProperty("CollapseInterval", 3000)
    m_PanelHeight = PropBag.ReadProperty("PanelHeight", m_def_PanelHeight)
    m_PanelWidth = PropBag.ReadProperty("PanelWidth", m_def_PanelWidth)
    m_Visible = PropBag.ReadProperty("Visible", m_def_Visible)
    m_Persistant = PropBag.ReadProperty("Persistant", m_PersistantDef)
    m_Gradient2 = PropBag.ReadProperty("Gradient2", m_Gradient2)
    m_Gradient1 = PropBag.ReadProperty("Gradient1", m_Gradient1)
    m_AutoShowCaptionOnCollapse = PropBag.ReadProperty("AutoShowCollapseCaptions", True)
    m_bSmartSizing = PropBag.ReadProperty("SmartSizing", False)
    m_PanelResizeTop = PropBag.ReadProperty("PanelResizeTop", False)
    m_PanelResizeLeft = PropBag.ReadProperty("PanelResizeLeft", False)
    m_PanelResizeRight = PropBag.ReadProperty("PanelResizeRight", False)
    m_PanelResizeBottom = PropBag.ReadProperty("PanelResizeBottom", False)
    
    m_PanelBottomDockFormResize = PropBag.ReadProperty("PanelBottomDockFormResize", False)
    m_PanelTopDockFormResize = PropBag.ReadProperty("PanelTopDockFormResize", False)
    m_PanelLeftDockFormResize = PropBag.ReadProperty("PanelLeftDockFormResize", False)
    m_PanelRightDockFormResize = PropBag.ReadProperty("PanelRightDockFormResize", False)
    
    
    GradClr1 = m_Gradient1
    GradClr2 = m_Gradient2
    Persist = m_Persistant
    m_MainApp = True

End Sub

Private Sub UserControl_Resize()

  Const constSource As String = m_constClassName & ".UserControl_Resize"

    On Error GoTo Err_UserControl_Resize
    ' set the control to 32 pixels wide
    UserControl.width = 32 * Screen.TwipsPerPixelX
    UserControl.Height = 32 * Screen.TwipsPerPixelY

Exit Sub

Err_UserControl_Resize:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub UserControl_Terminate()

  Dim i             As Integer
  Dim x             As Integer
  Const constSource As String = m_constClassName & ".UserControl_Terminate"

    On Error GoTo Err_UserControl_Terminate
    For x = 1 To m_DockedForms.Count
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "Width", m_DockedForms(x).width
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "Height", m_DockedForms(x).Height
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "Style", m_DockedForms(x).Style
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "State", m_DockedForms(x).State
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "FloatWidth", m_DockedForms(x).FloatingWidth
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "FloatHeight", m_DockedForms(x).FloatingHeight
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "FloatLeft", m_DockedForms(x).FloatingLeft
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "FloatTop", m_DockedForms(x).FloatingTop
        SaveSetting App.Title, "Docking", m_DockedForms(x).Key & "Align", m_DockedForms(x).Panel.Align
    Next '  X X
    For i = 1 To m_Panels.Count
        SaveSetting App.Title, "Panels", i & "Width", m_Panels(i).width
        SaveSetting App.Title, "Panels", i & "Height", m_Panels(i).Height
    Next '  I I
    Set m_Panels = Nothing
    Set m_DockedForms = Nothing
    Timer1(0).Enabled = False
    Timer1(1).Enabled = False
    Timer1(2).Enabled = False
    Timer1(3).Enabled = False
    Timer1(0).Interval = 0
    Timer1(1).Interval = 0
    Timer1(2).Interval = 0
    Timer1(3).Interval = 0
    DetachMessage Me, NewHWND, WM_ACTIVATEAPP
    If Not m_Size Is Nothing Then
        m_Size.Detach
        Set m_Size = Nothing
    End If

Exit Sub

Err_UserControl_Terminate:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  'Write property values to storage

    Call PropBag.WriteProperty("AutoShowCollapseCaptions", m_AutoShowCaptionOnCollapse, True)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, m_def_CaptionStyle)
    Call PropBag.WriteProperty("MaximizeButton", m_MaximizeButton, False)
    Call PropBag.WriteProperty("AutoExpand", m_AutoExpand, True)
    Call PropBag.WriteProperty("AutoCollapseTop", m_AutoCollapseTop, False)
    Call PropBag.WriteProperty("AutoCollapseLeft", m_AutoCollapseLeft, False)
    Call PropBag.WriteProperty("AutoCollapseBottom", m_AutoCollapseBottom, False)
    Call PropBag.WriteProperty("AutoCollapseRight", m_AutoCollapseRight, False)
    Call PropBag.WriteProperty("CollapseInterval", m_CollapseInterval, 3000)
    Call PropBag.WriteProperty("PanelHeight", m_PanelHeight, m_def_PanelHeight)
    Call PropBag.WriteProperty("PanelWidth", m_PanelWidth, m_def_PanelWidth)
    Call PropBag.WriteProperty("Visible", m_Visible, m_def_Visible)
    Call PropBag.WriteProperty("Persistant", m_Persistant, m_PersistantDef)
    Call PropBag.WriteProperty("Gradient1", m_Gradient1, m_Grad1)
    Call PropBag.WriteProperty("Gradient2", m_Gradient2, m_Grad2)
    Call PropBag.WriteProperty("SmartSizing", m_bSmartSizing, False)
    Call PropBag.WriteProperty("PanelResizeTop", m_PanelResizeTop, False)
    Call PropBag.WriteProperty("PanelResizeBottom", m_PanelResizeBottom, False)
    Call PropBag.WriteProperty("PanelResizeLeft", m_PanelResizeLeft, False)
    Call PropBag.WriteProperty("PanelResizeRight", m_PanelResizeRight, False)
    
    Call PropBag.WriteProperty("PanelBottomDockFormResize", m_PanelBottomDockFormResize, False)
    Call PropBag.WriteProperty("PanelTopDockFormResize", m_PanelTopDockFormResize, False)
    Call PropBag.WriteProperty("PanelLeftDockFormResize", m_PanelLeftDockFormResize, False)
    Call PropBag.WriteProperty("PanelRightDockFormResize", m_PanelRightDockFormResize, False)
    
    
End Sub

Public Property Let Visible(ByVal New_Visible As Boolean)
Attribute Visible.VB_Description = "Show/Hides the docking system frame"

    If Ambient.UserMode = False Then
        Err.Raise 387
    End If
    m_Visible = New_Visible
    PropertyChanged "Visible"
    LockWindowUpdate Extender.Parent.hwnd
    If New_Visible Then
        Timer1(0).Enabled = True
        Timer1(1).Enabled = True
        Timer1(2).Enabled = True
        Timer1(3).Enabled = True
        If m_Panels(tdAlignLeft).Expanded = False Then
            m_Panels(tdAlignLeft).Panel_Expand
        End If
        If m_Panels(tdAlignRight).Expanded = False Then
            m_Panels(tdAlignRight).Panel_Expand
        End If
        If m_Panels(tdAlignTop).Expanded = False Then
            m_Panels(tdAlignTop).Panel_Expand
        End If
        If m_Panels(tdAlignBottom).Expanded = False Then
            m_Panels(tdAlignBottom).Panel_Expand
        End If
      Else 'NOT NEW_VISIBLE...
        Timer1(0).Enabled = False
        Timer1(1).Enabled = False
        Timer1(2).Enabled = False
        Timer1(3).Enabled = False
    End If
    m_Panels(tdAlignLeft).Visible = New_Visible
    m_Panels(tdAlignRight).Visible = New_Visible
    m_Panels(tdAlignTop).Visible = New_Visible
    m_Panels(tdAlignBottom).Visible = New_Visible
    LockWindowUpdate ByVal 0&

End Property

Public Property Get Visible() As Boolean

  'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
  'MemberInfo=0,0,2,0

    Visible = m_Visible

End Property


Private Function ChangePanelRezise(tdAlign As AlignConstants, bln As Boolean)
On Error Resume Next

Dim i As Integer

        For i = 1 To m_Panels.Count
            If m_Panels(i).Align = tdAlign Then
                m_Panels(i).PanelSizing = bln
                Exit For
            End If
        Next

On Error GoTo 0
End Function
Private Function ChangePanelDockFormRezise(tdAlign As AlignConstants, bln As Boolean)
On Error Resume Next

Dim i As Integer

        For i = 1 To m_Panels.Count
            If m_Panels(i).Align = tdAlign Then
                m_Panels(i).DockedFormSizing = bln
                Exit For
            End If
        Next

On Error GoTo 0
End Function


