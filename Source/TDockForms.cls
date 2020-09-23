VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TDockForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DockForm"
Attribute VB_Ext_KEY = "Member0" ,"DockForm"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ******************************************************************************
' Class:             DockedForm
' Description:       Collection of Docked Forms, used by TabDock
'                    to keep track of the forms that are part of the
'                    engine.
' Created by:        Marclei V Silva
' Machine:           ZEUS
' Date-Time:         09/05/2000 3:07:33
' Last modification: 09/05/2000 3:07:33
' ******************************************************************************
Option Explicit
' local variable to hold collection
Private mCol                       As Collection
Private WithEvents ctlcommand      As VB.PictureBox
Attribute ctlcommand.VB_VarHelpID = -1

Friend Function Add(ByVal Item As Form, _
       ByVal Panel As TTabDockHost, _
       Style As tdDockStyles, _
       State As tdDockedState, _
       sKey As String, _
       hwnd As Long, _
       iPercent As Integer, _
       Optional bHasMaxButton As Boolean = False, _
       Optional bHasCloseButton As Boolean = False) As TDockForm

  Dim objNewMember As TDockForm
  Dim lWidth       As Long
  Dim lHeight      As Long
  Dim lLeft        As Long
  Dim lTop         As Long

    ' create a new object
    Set objNewMember = New TDockForm
    ' although we are passing the object references here, only
    ' a pointer to them will be saved for further tracking of the object
    ' this avoid vb creation of multiple references to the object
    ' and minimizes the chances of crashing when we finishing the
    ' application
    Set objNewMember.Panel = Panel
    Set objNewMember.Extender = Item
    'If (Style And tdDockFloat) = tdDockFloat Then
    'objNewMember.AttachHitTester Item
    'End If
    ' set the properties passed into the method
    If Persist Then
        lWidth = (GetSetting(App.Title, "Docking", sKey & "FloatWidth", Item.width / Screen.TwipsPerPixelX))
        lHeight = (GetSetting(App.Title, "Docking", sKey & "FloatHeight", Item.Height / Screen.TwipsPerPixelY))
        lLeft = GetSetting(App.Title, "Docking", sKey & "FloatLeft", Item.Left / Screen.TwipsPerPixelX)
        lTop = GetSetting(App.Title, "Docking", sKey & "FloatTop", Item.Top / Screen.TwipsPerPixelY)
        State = GetSetting(App.Title, "Docking", sKey & "State", State)
        Style = GetSetting(App.Title, "Docking", sKey & "Style", Style)
      Else 'NOT PERSIST...
        lWidth = Item.width / Screen.TwipsPerPixelX
        lHeight = Item.Height / Screen.TwipsPerPixelY
        lLeft = Item.Left / Screen.TwipsPerPixelX
        lTop = Item.Top / Screen.TwipsPerPixelY
    End If
    objNewMember.Style = Style
    objNewMember.State = State
    objNewMember.FloatingWidth = lWidth
    objNewMember.FloatingHeight = lHeight
    objNewMember.FloatingLeft = lLeft
    objNewMember.FloatingTop = lTop
    objNewMember.ParentHWND = hwnd
    objNewMember.Key = sKey
    objNewMember.HasMaxButton = bHasMaxButton
    objNewMember.HasCloseButton = bHasCloseButton
    ' new property percent - default is 10
    If iPercent = 0 Then
        objNewMember.Percent = 10
      Else
        objNewMember.Percent = iPercent
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
      Else
        mCol.Add objNewMember, sKey
    End If
    ' return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Friend Sub Append(DockForm As TDockForm, _
       Index As Integer)

    If Index <= mCol.Count Then
        mCol.Add DockForm, , Index
      Else
        mCol.Add DockForm
    End If

End Sub

Private Sub Class_Initialize()

  ' creates the collection when this class is created

    Set mCol = New Collection

End Sub

Private Sub Class_Terminate()

  ' destroys collection when this class is terminated

    Set mCol = Nothing

End Sub

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of items in the collection"

    Count = mCol.Count

End Property

Private Sub ctlcommand_Click()

  '-- end code

    Beep

End Sub

Friend Function HandleExists(hWndA As Long) As Boolean

  Dim Index As Integer

    For Index = 1 To mCol.Count
        If mCol(Index).hwnd = hWndA Then
            HandleExists = True
            Exit Function
        End If
    Next Index
    HandleExists = False

End Function

Public Property Get Item(Index As Variant) As TDockForm
Attribute Item.VB_Description = "Returns a docked form from the list"
Attribute Item.VB_UserMemId = 0
On Error Resume Next

    Set Item = mCol(Index)

End Property

Friend Function ItemByHandle(hWndA As Long) As TDockForm

  Dim Index As Integer

    For Index = 1 To mCol.Count
        If mCol(Index).hwnd = hWndA Then
            Set ItemByHandle = mCol(Index)
            Exit Function
        End If
    Next Index

End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mCol.[_NewEnum]

End Property

Public Sub Remove(Index As Variant)
Attribute Remove.VB_Description = "Removes a item (DockForm) from the collection list"

    mCol.Remove Index

End Sub

Friend Sub RemoveByHandle(hWndA As Long)

  Dim Index As Integer

    ' ******************************************************************************
    ' Routine       : RemoveByHandle
    ' Created by    : Marclei V Silva
    ' Machine       : ZEUS
    ' Date-Time     : 23/12/0020:20:03
    ' Inputs        : hWndA : The form handle
    ' Outputs       : nothing
    ' Credits       :
    ' Modifications :
    ' Description   : This procedure was create because
    '                 we can not reference the key attribute once
    '                 we do not know wether the user supplied one
    '                 This come to rescue us from this.
    ' ******************************************************************************
    For Index = 1 To mCol.Count
        If mCol(Index).hwnd = hWndA Then
            mCol.Remove Index
            Exit Sub
        End If
    Next Index

End Sub

Friend Function VisibleCount() As Integer

  Dim Count As Integer
  Dim w     As TDockForm

    'Friend Function GetIndex(DockForm As TDockForm) As Integer
    '    Dim Index As Integer
    '    Dim Rc As RECT
    '    Dim Rw As RECT
    '
    '    ' let's DockArrange window list when a new window is added to
    '    ' the list
    '    GetWindowRect DockForm.hWnd, Rc
    '    For Index = 1 To mCol.Count
    '        GetWindowRect mCol(Index).hWnd, Rw
    '        If Rw.Top > Rc.Top Then
    '            Exit For
    '        End If
    '    Next
    '    GetIndex = Index
    'End Function
    Count = 0
    For Each w In mCol
        If w.Visible Then
            Count = Count + 1
        End If
    Next '  W W
    VisibleCount = Count

End Function

Friend Function VisibleCountCollapsed() As Integer

  Dim Count As Integer
  Dim w     As TDockForm

    Count = 0
    For Each w In mCol
        If w.Tag Then
            Count = Count + 1
        End If
    Next w
    VisibleCountCollapsed = Count

End Function

