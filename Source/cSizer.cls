VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ===================================================================
' Class: cSizer.cls
' Created: Gary Noble
' Description: Displays A System Size Box When A Form Is Moved Or Sized
' Uses: This Class Takes Extensive Use Of The System Api To Manipulate
'       Window Movement.
'       Uses The SPI (SystemParametersInfo) Call To Take Care of The
'       Rectangle And Movement Params
' ===================================================================
Private Type RECT
    Left                                          As Long
    Top                                           As Long
    Right                                         As Long
    Bottom                                        As Long
End Type
Private Const SPI_GETDRAGFULLWINDOWS          As Integer = 38
Private Const SPI_SETDRAGFULLWINDOWS          As Integer = 37
Private Const SPIF_SENDWININICHANGE           As Long = &H2
Private Const SPIF_UPDATEINIFILE              As Long = &H1
Private Const WM_DESTROY                      As Long = &H2
Private Const WM_SIZING                       As Long = &H214
Private Const WM_MOVING                       As Long = &H216&
Private Const WM_ENTERSIZEMOVE                As Long = &H231&
Private Const WM_EXITSIZEMOVE                 As Long = &H232&
Private Const WM_ACTIVATE                     As Long = &H6
Private Const WM_SIZE                         As Long = &H5
Private m_hWnd                                As Long
Private m_bSwitchOff                          As Boolean
Private m_bNoFullDrag                         As Boolean
Private m_bActive                             As Boolean
Private m_bInSizeMove                         As Boolean
Implements ISubclass
Public Event EnterSizeMove()
Public Event ExitSizeMove()
Public Event Activate(ByVal bByMouse As Boolean)
Public Event Deactivate()
Public Event Moving(ByRef lLeft As Long, ByRef lTop As Long, ByRef lWidth As Long, ByRef lHeight As Long)
Public Event Sizing(ByRef lLeft As Long, ByRef lTop As Long, ByRef lWidth As Long, ByRef lHeight As Long)
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                          ByVal uParam As Long, _
                          lpvParam As Any, _
                          ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                          lpvSource As Any, _
                          ByVal cbCopy As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                          lpRect As RECT) As Long

Public Property Get Active() As Boolean

    Active = m_bActive

End Property

Public Sub Attach(ByVal hWndA As Long)

    Detach
    m_hWnd = hWndA
    AttachMessage Me, m_hWnd, WM_ENTERSIZEMOVE
    AttachMessage Me, m_hWnd, WM_EXITSIZEMOVE
    AttachMessage Me, m_hWnd, WM_DESTROY
    AttachMessage Me, m_hWnd, WM_MOVING
    AttachMessage Me, m_hWnd, WM_SIZING
    AttachMessage Me, m_hWnd, WM_ACTIVATE
    AttachMessage Me, m_hWnd, WM_SIZE

End Sub

Private Sub Class_Terminate()

    On Error Resume Next
        m_bSwitchOff = True
        pTurnOffFullDrag
        pResetFullDrag
    On Error GoTo 0

End Sub

Public Sub Detach()

    If Not m_hWnd = 0 Then
        DetachMessage Me, m_hWnd, WM_ENTERSIZEMOVE
        DetachMessage Me, m_hWnd, WM_EXITSIZEMOVE
        DetachMessage Me, m_hWnd, WM_DESTROY
        DetachMessage Me, m_hWnd, WM_MOVING
        DetachMessage Me, m_hWnd, WM_SIZING
        DetachMessage Me, m_hWnd, WM_SIZE
        DetachMessage Me, m_hWnd, WM_ACTIVATE
    End If

End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)

  '

End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse

  ' Process before windows:

    ISubClass_MsgResponse = emrPostProcess

End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

  Dim tR      As RECT
  Dim lWidth  As Long
  Dim lHeight As Long

    Select Case iMsg
      Case WM_ENTERSIZEMOVE
        ' Enter modal sizing/moving loop
        m_bInSizeMove = True
        RaiseEvent EnterSizeMove
      Case WM_EXITSIZEMOVE
        ' Exit modal sizing/moving loop
        RaiseEvent ExitSizeMove
        m_bInSizeMove = False
      Case WM_MOVING
        ' Form is moving:
        CopyMemory tR, ByVal lParam, Len(tR)
        lWidth = tR.Right - tR.Left
        lHeight = tR.Bottom - tR.Top
        RaiseEvent Moving(tR.Left, tR.Top, lWidth, lHeight)
        tR.Right = tR.Left + lWidth
        tR.Bottom = tR.Top + lHeight
        CopyMemory ByVal lParam, tR, Len(tR)
      Case WM_SIZING
        ' Form is sizing:
        CopyMemory tR, ByVal lParam, Len(tR)
        lWidth = tR.Right - tR.Left
        lHeight = tR.Bottom - tR.Top
        RaiseEvent Sizing(tR.Left, tR.Top, lWidth, lHeight)
        tR.Right = tR.Left + lWidth
        tR.Bottom = tR.Top + lHeight
        CopyMemory ByVal lParam, tR, Len(tR)
      Case WM_ACTIVATE
        ' Form is activated/deactivated:
        If wParam = 0 Then
            ' deactivate:
            pResetFullDrag
            m_bActive = False
            RaiseEvent Deactivate
          Else
            ' Activate
            pTurnOffFullDrag
            m_bActive = True
            RaiseEvent Activate(wParam = 2)
        End If
      Case WM_SIZE
        ' This ensures that you don't have to separately
        ' check for maximize/minimize/restore/in code
        ' movement, as these aren't
        ' controlled by the UI sizing loop:
        If Not (m_bInSizeMove) Then
            GetWindowRect m_hWnd, tR
            lWidth = tR.Right - tR.Left
            lHeight = tR.Bottom - tR.Top
            RaiseEvent Sizing(tR.Left, tR.Top, lWidth, lHeight)
        End If
      Case WM_DESTROY
        ' Ensures the class terminates regardless
        ' of whether the user explicitly detaches
        ' it or not:
        'debug.print "WM_DESTROY"
        pResetFullDrag
        Detach
    End Select

End Function

Public Property Let NoFullDrag(ByVal bState As Boolean)

    m_bNoFullDrag = bState
    If bState Then
        If m_bActive Then
            pTurnOffFullDrag
        End If
      Else
        pResetFullDrag
    End If

End Property

Public Property Get NoFullDrag() As Boolean

    NoFullDrag = m_bNoFullDrag

End Property

Private Sub pResetFullDrag()

    If m_bSwitchOff Then
        SystemParametersInfo SPI_SETDRAGFULLWINDOWS, 1&, ByVal 0&, SPIF_SENDWININICHANGE
        m_bSwitchOff = False
    End If

End Sub

Private Sub pTurnOffFullDrag()

  Dim lR As Long

    If m_bNoFullDrag Then
        If Not (SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0&, lR, 0) = 0) Then
            If Not lR = 0 Then
                m_bSwitchOff = True
                lR = SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, 0&, ByVal 0&, SPIF_SENDWININICHANGE)
            End If
        End If
    End If

End Sub

