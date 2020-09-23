VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private iInterval     As Long
Private id            As Long
' User can attach any Variant data they want to the timer
Public Item           As Variant
Attribute Item.VB_VarDescription = "Gets/sets a piece of additional data to store with the timer."
Public Event ThatTime()
Attribute ThatTime.VB_Description = "Raised when the timer fires."
' SubTimer is independent of VBCore, so it hard codes error handling
Public Enum EErrorTimer
    eeBaseTimer = 13650     ' CTimer
    eeTooManyTimers         ' No more than 10 timers allowed per class
    eeCantCreateTimer       ' Can't create system timer
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private eeBaseTimer, eeTooManyTimers, eeCantCreateTimer
#End If

Private Fake_To_Protect_ECP_at_end_of_Declarations As Boolean ' Move_Enum_Away_from_End_ofDeclarations_And_Delete_Me

Private Sub Class_Terminate()

    Interval = 0

End Sub

Friend Sub ErrRaise(e As Long)

  Dim sText   As String
  Dim sSource As String

    If e > 1000 Then
        sSource = App.EXEName & ".WindowProc"
        Select Case e
          Case eeTooManyTimers
            sText = "No more than 10 timers allowed per class"
          Case eeCantCreateTimer
            sText = "Can't create system timer"
        End Select
        Err.Raise e Or vbObjectError, sSource, sText
      Else
        ' Raise standard Visual Basic error
        Err.Raise e, sSource
    End If

End Sub

Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Gets/sets the interval at which the timer fires.  Set to zero to stop the timer."

    Interval = iInterval

End Property

Public Property Let Interval(iIntervalA As Long)

  Dim f As Boolean

    ' Can't just change interval--you must kill timer and start a new one
    If iIntervalA > 0 Then
        ' Don't mess with it if interval is the same
        If iInterval = iIntervalA Then
            Exit Property
        End If
        ' Must destroy any existing timer to change interval
        If iInterval Then
            f = TimerDestroy(Me)
            'debug.Assert f      ' Shouldn't fail
        End If
        ' Create new timer with new interval
        iInterval = iIntervalA
        If TimerCreate(Me) = False Then
            ErrRaise eeCantCreateTimer
        End If
      Else
        If (iInterval > 0) Then
            iInterval = 0
            f = TimerDestroy(Me)
            'debug.Assert f      ' Shouldn't fail
        End If
    End If

End Property

Public Sub PulseTimer()
Attribute PulseTimer.VB_MemberFlags = "40"

  ' Must be public so that Timer object can't terminate while client's ThatTime
  ' event is being processed--Friend wouldn't prevent this disaster

    RaiseEvent ThatTime

End Sub

Friend Property Get TimerID() As Long

    TimerID = id

End Property

Friend Property Let TimerID(idA As Long)

    id = idA

End Property

