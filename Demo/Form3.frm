VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form Layout"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4890
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      Picture         =   "Form3.frx":058A
      ScaleHeight     =   2295
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   300
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

  'debug.Print "FORM3:MOUSEDOWN"

End Sub

Private Sub Form_MouseUp(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)

  'debug.Print "FORM3:MOUSEUP"

End Sub

Private Sub Form_Resize()

    On Error Resume Next
        Picture1.Move 40, Picture1.Top, Me.ScaleWidth - 80, Me.ScaleHeight - (Picture1.Top + 20)
        '  Picture1.Move 250 + (ScaleWidth / 2) - (Picture1.Width / 2), 200 + (ScaleHeight / 2) - (Picture1.Height / 2)
    On Error GoTo 0

End Sub

'-- end code

':) Ulli's VB Code Formatter V2.14.7 (02/05/2003 16:50:14) 1 + 31 = 32 Lines
