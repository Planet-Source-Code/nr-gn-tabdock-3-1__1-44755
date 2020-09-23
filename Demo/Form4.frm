VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Tabbed View"
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   3750
   ControlBox      =   0   'False
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   552
      Left            =   300
      ScaleHeight     =   555
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   660
      Width           =   2952
      Begin VB.Label Label1 
         Caption         =   $"Form4.frx":058A
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3252
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   1332
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3492
      _ExtentX        =   6165
      _ExtentY        =   2355
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Readme"
            Object.ToolTipText     =   "Show TabDock Control ReadMe file"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Comments"
            Object.ToolTipText     =   "Show user comments about the control"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Revisions"
            Object.ToolTipText     =   "Revisions made to this control"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "MDIForm"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    On Error Resume Next
    TabStrip1.Move MDIForm1.TabDock.DockedFormCaptionOffsetLeft(Me.Name), MDIForm1.TabDock.DockedFormCaptionOffsetTop(Me.Name), Me.ScaleWidth - MDIForm1.TabDock.DockedFormCaptionOffsetRight(Me.Name), Me.ScaleHeight - MDIForm1.TabDock.DockedFormCaptionOffsetBottom(Me.Name)
    Picture1.Move TabStrip1.Left + 20, TabStrip1.Top + 350, TabStrip1.Width - 50, TabStrip1.Height - 400
    Label1.Move 50, 50, Picture1.ScaleWidth - 100, Picture1.ScaleHeight - 100
    On Error GoTo 0

End Sub

Private Sub Form_Resize_old()

    Dim formItem            As Integer
    Dim formDockedTopBottom As Boolean

    On Error Resume Next
    ' resize controls
    On Error Resume Next
    formItem = MDIForm1.TabDock.DockedFormIndex(Me.Name)
    formDockedTopBottom = MDIForm1.TabDock.IsFormDockedTopBottom(Me.Name)
    If Not formItem Then
        If formDockedTopBottom Then
            TabStrip1.Move (MDIForm1.TabDock.DockedFormCaptionHeight + 4) * Screen.TwipsPerPixelY, 30, Me.ScaleWidth - ((MDIForm1.TabDock.DockedFormCaptionHeight + 7) * Screen.TwipsPerPixelY), Abs(Me.ScaleHeight - 60)
            Else 'FORMDOCKEDTOPBOTTOM = FALSE/0
            TabStrip1.Move 10, 30, Me.ScaleWidth, Me.ScaleHeight
        End If
    End If
    Picture1.Move TabStrip1.Left + 20, TabStrip1.Top + 350, TabStrip1.Width - 50, TabStrip1.Height - 400
    Label1.Move 50, 50, Picture1.ScaleWidth - 100, Picture1.ScaleHeight - 100
    On Error GoTo 0

End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.Index
        Case 1
        ' load readme file
        MDIForm1.LoadNewDoc App.Path & "\readme.rtf"
        Case 2
        ' load comments file
        MDIForm1.LoadNewDoc App.Path & "\comments.rtf"
        Case 3
        ' load revisions file
        MDIForm1.LoadNewDoc App.Path & "\revisions.rtf"
        Case 4
        ' load revisions file
        MDIForm1.LoadNewDoc App.Path & "\project.rtf"
    End Select
    On Error Resume Next
    Me.SetFocus
    On Error GoTo 0

End Sub

'-- end code

