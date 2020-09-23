VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   BackColor       =   &H80000016&
   Caption         =   "Members"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":058A
            Key             =   "closed"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2415
      Left            =   435
      TabIndex        =   0
      Top             =   0
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   4260
      _Version        =   393217
      Indentation     =   265
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    With TreeView1.Nodes
        .Add , , , "Item 1", "closed", "closed"
        .Add 1, tvwChild, , "SubItem 1", "closed", "closed"
        .Add , , , "Item 2", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 1", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 2", "closed", "closed"
        .Add 3, tvwChild, , "SubItem 3", "closed", "closed"
    End With 'TREEVIEW1.NODES

End Sub

Private Sub Form_Resize()

    '-- end code

    On Error Resume Next
    TreeView1.Move MDIForm1.TabDock.DockedFormCaptionOffsetLeft(Me.Name), MDIForm1.TabDock.DockedFormCaptionOffsetTop(Me.Name), Me.ScaleWidth - MDIForm1.TabDock.DockedFormCaptionOffsetRight(Me.Name), Me.ScaleHeight - MDIForm1.TabDock.DockedFormCaptionOffsetBottom(Me.Name)
    On Error GoTo 0

End Sub

Private Sub Form_Resize_old()

    Dim formItem            As Integer
    Dim formDockedTopBottom As Boolean

    On Error Resume Next
    formItem = MDIForm1.TabDock.DockedFormIndex(Me.Name)
    formDockedTopBottom = MDIForm1.TabDock.IsFormDockedTopBottom(Me.Name)
    If Not formItem Then
        If formDockedTopBottom Then
            TreeView1.Move (MDIForm1.TabDock.DockedFormCaptionHeight + 4) * Screen.TwipsPerPixelY, 30, Me.ScaleWidth - ((MDIForm1.TabDock.DockedFormCaptionHeight + 7) * Screen.TwipsPerPixelY), Abs(Me.ScaleHeight - 60)
            Else 'FORMDOCKEDTOPBOTTOM = FALSE/0
            TreeView1.Move 10, 30, Me.ScaleWidth, Me.ScaleHeight
        End If
    End If
    On Error GoTo 0

End Sub
