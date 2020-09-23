VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Immediate"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   3180
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":025C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":036E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()

    On Error Resume Next
    Text1.Move MDIForm1.TabDock.DockedFormCaptionOffsetLeft(Me.Name), MDIForm1.TabDock.DockedFormCaptionOffsetTop(Me.Name), Me.ScaleWidth - MDIForm1.TabDock.DockedFormCaptionOffsetRight(Me.Name), Me.ScaleHeight - MDIForm1.TabDock.DockedFormCaptionOffsetBottom(Me.Name)
    On Error GoTo 0

End Sub
