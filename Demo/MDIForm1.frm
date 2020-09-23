VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\Source\TabDock.vbp"
Begin VB.MDIForm MDIForm1 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "TabDock Control Demo Application"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9705
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDock.TTabDock TabDock 
      Left            =   1320
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      BorderStyle     =   10
      CaptionStyle    =   15
      AutoCollapseBottom=   -1  'True
      CollapseInterval=   7000
      Gradient1       =   0
      SmartSizing     =   -1  'True
      PanelResizeLeft =   -1  'True
      PanelResizeRight=   -1  'True
      PanelBottomDockFormResize=   -1  'True
      PanelTopDockFormResize=   -1  'True
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6540
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "TabDockHost v 1.0"
            TextSave        =   "TabDockHost v 1.0"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &1"
         Index           =   0
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &2"
         Index           =   1
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &3"
         Index           =   2
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &4"
         Index           =   3
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &6"
         Index           =   4
      End
      Begin VB.Menu mnuViewForm 
         Caption         =   "Form &7"
         Index           =   5
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPanels 
         Caption         =   "Panels"
      End
   End
   Begin VB.Menu mnuDocking 
      Caption         =   "&Dock"
      Begin VB.Menu mnuDockForm 
         Caption         =   "Form &1"
         Index           =   0
      End
      Begin VB.Menu mnuDockForm 
         Caption         =   "Form &2"
         Index           =   1
      End
      Begin VB.Menu mnuDockForm 
         Caption         =   "Form &6"
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuTBHelp 
         Caption         =   "TabDock Help"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPUDockable 
         Caption         =   "Dockable"
      End
      Begin VB.Menu mnuPUHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPUAbout 
         Caption         =   "About TabDock..."
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ******************************************************************************
' Project       : DemoTabDOck.vbp
' Created by    : Marclei V Silva
' Company Name  : Spnorte Consultoria
' Machine       : ZEUS
' Date-Time     : 25/06/2000 - 1:47:49
' Modifications : See documents in this prokect
' Description   : Demo project for TabDock Control
' ******************************************************************************
Option Explicit
Option Compare Text


Public Sub LoadNewDoc(strFileName As String)

    Dim frmNewDoc As Form5
    Dim i         As Integer
    Dim bLoaded   As Boolean

    ' This will load a MDI Child form. Note the
    ' interaction it has with our component
    ' if there is any pending opened form
    ' just unload them all
    On Error Resume Next
    bLoaded = False
    For i = 0 To Forms.Count - 1
        If Forms(i).Tag = strFileName Then
            bLoaded = True
            Set frmNewDoc = Forms(i)
            Exit For
        End If
    Next '  I
    If Not bLoaded Then
        Set frmNewDoc = New Form5
        frmNewDoc.Caption = strFileName
        frmNewDoc.Tag = strFileName
        On Error Resume Next
        frmNewDoc.rtfText.LoadFile strFileName
    End If
    frmNewDoc.Show
    frmNewDoc.SetFocus
    On Error GoTo 0

End Sub

Private Sub MDIForm_Load()

    ' This has no more purpose than to assign this form's hwnd
    ' to the control so when a docked form is on focus the main form
    ' is too. If you don't do this thats the only feature you'll loose.

    TabDock.GrabMain Me.hWnd
    ' load border styles into combo1
    SetBorderStyles
    ' load caption styles
    SetCaptionStyles
    ' Let´s add some forms to the TabDock Control
    ' Here you can set the state, style, docking alignment
    ' and Key properties.
    TabDock.AddForm Form1, tdDocked, tdAlignBottom, "Form1", tdDockBottom Or tdDockFloat, 60, True, False
    ' Use the key properties to retrive information
    ' about the docked form later see FormDocked() and
    ' FormUnDocked() events below
    TabDock.AddForm Form2, tdDocked, tdAlignLeft, "Form2", tdDockBottom Or tdDockTop Or tdDockFloat Or tdDockLeft, , True, True
    ' Form 3 may not float and can only dock on the left panel
    ' if you want a fixed form this is the way you should add it
    TabDock.AddForm Form3, tdDocked, tdAlignRight, "Form3", tdDockRight, 30, True, False
    ' Form 4 can only dock on the top panel and can not float
    ' note that Form4 has no border
    TabDock.AddForm Form4, tdDocked, tdAlignTop, "Form4", tdDockTop Or tdDockFloat, , True, True
    ' Form 6 will also have style property set
    ' Let's make it dock only on left and right panels
    ' and allow floating or else it will not be able to dock
    ' in any other panel
    TabDock.AddForm Form6, tdDocked, tdAlignTop, "Form6", tdDockTop Or tdDockBottom Or tdDockFloat, , False, True
    ' Form 7 docks in Right panel and cannot float
    TabDock.AddForm Form7, tdDocked, tdAlignRight, "Form7", tdDockRight Or tdDockFloat, 70, False, False
    ' right panel is fixed
    '    TabDock.Panels(tdAlignRight).Resizable = False
    ' let's change right panel width before showing the docking system
    'TabDock.Panels(tdAlignRight).Width = 3800
    ' let's change top panel back color
    '    TabDock.Panels(tdAlignTop).BackColor = vbButtonShadow
    ' After you've had added your forms, the Docking
    ' system engine will do the heavy job for you
    ' you don't have to configure anything else.
    ' call this method to show the TabTock Panels
    TabDock.Show
    ' This subrotine will set up the menu items
    ' based on the startup configuration we've made
    ' right above
    SetupMenu
    ' set version info...
    StatusBar1.SimpleText = App.ProductName & " - version " & App.Major & "." & App.Minor & "." & App.Revision
    ' load default doc
    LoadNewDoc App.Path & "\readme.rtf"

End Sub



Private Sub MDIForm_Unload(Cancel As Integer)

    On Error Resume Next
    ' Let´s clean it up. Many app may hang if you do not
    ' unload all forms reference
    Unload Form1
    Unload Form2
    Unload Form3
    Unload Form4
    Unload Form6
    Unload Form7
    UnloadAll
    On Error GoTo 0

End Sub

Private Sub mnuAbout_Click()

    Dim strAbout As String

    ' build about message
    strAbout = "TabDock Control version " & App.Major & "." & App.Minor & vbNewLine & "by Marclei V Silva" & vbNewLine & _
              "email: marclei@spnorte.com" & vbNewLine & "Home Page: http://www.spnorte.com" & vbNewLine & vbNewLine & _
              "Description" & vbNewLine & "Dock your forms like" & vbNewLine & "Vb IDE forms does without all the hazards!" & vbNewLine & _
              vbNewLine & "WaRnInG! WaRnInG! WaRnInG! WaRnInG! WaRnInG! WaRnInG!" & vbNewLine & "This control is still under development, so please," & vbNewLine & _
              "avoid using this in commercial applications" & vbNewLine

    ' show about message
    MsgBox strAbout

End Sub

Private Sub mnuDockForm_Click(Index As Integer)

    Dim Key As String

    ' Another feature of the TabDock Host is to
    ' dock and undock the form dynamically
    mnuDockForm(Index).Checked = Not mnuDockForm(Index).Checked
    ' Select the form you wish to operate with
    Select Case Index
        Case 0
        Key = "Form1"
        Case 1
        Key = "Form2"
        Case 2
        Key = "Form6"
    End Select
    ' Now toggle docking
    If mnuDockForm(Index).Checked Then
        TabDock.FormDock Key
        Else 'MNUDOCKFORM(INDEX).CHECKED = FALSE/0
        TabDock.FormUndock Key
    End If

End Sub

Private Sub mnuFileExit_Click()

    ' Unload all forms

    Unload Me

End Sub

Private Sub mnuFileNew_Click()

    ' Load a new document

    LoadNewDoc ""

End Sub

Private Sub mnuFileSave_Click()

    ' save the readme file as text

    If ActiveForm Is Nothing Then
        Exit Sub
    End If
    ActiveForm.rtfText.SaveFile ActiveForm.Tag, 0
    
    
    Dim f As TDockForm
    
    Set f = TabDock.DockedForms("Form2")
    f.Panel.Panel_Collapse
    
    
    
    
    
    

End Sub

Private Sub mnuPUAbout_Click()

    ' show about box

    mnuAbout_Click

End Sub

Private Sub mnuPUDockable_Click()

    ' if is docked then undock else dock it ;-)

    If TabDock.DockedForms(mnuPopup.Tag).State = tdDocked Then
        TabDock.FormUndock mnuPopup.Tag
        Else 'NOT TABDOCK.DOCKEDFORMS(MNUPOPUP.TAG).STATE...
        TabDock.FormDock mnuPopup.Tag
    End If

End Sub

Private Sub mnuPUHide_Click()

    ' hide the current form

    TabDock.FormHide mnuPopup.Tag

End Sub

Private Sub mnuViewForm_Click(Index As Integer)

    Dim Key As String

    ' This is a simple use of the TabDock Host.
    ' Based on the menu clicked item we will hide or
    ' show the selected form
    mnuViewForm(Index).Checked = Not mnuViewForm(Index).Checked
    ' Select the form you wish to operate with
    Select Case Index
        Case 0
        Key = "Form1"
        Case 1
        Key = "Form2"
        Case 2
        Key = "Form3"
        Case 3
        Key = "Form4"
        Case 4
        Key = "Form6"
        Case 5
        Key = "Form7"
    End Select
    ' Now toggle visibility
    If mnuViewForm(Index).Checked Then
        TabDock.FormShow Key
        Else 'MNUVIEWFORM(INDEX).CHECKED = FALSE/0
        TabDock.FormHide Key
    End If

End Sub

Private Sub mnuViewPanels_Click()

    ' Here we can access the Panels and perform
    ' many interesting actions

    mnuViewPanels.Checked = Not mnuViewPanels.Checked
    ' let's toggle panels visibility
    TabDock.Visible = mnuViewPanels.Checked

End Sub

Private Sub SetBorderStyles()

    ' We will add some border styles, so that the user may select
    ' the one that is more interesting

    With Form6.Combo1
        .Clear
        .AddItem "0 - None"
        .AddItem "1 - RaisedOuter"
        .AddItem "2 - RaisedInner"
        .AddItem "3 - Raised"
        .AddItem "4 - SunkenOuter"
        .AddItem "5 - SunkenInner"
        .AddItem "6 - Sunken"
        .AddItem "7 - Etched"
        .AddItem "8 - Bump"
        .AddItem "9 - Mono"
        .AddItem "10- Flat"
        .AddItem "11- Soft"
    End With 'FORM6.COMBO1
    ' select current style
    Form6.Combo1.ListIndex = TabDock.BorderStyle

End Sub

Private Sub SetCaptionStyles()

    ' We will add some border styles, so that the user may select
    ' the one that is more interesting

    With Form6.Combo2
        .Clear
        .AddItem "0 - Normal"
        .AddItem "1 - Etched"
        .AddItem "2 - Soft"
        .AddItem "3 - Raised"
        .AddItem "4 - RaisedInner"
        .AddItem "5 - SunkenOuter"
        .AddItem "6 - Sunken"
        .AddItem "7 - SingleRaisedBar"
        .AddItem "8 - Gradient"
        .AddItem "9 - SingleRaisedInner"
        .AddItem "10 - SingleSoft"
        .AddItem "11 - SingleEtched"
        .AddItem "12 - SingleSunken"
        .AddItem "13 - SingleSunkenOuter"
        .AddItem "14 - OfficeXP"
        .AddItem "15 - VS.Net"
    End With 'FORM6.COMBO2
    ' select current style
    Form6.Combo2.ListIndex = TabDock.CaptionStyle

End Sub

Private Sub SetupMenu()

    ' Check all items of the View menu once
    ' all forms are initially visible

    mnuViewForm(0).Checked = True
    mnuViewForm(1).Checked = False
    mnuViewForm(2).Checked = True
    mnuViewForm(3).Checked = True
    mnuViewForm(4).Checked = True
    mnuViewForm(5).Checked = True
    ' Here let's update the docking state menu items
    mnuDockForm(0).Checked = True
    mnuDockForm(1).Checked = False
    mnuDockForm(2).Checked = True
    ' panels are initially visible
    mnuViewPanels.Checked = True

End Sub

Private Sub TabDock_CaptionClick(ByVal DockedForm As TabDock.TDockForm, _
                                  ByVal Button As Integer, _
                                  ByVal X As Single, _
                                  ByVal Y As Single)

    ' This event replace previous context menu items, so that
    ' you can personalize the menus the way you want
    ' check if it was a right click

    If Button <> vbRightButton Then
        Exit Sub
    End If
    ' save form key for further use
    mnuPopup.Tag = DockedForm.Key
    ' enable this item only for floating forms
    mnuPUDockable.Enabled = (DockedForm.Style And tdDockFloat)
    ' check dockable menu
    mnuPUDockable.Checked = (DockedForm.State = tdDocked)
    ' show popup menu
    PopupMenu mnuPopup

End Sub

Private Sub TabDock_FormDocked(ByVal DockedForm As TabDock.TDockForm)

    ' based on the form key we can take any action we want

    Select Case DockedForm.Key
        Case "Form1"
        mnuDockForm(0).Checked = True
        Case "Form2"
        mnuDockForm(1).Checked = True
        Case "Form6"
        mnuDockForm(2).Checked = True
        'Debug.Print "Docked"
        Form6.Cls
    End Select

End Sub

Private Sub TabDock_FormHide(ByVal DockedForm As TabDock.TDockForm)

    ' based on the form key we can take any action we want

    Select Case DockedForm.Key
        Case "Form1"
        mnuViewForm(0).Checked = False
        Case "Form2"
        mnuViewForm(1).Checked = False
        Case "Form3"
        mnuViewForm(2).Checked = False
        Case "Form4"
        mnuViewForm(3).Checked = False
        Case "Form6"
        mnuViewForm(4).Checked = False
        Case "Form7"
        mnuViewForm(5).Checked = False
    End Select

End Sub

Private Sub TabDock_FormShow(ByVal DockedForm As TabDock.TDockForm)

    ' based on the form key we can take any action we want

    Select Case DockedForm.Key
        Case "Form1"
        mnuViewForm(0).Checked = True
        Case "Form2"
        mnuViewForm(1).Checked = True
        Case "Form3"
        mnuViewForm(2).Checked = True
        Case "Form4"
        mnuViewForm(3).Checked = True
        Case "Form6"
        mnuViewForm(4).Checked = True
        Case "Form7"
        mnuViewForm(5).Checked = True
    End Select

End Sub

Private Sub TabDock_FormUnDocked(ByVal DockedForm As TabDock.TDockForm)

    ' based on the form key we can take any action we want

    Select Case DockedForm.Key
        Case "Form1"
        mnuDockForm(0).Checked = False
        Case "Form2"
        mnuDockForm(1).Checked = False
        Case "Form6"
        mnuDockForm(2).Checked = False
    End Select

End Sub

Private Sub TabDock_PanelClick(ByVal Panel As TabDock.TTabDockHost)

    If Panel.Align = tdAlignRight Then
        MsgBox "This panel is fixed. It cannot be moved"
    End If

End Sub

Private Sub TabDock_PanelResize(ByVal Panel As TabDock.TTabDockHost)

    ' if active panel align = tdAlignTop

    If Panel.Align = tdAlignTop Then
        ' show a custom message
        'MsgBox "Panel resize event was fired. " & vbCrLf & "Top panel can not be resized!" & vbCrLf & "Click OK to restore panel size", vbInformation
        ' restore panel height
        'Panel.Height = 900
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    '-- end code


End Sub

Private Sub UnloadAll()

    Dim i As Integer

    ' if there is any pending opened form
    ' just unload them all
    On Error Resume Next
    For i = 0 To Forms.Count - 1
        Unload Forms(i)
    Next '  I
    On Error GoTo 0

End Sub
