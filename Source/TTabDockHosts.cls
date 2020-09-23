VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TTabDockHosts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TTabHost"
Attribute VB_Ext_KEY = "Member0" ,"TTabHost"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' ********************************************************************
' Project     :  TabDock
' Module      :  TTabDockHosts.cls
' Description :  Collection of Hosts
' Created by  :  Marclei V Silva
' Machine     :  ZEUS
' Date-Time   :  12/05/2000 0:40:08
' ********************************************************************
Option Explicit
' local variable to hold collection
Private mCol     As Collection

Public Function Add(Align As Integer, _
                    Height As Variant, _
                    width As Variant, _
                    Visible As Boolean, _
                    ByVal Parent As TTabDock, _
                    ByVal Container As Object, _
                    Optional sKey As String) As TTabDockHost

  Dim objNewMember As TTabDockHost

    ' create a new object
    Set objNewMember = New TTabDockHost
    ' set the properties passed into the method
    Container.FillStyle = 1
    Container.FillColor = &H0&
    Container.ForeColor = &H80000012
    Set objNewMember.Container = Container
    '    objNewMember.BackColor = &H8000000F
    objNewMember.BorderStyle = 0
    objNewMember.Height = Height
    objNewMember.width = width
    objNewMember.Align = Align
    objNewMember.Visible = Visible
    objNewMember.Index = mCol.Count + 1
    Set objNewMember.Parent = Parent
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
      Else
        mCol.Add objNewMember, sKey
    End If
    ' return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Private Sub Class_Initialize()

  'creates the collection when this class is created

    Set mCol = New Collection

End Sub

Private Sub Class_Terminate()

  'destroys collection when this class is terminated

    Set mCol = Nothing

End Sub

Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of hosts in the collection"

  '-- end code

    Count = mCol.Count

End Property

Public Property Get Item(Index As Variant) As TTabDockHost
Attribute Item.VB_Description = "Returns a TabDockHost object reference"
Attribute Item.VB_UserMemId = 0

    Set Item = mCol(Index)

End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"

    Set NewEnum = mCol.[_NewEnum]

End Property

Public Sub Remove(Index As Variant)

    mCol.Remove Index

End Sub

