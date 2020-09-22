VERSION 5.00
Begin VB.UserControl uctBuildQuery 
   BackColor       =   &H8000000D&
   CanGetFocus     =   0   'False
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   990
   ClipControls    =   0   'False
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   Picture         =   "uctBuildQuery.ctx":0000
   ScaleHeight     =   705
   ScaleWidth      =   990
   ToolboxBitmap   =   "uctBuildQuery.ctx":25B6
   Windowless      =   -1  'True
End
Attribute VB_Name = "uctBuildQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=====================================================
'Control        Easy query builder
'-----------------------------------------------------
'Written By C.Robb.
'=====================================================
Option Explicit

'______________________________
'Private members
Private oInsertQuery As EasyQueries.InsertQuery
Private oUpdateQuery As EasyQueries.UpdateQuery
Private oSelectQuery As EasyQueries.SelectQuery
Private oDeleteQuery As EasyQueries.DeleteQuery

'======================================================
'Initialise and terminate events
'======================================================
Private Sub UserControl_Initialize()
    Set oInsertQuery = New EasyQueries.InsertQuery
    Set oUpdateQuery = New EasyQueries.UpdateQuery
    Set oSelectQuery = New EasyQueries.SelectQuery
    Set oDeleteQuery = New EasyQueries.DeleteQuery
End Sub

'=======================================================
'Resize the control
'=======================================================
Private Sub UserControl_Resize()
    Height = 705
    Width = 990
End Sub

Private Sub UserControl_Terminate()
    If Not oInsertQuery Is Nothing Then Set oInsertQuery = Nothing
    If Not oUpdateQuery Is Nothing Then Set oUpdateQuery = Nothing
    If Not oSelectQuery Is Nothing Then Set oSelectQuery = Nothing
    If Not oDeleteQuery Is Nothing Then Set oDeleteQuery = Nothing
End Sub

'======================================================
'Expose the insertquery object
'======================================================
Public Property Get InsertQuery() As EasyQueries.InsertQuery
    Set InsertQuery = oInsertQuery
End Property

'======================================================
'Expose the updatequery object
'======================================================
Public Property Get UpdateQuery() As EasyQueries.UpdateQuery
    Set UpdateQuery = oUpdateQuery
End Property

'======================================================
'Expose the selectquery object
'======================================================
Public Property Get SelectQuery() As EasyQueries.SelectQuery
    Set SelectQuery = oSelectQuery
End Property

'======================================================
'Expose the deletequery object
'======================================================
Public Property Get DeleteQuery() As EasyQueries.DeleteQuery
    Set DeleteQuery = oDeleteQuery
End Property

