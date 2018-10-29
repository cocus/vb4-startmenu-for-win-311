VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "CProgmanGroupItem"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Option Explicit

Private mvarPointX As Integer
Private mvarPointY As Integer
Private mvarIcon As Integer
Private mvarcbHeader As Integer
Private mvarcbANDPlane As Integer
Private mvarcbXORPlane As Integer
Private mvarpHeader As Integer
Private mvarpANDPlane As Integer
Private mvarpXORPlane As Integer
Private mvarItemName As String
Private mvarItemCommand As String
Private mvarItemIconPath As String

Public Property Let ItemIconPath(ByVal vData As String)
    mvarItemIconPath = vData
End Property

Public Property Get ItemIconPath() As String
    ItemIconPath = mvarItemIconPath
End Property



Public Property Let ItemCommand(ByVal vData As String)
    mvarItemCommand = vData
End Property

Public Property Get ItemCommand() As String
    ItemCommand = mvarItemCommand
End Property



Public Property Let ItemName(ByVal vData As String)
    mvarItemName = vData
End Property

Public Property Get ItemName() As String
    ItemName = mvarItemName
End Property



Public Property Let pXORPlane(ByVal vData As Integer)
    mvarpXORPlane = vData
End Property

Public Property Get pXORPlane() As Integer
    pXORPlane = mvarpXORPlane
End Property



Public Property Let pANDPlane(ByVal vData As Integer)
    mvarpANDPlane = vData
End Property

Public Property Get pANDPlane() As Integer
    pANDPlane = mvarpANDPlane
End Property



Public Property Let pHeader(ByVal vData As Integer)
    mvarpHeader = vData
End Property

Public Property Get pHeader() As Integer
    pHeader = mvarpHeader
End Property



Public Property Let cbXORPlane(ByVal vData As Integer)
    mvarcbXORPlane = vData
End Property

Public Property Get cbXORPlane() As Integer
    cbXORPlane = mvarcbXORPlane
End Property



Public Property Let cbANDPlane(ByVal vData As Integer)
    mvarcbANDPlane = vData
End Property

Public Property Get cbANDPlane() As Integer
    cbANDPlane = mvarcbANDPlane
End Property



Public Property Let cbHeader(ByVal vData As Integer)
    mvarcbHeader = vData
End Property

Public Property Get cbHeader() As Integer
    cbHeader = mvarcbHeader
End Property



Public Property Let Icon(ByVal vData As Integer)
    mvarIcon = vData
End Property

Public Property Get Icon() As Integer
    Icon = mvarIcon
End Property



Public Property Let PointY(ByVal vData As Integer)
    mvarPointY = vData
End Property

Public Property Get PointY() As Integer
    PointY = mvarPointY
End Property



Public Property Let PointX(ByVal vData As Integer)
    mvarPointX = vData
End Property

Public Property Get PointX() As Integer
    PointX = mvarPointX
End Property




Public Sub Dump()
    Debug.Print " ===================================== "
    Debug.Print "PointX = "; PointX
    Debug.Print "PointY = "; PointY
    Debug.Print "Icon = "; Icon
    Debug.Print "cbHeader = "; cbHeader
    Debug.Print "cbANDPlane = "; cbANDPlane
    Debug.Print "cbXORPlane = "; cbXORPlane
    Debug.Print "pHeader = "; pHeader
    Debug.Print "pANDPlane = "; pANDPlane
    Debug.Print "pXORPlane = "; pXORPlane
    Debug.Print "ItemName = "; ItemName
    Debug.Print "ItemCommand = "; ItemCommand
    Debug.Print "ItemIconPath = "; ItemIconPath
End Sub

