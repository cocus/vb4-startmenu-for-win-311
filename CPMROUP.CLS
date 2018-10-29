VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
END
Attribute VB_Name = "CProgmanGroup"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECTS
    left        As Integer
    top         As Integer
    right       As Integer
    bottom      As Integer
End Type

Private Type POINTS
    x           As Integer
    y           As Integer
End Type

Private Type GROUPDEF
    dwMagic     As Long
    wCheckSum   As Integer
    cbGroup     As Integer
    nCmdShow    As Integer
    rcNormal    As RECTS
    ptMin       As POINTS
    pName       As Integer
    cxIcon      As Integer
    cyIcon      As Integer
    wIconFormat As Integer
    wReserved   As Integer
    cItems      As Integer
End Type

Private Type tProgmanGroup
    grpdef      As GROUPDEF
    Name        As String
    Items()     As Integer
End Type



Private Type ITEMDEF
    pt As POINTS
    iIcon As Integer
    cbHeader As Integer
    cbANDPlane As Integer
    cbXORPlane As Integer
    pHeader As Integer
    pANDPlane As Integer
    pXORPlane As Integer
    pName As Integer
    pCommand As Integer
    pIconPath As Integer
End Type


Private Type tProgmanItem
    itmdef      As ITEMDEF
    Name        As String
    Command     As String
    IconPath    As String
End Type


Private c_tItems                As Collection

Private c_tGrpDef               As tProgmanGroup

Private Function ReadStringFromFile(ByVal iFile As Integer, ByVal iMaxSize As Integer, Optional ByVal iPos As Variant) As String
    If Not (IsMissing(iPos) Or (iPos = 0)) Then
        Seek #iFile, iPos
    End If

    Dim bData           As Byte

    Do
        Get iFile, , bData
       
        If (bData = 0) Or _
           (iMaxSize = 0) Then
            Exit Do
        End If

        iMaxSize = iMaxSize - 1

        ReadStringFromFile = ReadStringFromFile + Chr(bData)
    Loop
End Function

Public Function ReadGroup(ByVal sPath As String) As Boolean
    Dim iFile                   As Integer
    Dim iIndex                  As Integer
    Dim iItemOffset             As Integer
    
    Dim tGrpItm                 As ITEMDEF
    
    Dim tClass                  As CProgmanGroupItem

    On Error GoTo ErrHandler

    '// Get an unused file number
    iFile = FreeFile

    Open sPath For Binary Access Read As #iFile
        With c_tGrpDef
            Get #iFile, , .grpdef

            With c_tGrpDef.grpdef
                '// Read the Group Name
                If .pName Then
                    'Seek #1, tGrpDef.grpdef.pName + 1
                    c_tGrpDef.Name = ReadStringFromFile(iFile, 30, .pName + 1)
                End If

                Debug.Print "dwMagic = "; Hex(.dwMagic)
                Debug.Print "wCheckSum = "; .wCheckSum
                Debug.Print "cbGroup = "; .cbGroup
                Debug.Print "nCmdShow = "; .nCmdShow
                Debug.Print "rcNormal = "; .rcNormal.left; ", "; .rcNormal.top; ", "; .rcNormal.right; ", "; .rcNormal.bottom
                Debug.Print "ptMin = "; .ptMin.x; ", "; .ptMin.y
                Debug.Print "pName = "; .pName
                Debug.Print "cxIcon = "; .cxIcon
                Debug.Print "cyIcon = "; .cyIcon
                Debug.Print "wIconFormat = "; .wIconFormat
                Debug.Print "wReserved = "; .wReserved
                Debug.Print "cItems = "; .cItems
            End With
            Debug.Print "Name = '"; c_tGrpDef.Name; "'"

            '// read the old item table
            ReDim .Items(.grpdef.cItems - 1)
            Seek #iFile, Len(.grpdef) + 1
            For iIndex = 0 To UBound(.Items)
                Get #iFile, , .Items(iIndex)
            Next iIndex

            '// Read in the items.
            For iIndex = 0 To UBound(.Items)
                If Not ((.Items(iIndex) = 0) Or _
                       (.Items(iIndex) > .grpdef.cbGroup)) Or True Then
                    
                    Seek #iFile, .Items(iIndex) + 1
                    Get #iFile, , tGrpItm
                    
                    Set tClass = New CProgmanGroupItem
                    
                    With tGrpItm
                        tClass.PointX = .pt.x
                        tClass.PointY = .pt.y
                        tClass.Icon = .iIcon
                        tClass.cbHeader = .cbHeader
                        tClass.cbANDPlane = .cbANDPlane
                        tClass.cbXORPlane = .cbXORPlane
                        tClass.pHeader = .pHeader
                        tClass.pANDPlane = .pANDPlane
                        tClass.pXORPlane = .pXORPlane
                        
                        tClass.ItemName = ReadStringFromFile(iFile, 30, tGrpItm.pName + 1)
                        tClass.ItemCommand = ReadStringFromFile(iFile, 30, tGrpItm.pCommand + 1)
                        tClass.ItemIconPath = ReadStringFromFile(iFile, 30, tGrpItm.pIconPath + 1)
                    End With
                    
                    Call tClass.Dump
                    
                    c_tItems.Add tClass
                End If
            Next iIndex
        End With
    Close #1
    
    Exit Function
ErrHandler:
    ReadGroup = False
End Function

Private Sub Class_Initialize()
    Set c_tItems = New Collection
End Sub


Public Property Get Items() As Collection
    Set Items = c_tItems
End Property


