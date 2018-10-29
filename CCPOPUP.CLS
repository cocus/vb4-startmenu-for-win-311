VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCustomPopupMenu"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit


Private c_fPopup            As frmPopup

Private c_oItems            As Collection

Public Sub F_EVENT_Unload(Cancel As Integer)
    Debug.Print "[CCustomPopupMenu] PopupForm Unload!"
End Sub

Public Sub F_EVENT_ItemClick(ByVal iIndex As Integer)
    Debug.Print "[CCustomPopupMenu] ItemClick: "; iIndex; "!"
    On Local Error Resume Next
    Debug.Print "Going to run "; c_oItems(iIndex + 1).ItemCommand
    Shell c_oItems(iIndex + 1).ItemCommand
    c_fPopup.Hide
    Unload c_fPopup
    Set c_fPopup = Nothing
End Sub

Private Sub Class_Initialize()
    Set c_fPopup = New frmPopup
    Call c_fPopup.SetCustomPopupMenuClass(Me)
    Debug.Print "[CCustomPopupMenu] Initialized!"
End Sub

Private Sub Class_Terminate()
    Debug.Print "[CCustomPopupMenu] Terminated!"
    Set c_fPopup = Nothing
End Sub

Public Sub PopulateFromGroupItems(ByRef cItems As Collection)
    Dim iIndex          As Integer
    Dim oGroupItem      As Variant

    iIndex = 0

    Set c_oItems = cItems

    With c_fPopup
        Call .SetItemsCollection(cItems)
    
        For Each oGroupItem In cItems
            If (iIndex) Then
                Load .lblLabels(iIndex)
                With .lblLabels(iIndex)
                    .left = c_fPopup.lblLabels(iIndex - 1).left
                    .top = c_fPopup.lblLabels(0).top + iIndex * (c_fPopup.picIcons(0).Height + 2)
                End With

                Load .picIcons(iIndex)
                With .picIcons(iIndex)
                    .AutoRedraw = True
                    .Height = c_fPopup.picIcons(0).Height
                    .Width = c_fPopup.picIcons(0).Width
                    .top = c_fPopup.picIcons(0).top + iIndex * (c_fPopup.picIcons(0).Height + 2)
                End With
            End If

            With .lblLabels(iIndex)
                .Caption = oGroupItem.ItemName
                .Visible = True
            End With

            With .picIcons(iIndex)
                Dim hIcon           As Variant
                hIcon = GetIconFromPath(oGroupItem.ItemIconPath)
                If hIcon Then
                    c_fPopup.picIcons(iIndex).Cls
                    DrawIcon .hdc, 0, 0, hIcon
                    Call DestroyIcon(hIcon)
                End If
                .Visible = True
            End With

            iIndex = iIndex + 1
        Next
    End With
    
    Call c_fPopup.Show
End Sub

