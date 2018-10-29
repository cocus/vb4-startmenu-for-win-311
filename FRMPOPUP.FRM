VERSION 4.00
Begin VB.Form frmPopup 
   Caption         =   "Form2"
   ClientHeight    =   13170
   ClientLeft      =   1020
   ClientTop       =   750
   ClientWidth     =   6690
   Height          =   13575
   Left            =   960
   LinkTopic       =   "Form2"
   ScaleHeight     =   13170
   ScaleWidth      =   6690
   Top             =   405
   Width           =   6810
   Begin VB.Timer tmrCheckMouse 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   3600
   End
   Begin VB.PictureBox picIcons 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "AppName"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape shpSelectionBack 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private f_cPopupMenu                    As CCustomPopupMenu

Private f_cItems                        As Collection

Private f_iHoveringItem                 As Integer

Public Sub SetItemsCollection(ByRef cItems As Collection)
    Set f_cItems = cItems
End Sub

Public Sub SetCustomPopupMenuClass(ByRef cPopup As CCustomPopupMenu)
    Set f_cPopupMenu = cPopup
End Sub

Private Sub Form_Load()
    f_iHoveringItem = -1
End Sub

Private Sub Form_Resize()
    shpSelectionBack.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (f_cPopupMenu Is Nothing) Then
        Call f_cPopupMenu.F_EVENT_Unload(Cancel)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iIndex              As Integer

    For iIndex = 0 To picIcons.UBound
        With picIcons(iIndex)
            If ((.top <= Y) And _
                ((.top + .Height) >= Y)) Then
                Call SetHover(iIndex)
                Exit Sub
            End If
        End With
    Next iIndex

    Call SetHover(-1)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iIndex              As Integer

    For iIndex = 0 To picIcons.UBound
        With picIcons(iIndex)
            If ((.top <= Y) And _
                ((.top + .Height) >= Y)) Then
                Call MouseUp(iIndex)
                Exit Sub
            End If
        End With
    Next iIndex
End Sub

Private Sub lblLabels_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetHover(Index)
End Sub

Private Sub lblLabels_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MouseUp(Index)
End Sub

Private Sub picIcons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetHover(Index)
End Sub

Private Sub picIcons_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call MouseUp(Index)
End Sub

Private Sub SetHover(ByVal iNewItem As Integer)
    Dim hIcon           As Variant

    If f_cItems Is Nothing Then
        Exit Sub
    End If

    If Not ((f_iHoveringItem = -1) Or _
            (f_iHoveringItem = iNewItem)) Then
        With lblLabels(f_iHoveringItem)
            .BackColor = vbButtonFace
            .ForeColor = vbButtonText
        End With
        
        With picIcons(f_iHoveringItem)
            .BackColor = vbButtonFace
            .ForeColor = vbButtonText
            
            hIcon = GetIconFromPath(f_cItems(f_iHoveringItem + 1).ItemIconPath)
            If hIcon Then
                picIcons(f_iHoveringItem).Cls
                DrawIcon .hdc, 0, 0, hIcon
                Call DestroyIcon(hIcon)
            End If
        End With
        
        shpSelectionBack.Visible = False
    End If
    
    If Not ((iNewItem = -1) Or _
            (f_iHoveringItem = iNewItem)) Then
        With lblLabels(iNewItem)
            .BackColor = vbHighlight
            .ForeColor = vbHighlightText
        End With
        With picIcons(iNewItem)
            .BackColor = vbHighlight
            .ForeColor = vbHighlightText

            hIcon = GetIconFromPath(f_cItems(iNewItem + 1).ItemIconPath)
            If hIcon Then
                picIcons(iNewItem).Cls
                DrawIcon .hdc, 0, 0, hIcon
                Call DestroyIcon(hIcon)
            End If
        End With
        
        With shpSelectionBack
            .ZOrder 1
            .top = picIcons(iNewItem).top
            .Visible = True
        End With
        
        tmrCheckMouse.Enabled = True
    End If
    
    If iNewItem = -1 Then
        tmrCheckMouse.Enabled = False
    End If
    
    f_iHoveringItem = iNewItem
End Sub

Private Sub MouseUp(ByVal iIndex As Integer)
    If Not (f_cPopupMenu Is Nothing) Then
        Call f_cPopupMenu.F_EVENT_ItemClick(iIndex)
    End If
End Sub

Private Sub tmrCheckMouse_Timer()
    If Not (IsMouseInWindows(Me.hwnd)) Then
        Call SetHover(-1)
    End If
End Sub


