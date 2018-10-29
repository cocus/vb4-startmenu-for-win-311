VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   14250
   ClientLeft      =   12165
   ClientTop       =   630
   ClientWidth     =   11520
   Height          =   14655
   Left            =   12105
   LinkTopic       =   "Form1"
   ScaleHeight     =   14250
   ScaleWidth      =   11520
   Top             =   285
   Width           =   11640
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   13815
      Left            =   3240
      ScaleHeight     =   13755
      ScaleWidth      =   7395
      TabIndex        =   4
      Top             =   240
      Width           =   7455
      Begin VB.PictureBox Picture3 
         Height          =   13695
         Left            =   0
         ScaleHeight     =   13635
         ScaleWidth      =   7035
         TabIndex        =   6
         Top             =   0
         Width           =   7095
         Begin VB.PictureBox picIcons 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Index           =   0
            Left            =   0
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "AppName"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   8
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
      Begin VB.VScrollBar VScroll1 
         Height          =   2415
         Left            =   7080
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   600
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit


#If Win32 Then
#Else
Private Const GW_HWNDNEXT = 2
Private Const GCW_HICON As Long = (-14)
Private Const GWW_HINSTANCE As Long = (-6)
#End If


#If Win32 Then
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Declare Function DrawIcon Lib "user32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function CopyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
#Else

Private Declare Function FindWindow Lib "User.dll" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Integer
Private Declare Function GetWindow Lib "User.dll" (ByVal hWnd As Integer, ByVal wCmd As Integer) As Integer

Private Declare Function EnumWindows Lib "User.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Integer

Private Declare Function GetWindowText Lib "User.dll" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer) As Integer
Private Declare Function GetWindowTextLength Lib "User.dll" (ByVal hWnd As Integer) As Integer

Private Declare Function IsIconic Lib "User.dll" (ByVal hWnd As Integer) As Integer
Private Declare Function IsWindowVisible Lib "User.dll" (ByVal hWnd As Integer) As Integer

Private Declare Function GetModuleFileName Lib "Kernel.dll" (ByVal hModule As Integer, ByVal lpFilename As String, ByVal nSize As Integer) As Integer

Private Declare Function GetWindowWord Lib "User.dll" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Private Declare Function GetClassWord Lib "User.dll" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer

Private Declare Function lstrlen Lib "Kernel.dll" (ByVal lpString As Any) As Integer

Private Declare Function ExtractIcon Lib "shell.dll" (ByVal hisnt As Integer, ByVal lpszExeName As String, ByVal iIcon As Integer) As Integer
Private Declare Function CopyIcon Lib "User.dll" (ByVal hInst As Integer, ByVal hIcon As Integer) As Integer
Private Declare Function DrawIcon Lib "User.dll" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer, ByVal hIcon As Integer) As Integer
Private Declare Function DestroyIcon Lib "User.dll" (ByVal hIcon As Integer) As Integer

#End If

Private f_iHoveringItem                 As Integer
Private f_oProgmanGroup                 As New CProgmanGroup

Private Sub Command2_Click()
    Dim j As CCustomPopupMenu
    Set j = New CCustomPopupMenu
    j.PopulateFromGroupItems f_oProgmanGroup.Items
    Exit Sub
    Dim iIndex          As Integer
    Dim oGroupItem      As Variant

    iIndex = 0

    With f_oProgmanGroup
        For Each oGroupItem In .Items
            If (iIndex) Then
                Load lblLabels(iIndex)
                With lblLabels(iIndex)
                    .left = lblLabels(iIndex - 1).left
                    .top = lblLabels(0).top + iIndex * (picIcons(0).Height + 30)
                End With

                Load picIcons(iIndex)
                With picIcons(iIndex)
                    .AutoRedraw = True
                    .Height = picIcons(0).Height
                    .Width = picIcons(0).Width
                    .top = picIcons(0).top + iIndex * (picIcons(0).Height + 30)
                End With
            End If

            With lblLabels(iIndex)
                .Caption = oGroupItem.ItemName
                .Visible = True
            End With

            With picIcons(iIndex)
                Dim hIcon           As Variant
                hIcon = GetIconFromPath(oGroupItem.ItemIconPath)
                If hIcon Then
                    picIcons(iIndex).Cls
                    DrawIcon .hdc, 0, 0, hIcon
                    Call DestroyIcon(hIcon)
                End If
                .Visible = True
            End With

            iIndex = iIndex + 1
        Next
    End With
End Sub

Private Sub Command3_Click()
    Dim f As Class1
    Set f = New Class1
    
End Sub

Private Sub Form_Load()
    With f_oProgmanGroup
        Call .ReadGroup("ACCESORI.GRP")
    End With
    
    f_iHoveringItem = -1
End Sub

Private Sub lblLabels_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetHover(Index)
End Sub

Private Sub picIcons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetHover(Index)
End Sub

Private Sub SetHover(ByVal iNewItem As Integer)
    Dim hIcon           As Variant

    If Not ((f_iHoveringItem = -1) Or _
            (f_iHoveringItem = iNewItem)) Then
        With lblLabels(f_iHoveringItem)
            .BackColor = vbButtonFace
            .ForeColor = vbButtonText
        End With
        
        With picIcons(f_iHoveringItem)
            .BackColor = vbButtonFace
            .ForeColor = vbButtonText
            
            hIcon = GetIconFromPath(f_oProgmanGroup.Items(f_iHoveringItem + 1).ItemIconPath)
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

            hIcon = GetIconFromPath(f_oProgmanGroup.Items(iNewItem + 1).ItemIconPath)
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
    End If
    
    f_iHoveringItem = iNewItem
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If f_iHoveringItem > 0 Then
        Call SetHover(-1)
    End If
End Sub

Private Sub Command1_Click()
    'Set the form's graphics mode to persistent
    Me.AutoRedraw = True
    
    Dim w As Integer
    w = FindWindow(0, 0)
    Do
        Debug.Print "Resultado: "; w
        If IsWindowVisible(w) Then
            List1.AddItem GetWindowTextVB(w)
        End If
        w = GetWindow(w, GW_HWNDNEXT)

       
        If w = 0 Then
            Exit Do
        End If
    Loop
End Sub


Private Function GetWindowTextVB(ByVal hWnd As Integer) As String
    Dim sSave As String, Ret As Integer
    Ret = GetWindowTextLength(hWnd)
    Debug.Print "Ret="; Ret
    sSave = Space(Ret)
    GetWindowText hWnd, sSave, Ret + 1
    GetWindowTextVB = CStr(hWnd) + " " + sSave
End Function

Private Function TrimStringToZero(ByVal sInputString As String) As String
    Dim iPos        As Integer
    iPos = lstrlen(sInputString)
    TrimStringToZero = left$(sInputString, iPos)
End Function

Private Function GetIconFromPath(ByVal sPath As String, Optional ByVal iIndex As Variant) As Integer
    Dim hIcon           As Integer

    If IsMissing(iIndex) Then
        iIndex = 0
    End If

    GetIconFromPath = ExtractIcon(App.hInstance, sPath, iIndex)
    
    If GetIconFromPath Then
        Dim hIconCopy As Integer
        
        hIconCopy = CopyIcon(App.hInstance, GetIconFromPath)
        Call DestroyIcon(GetIconFromPath)
        GetIconFromPath = hIconCopy
    End If
End Function


Private Sub List1_DblClick()
    Dim s As String
    
    s = List1.Text
    Debug.Print s
    'foreach
    
    Dim Items As Variant
    Items = Split(s, " ")
    If UBound(Items) > 0 Then
        Debug.Print "hwnd ="; Items(0)
        
        Picture1.Cls
        Dim hIcon As Integer
        hIcon = GetClassWord(Items(0), GCW_HICON)
        If hIcon = 0 Then
            
            Dim sPath As String
            sPath = Space$(230)
            Dim iLength As Integer
            iLength = GetModuleFileName(GetWindowWord(Items(0), GWW_HINSTANCE), sPath, 230)
            
            If (iLength) Then
                sPath = left$(sPath, iLength)
                Debug.Print "ModuleFileName: "; sPath
                hIcon = ExtractIcon(App.hInstance, sPath, 0)
                
                If hIcon Then
                    Dim hIconCopy As Integer
                    
                    hIconCopy = CopyIcon(App.hInstance, hIcon)
                    Call DestroyIcon(hIcon)
                    hIcon = hIconCopy
                End If
            End If
        End If
        
        
        If hIcon > 1 Then
            DrawIcon Picture1.hdc, 0, 0, hIcon
            Call DestroyIcon(hIcon)
        End If
    End If
End Sub



