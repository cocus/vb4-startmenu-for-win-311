Attribute VB_Name = "Module1"
Option Explicit






#If Win32 Then
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type


#Else
Private Type POINTAPI
    X As Integer
    Y As Integer
End Type

Private Type RECT
    left As Integer
    top As Integer
    right As Integer
    bottom As Integer
End Type

#End If



#If Win32 Then
#Else
Private Const GW_HWNDNEXT = 2
Private Const GCW_HICON As Long = (-14)
Private Const GWW_HINSTANCE As Long = (-6)
#End If


#If Win32 Then
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long

Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long


Public Declare Function DrawIcon Lib "user32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function CopyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
#Else


'// https://binaryworld.net/main/CodeDetail.aspx?CodeId=43
Declare Function VarPtr Lib "VB40016.DLL" (variable As Any) As Long


Private Declare Function lstrlen Lib "Kernel.dll" (ByVal lpString As Any) As Integer

Private Declare Sub GetCursorPos Lib "User.dll" (ByRef lpPoint As POINTAPI)
Private Declare Sub GetCursorPosPtr Lib "User.dll" Alias "GetCursorPos" (lpPoint As Long)

Private Declare Sub ClientToScreen Lib "User.dll" (ByVal hWnd As Integer, ByRef lpPoint As POINTAPI)
Private Declare Sub SetRect Lib "User.dll" (ByRef lpRect As RECT, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Private Declare Function WindowFromPoint Lib "User.dll" (ByVal Pnt As Any) As Integer
Private Declare Function GetParent Lib "User.dll" (ByVal hWnd As Integer) As Integer

Private Declare Function PtInRect Lib "User.dll" (ByRef lpRect As RECT, ByVal Pnt As Any) As Integer
Private Declare Sub GetWindowRect Lib "User.dll" (ByVal hWnd As Integer, ByRef lpRect As RECT)


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


Private Declare Function ExtractIcon Lib "shell.dll" (ByVal hisnt As Integer, ByVal lpszExeName As String, ByVal iIcon As Integer) As Integer
Private Declare Function CopyIcon Lib "User.dll" (ByVal hInst As Integer, ByVal hIcon As Integer) As Integer
Public Declare Function DrawIcon Lib "User.dll" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal hIcon As Integer) As Integer
Public Declare Function DestroyIcon Lib "User.dll" (ByVal hIcon As Integer) As Integer

#End If


#If Win32 Then
Public Function GetWindowUnderMouse() As Long
#Else
Public Function GetWindowUnderMouse() As Integer
#End If
#If Win32 Then
    Dim pt                  As POINTAPI

    Call GetCursorPos(pt)

    GetWindowUnderMouse = WindowFromPoint(pt.X, pt.Y)
#Else
    Dim pt                  As Long

    Call GetCursorPosPtr(pt)

    GetWindowUnderMouse = WindowFromPoint(pt)
#End If
End Function



'---------------------------------------------------------------------------------------
' Procedure : IsMouseInWindows
' Author    : Cocus
' Date      : 6/10/2018
' Purpose   : Check if the mouse is hovering a window or one of its childs.
'---------------------------------------------------------------------------------------
'
#If Win32 Then
Public Function IsMouseInWindows(ByVal hWnd As Long) As Boolean
    
    Dim hWindowUnderMouse       As Long
#Else
Public Function IsMouseInWindows(ByVal hWnd As Integer) As Boolean

    Dim hWindowUnderMouse       As Integer
#End If

    hWindowUnderMouse = GetWindowUnderMouse()

    If hWindowUnderMouse = hWnd Then
        IsMouseInWindows = True
    Else
        If GetParent(hWindowUnderMouse) = hWnd Then
            IsMouseInWindows = True
        End If
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetIconFromPath
' Author    : Cocus
' Date      : 6/10/2018
' Purpose   : Get an icon handle from the specified path and icon index.
'---------------------------------------------------------------------------------------
'
#If Win32 Then
Public Function GetIconFromPath(ByVal sPath As String, Optional ByVal iIndex As Integer) As Long
#Else
Public Function GetIconFromPath(ByVal sPath As String, Optional ByVal iIndex As Variant) As Integer
#End If
    Dim hIcon           As Integer

    If IsMissing(iIndex) Then
        iIndex = 0
    End If

    GetIconFromPath = ExtractIcon(App.hInstance, sPath, iIndex)
    
    If GetIconFromPath Then
        #If Win32 Then
            Dim hIconCopy       As Long
            hIconCopy = CopyIcon(GetIconFromPath)
        #Else
            Dim hIconCopy       As Integer
            hIconCopy = CopyIcon(App.hInstance, GetIconFromPath)
        #End If
        Call DestroyIcon(GetIconFromPath)
        GetIconFromPath = hIconCopy
    End If
End Function





