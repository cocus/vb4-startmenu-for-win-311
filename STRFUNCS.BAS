Attribute VB_Name = "Module2"
Option Explicit
 
Public Const vbTextCompare = 1
Public Const vbBinaryCompare = 0
 
Public Function Join(source() As String, Optional sDelim) As String
    Dim sOut As String, iC As Integer
    If IsMissing(sDelim) Then
        sDelim = " "
    End If
On Error GoTo errh:
    For iC = LBound(source) To UBound(source) - 1
        sOut = sOut & source(iC) & CStr(sDelim)
    Next
    sOut = sOut & source(iC)
    Join = sOut
    Exit Function
errh:
    Err.Raise Err.Number
End Function
 
Public Function Split(ByVal sIn As String, Optional sDelim, Optional nLimit, Optional bCompare) As Variant
    Dim sRead As String, sOut() As String, nC As Integer
    If sDelim = "" Then
        Split = sIn
    End If
    If IsMissing(nLimit) Then
        nLimit = -1
    End If
    If IsMissing(bCompare) Then
        bCompare = vbBinaryCompare
    End If
    sRead = ReadUntil(sIn, CStr(sDelim), bCompare)
    Do
        ReDim Preserve sOut(nC)
        sOut(nC) = sRead
        nC = nC + 1
        If nLimit <> -1 And nC >= nLimit Then Exit Do
        sRead = ReadUntil(sIn, CStr(sDelim))
    Loop While sRead <> ""
    ReDim Preserve sOut(nC)
    sOut(nC) = sIn
    Split = sOut
End Function
 
Public Function ReadUntil(ByRef sIn As String, sDelim As String, Optional bCompare) As String
    Dim nPos As String
    If IsMissing(bCompare) Then
        bCompare = vbBinaryCompare
    End If
    nPos = InStr(1, sIn, sDelim, bCompare)
    If nPos > 0 Then
        ReadUntil = left(sIn, nPos - 1)
        sIn = Mid(sIn, nPos + Len(sDelim))
    End If
End Function
 
Public Function StrReverse(ByVal sIn As String) As String
    Dim nC As Integer, sOut As String
    For nC = Len(sIn) To 1 Step -1
    sOut = sOut & Mid(sIn, nC, 1)
    Next
    StrReverse = sOut
End Function
 
Public Function InStrRev(ByVal sIn As String, sFind As String, Optional nStart, Optional bCompare) As Long
    Dim nPos As Long
    If IsMissing(nStart) Then
        nStart = 1
    End If
    If IsMissing(bCompare) Then
        bCompare = vbBinaryCompare
    End If
    sIn = StrReverse(sIn)
    sFind = StrReverse(sFind)
    nPos = InStr(CLng(nStart), sIn, sFind, bCompare)
    If nPos = 0 Then
        InStrRev = 0
    Else
        InStrRev = Len(sIn) - nPos - Len(sFind) + 2
    End If
End Function
 
Public Function Replace(sIn As String, sFind As String, sReplace As String, Optional nStart, Optional nCount, Optional bCompare) As String
    Dim nC As Long, nPos As Integer, sOut As String
    If IsMissing(nStart) Then
        nStart = 1
    End If
    If IsMissing(nCount) Then
        nCount = -1
    End If
    If IsMissing(bCompare) Then
        bCompare = vbBinaryCompare
    End If
    sOut = sIn
    nPos = InStr(CLng(nStart), sOut, sFind, bCompare)
    If nPos = 0 Then GoTo EndFn:
    Do
        nC = nC + 1
        sOut = left(sOut, nPos - 1) & sReplace & _
           Mid(sOut, nPos + Len(sFind))
        If CLng(nCount) <> -1 And nC >= CLng(nCount) Then Exit Do
        nPos = InStr(CLng(nStart), sOut, sFind, bCompare)
    Loop While nPos > 0
EndFn:
    Replace = sOut
End Function

