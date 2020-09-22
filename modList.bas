Attribute VB_Name = "modList"

Option Explicit
'
Public Function ListLen(List As String, Optional delim As String = ",") As Long
    Dim tLen As Long
    Dim tPos As Long
    Dim i As Long
    If List = "" Then
        ListLen = 0
        Exit Function
    End If
    tLen = 0
    tPos = 0
    For i = 1 To Len(List)
        tPos = InStr(tPos + Len(delim), List, delim, vbBinaryCompare)
        If tPos > 0 Then
            tLen = tLen + 1
        Else
            Exit For
        End If
    Next i
    
    ListLen = tLen + 1
End Function

Public Function ListGetAt(List As String, Position As Long, Optional delim As String = ",") As String
    Dim i As Long
    Dim tPos As Long, prePos As Long
    Dim tLen As Long
    ListGetAt = ""
    For i = 1 To ListLen(List, delim)
        If prePos = 0 Then
            prePos = 1
        Else
            prePos = tPos + Len(delim)
        End If
        tPos = InStr(prePos, List, delim, vbBinaryCompare)
        
        If tPos > 0 Then
            tLen = tLen + 1
            If tLen = Position Then
                ListGetAt = Mid(List, prePos, tPos - prePos)
                Exit Function
            End If
        End If
    Next i
    If Position = ListLen(List, delim) Then
        ListGetAt = Mid(List, prePos, Len(List) - prePos + 1)
    End If
End Function

Public Function ListAppend(List As String, Item As String, Optional delim As String = ",") As String
    If List = "" Then
        ListAppend = Item
    Else
        ListAppend = List & delim & Item
    End If
End Function

Public Function ListFind(List As String, Item As String, Optional delim As String = ",") As Long
    Dim lstLen As Long, Position As Long
    
    lstLen = ListLen(List, delim)
    
    For Position = 1 To lstLen
        If StrComp(ListGetAt(List, Position, delim), Item, vbBinaryCompare) = 0 Then
            ListFind = Position
            Exit Function
        End If
    Next Position
    ListFind = 0
End Function

Public Function ListFindNoCase(List As String, Item As String, Optional delim As String = ",") As Long
    Dim lstLen As Long, Position As Long
    
    lstLen = ListLen(List, delim)
    
    For Position = 1 To lstLen
        If StrComp(ListGetAt(List, Position, delim), Item, vbTextCompare) = 0 Then
            ListFindNoCase = Position
            Exit Function
        End If
    Next Position
    ListFindNoCase = 0
End Function

Public Function ListReplaceAt(List As String, Item As String, Position As Long, Optional delim As String = ",") As String
    Dim i As Long
    If Position = 0 Then
        ListReplaceAt = List
        Exit Function
    End If
    If List = "" Then
        ListReplaceAt = ""
        Exit Function
    End If
    For i = 1 To ListLen(List, delim)
        If i = Position Then
            ListReplaceAt = ListAppend(ListReplaceAt, Item, delim)
        Else
            ListReplaceAt = ListAppend(ListReplaceAt, ListGetAt(List, i, delim), delim)
        End If
    Next i
End Function

Public Function ListDeleteAt(List As String, Position As Long, Optional delim As String = ",") As String
    Dim i As Long
    
    If Position = 0 Or Position > ListLen(List, delim) Then
        ListDeleteAt = List
        Exit Function
    End If
    
    If List = "" Then
        ListDeleteAt = ""
        Exit Function
    End If
    
    For i = 1 To ListLen(List, delim)
        If i <> Position Then
            ListDeleteAt = ListAppend(ListDeleteAt, ListGetAt(List, i, delim), delim)
        End If
    Next i
    
End Function
