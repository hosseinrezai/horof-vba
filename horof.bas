Function horof(x As String, Optional f As Boolean = 0)
Dim text As String
Dim a As String
Dim b As String
Dim manfi As String
If (Left(x, 1) = "-") Then
        manfi = ChrW(1605) & ChrW(1606) & ChrW(1601) & ChrW(1740) & ChrW(32)
        x = Right(x, Len(x) - 1)
    Else
        mafni = ""
End If
If (Application.Max(InStr(x, "."), InStr(x, ChrW(1643))) = 0) Then
    a = x
    b = 0
Else
    a = Left(x, Application.Max(InStr(x, "."), InStr(x, ChrW(1643))))
    a = Replace(a, ".", "")
    a = Replace(a, ChrW(1643), "")
    b = Right(x, Len(x) - Application.Max(InStr(x, "."), InStr(x, ChrW(1643))))
End If
While (Left(a, 1) = "0")
    a = Right(a, Len(a) - 1)
Wend
Select Case Len(a)
    Case 1
        Select Case a
        Case 0
            'text = 'ChrW(1589) & ChrW(1601) & ChrW(1585)
        Case 1
            text = ChrW(1740) & ChrW(1705)
        Case 2
            text = ChrW(1583) & ChrW(1608)
        Case 3
            text = ChrW(1587) & ChrW(1607)
        Case 4
            text = ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585)
        Case 5
            text = ChrW(1662) & ChrW(1606) & ChrW(1580)
        Case 6
            text = ChrW(1588) & ChrW(1588)
        Case 7
            text = ChrW(1607) & ChrW(1601) & ChrW(1578)
        Case 8
            text = ChrW(1607) & ChrW(1588) & ChrW(1578)
        Case 9
            text = ChrW(1606) & ChrW(1607)
        Case Else
        End Select
    Case 2
        Select Case a
        Case 10
            text = ChrW(1583) & ChrW(1607)
        Case 11
            text = ChrW(1740) & ChrW(1575) & ChrW(1586) & ChrW(1583) & ChrW(1607)
        Case 12
            text = ChrW(1583) & ChrW(1608) & ChrW(1575) & ChrW(1586) & ChrW(1583) & ChrW(1607)
        Case 13
            text = ChrW(1587) & ChrW(1740) & ChrW(1586) & ChrW(1583) & ChrW(1607)
        Case 14
            text = ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1607)
        Case 15
            text = ChrW(1662) & ChrW(1575) & ChrW(1606) & ChrW(1586) & ChrW(1583) & ChrW(1607)
        Case 16
            text = ChrW(1588) & ChrW(1575) & ChrW(1606) & ChrW(1586) & ChrW(1583) & ChrW(1607)
        Case 17
            text = ChrW(1607) & ChrW(1601) & ChrW(1583) & ChrW(1607)
        Case 18
            text = ChrW(1607) & ChrW(1580) & ChrW(1583) & ChrW(1607)
        Case 19
            text = ChrW(1606) & ChrW(1608) & ChrW(1586) & ChrW(1583) & ChrW(1607)
        Case Else
                Select Case Left(a, 1)
                    Case 2
                        text = ChrW(1576) & ChrW(1740) & ChrW(1587) & ChrW(1578)
                    Case 3
                        text = ChrW(1587) & ChrW(1740)
                    Case 4
                        text = ChrW(1670) & ChrW(1607) & ChrW(1604)
                    Case 5
                        text = ChrW(1662) & ChrW(1606) & ChrW(1580) & ChrW(1575) & ChrW(1607)
                    Case 6
                        text = ChrW(1588) & ChrW(1589) & ChrW(1578)
                    Case 7
                        text = ChrW(1607) & ChrW(1601) & ChrW(1578) & ChrW(1575) & ChrW(1583)
                    Case 8
                        text = ChrW(1607) & ChrW(1588) & ChrW(1578) & ChrW(1575) & ChrW(1583)
                    Case 9
                        text = ChrW(1606) & ChrW(1608) & ChrW(1583)
                Case Else
                End Select
                If (Right(a, 1) <> "0") Then
                    If (Left(a, 1) <> "0") Then
                        text = text & ChrW(32) & ChrW(1608) & ChrW(32)
                    End If
                    text = text & horof(Right(a, 1), 1)
                End If
        End Select
    Case 3
        Select Case Left(a, 1)
        Case 1
            text = ChrW(1589) & ChrW(1583)
        Case 2
            text = ChrW(1583) & ChrW(1608) & ChrW(1740) & ChrW(1587) & ChrW(1578)
        Case 3
            text = ChrW(1587) & ChrW(1740) & ChrW(1589) & ChrW(1583)
        Case 4
            text = ChrW(1670) & ChrW(1607) & ChrW(1575) & ChrW(1585) & ChrW(1589) & ChrW(1583)
        Case 5
            text = ChrW(1662) & ChrW(1575) & ChrW(1606) & ChrW(1589) & ChrW(1583)
        Case 6
            text = ChrW(1588) & ChrW(1588) & ChrW(1589) & ChrW(1583)
        Case 7
            text = ChrW(1607) & ChrW(1601) & ChrW(1578) & ChrW(1589) & ChrW(1583)
        Case 8
            text = ChrW(1607) & ChrW(1588) & ChrW(1578) & ChrW(1589) & ChrW(1583)
        Case 9
            text = ChrW(1606) & ChrW(1607) & ChrW(1589) & ChrW(1583)
        Case Else
        End Select
        If (Right(a, 2) <> "00") Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32) & horof(Right(a, 2), 1)
        End If
    Case 4 To 6
        pow = 3
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 7 To 9
        pow = 6
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 10 To 12
        pow = 9
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 13 To 15
        pow = 12
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1576) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 16 To 18
        pow = 15
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1576) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 19 To 21
        pow = 18
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1578) & ChrW(1585) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 22 To 24
        pow = 21
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1578) & ChrW(1585) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 25 To 27
        pow = 24
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1705) & ChrW(1608) & ChrW(1570) & ChrW(1583) & ChrW(1585) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 28 To 30
        pow = 27
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1705) & ChrW(1575) & ChrW(1583) & ChrW(1585) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case 31 To 33
        pow = 30
        text = horof(Left(a, Len(a) - pow), 1) & ChrW(32) & ChrW(1705) & ChrW(1608) & ChrW(1740) & ChrW(1606) & ChrW(1578) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606)
        If (Right(a, pow) <> String(pow, "0")) Then
            text = text & ChrW(32) & ChrW(1608) & ChrW(32)
            text = text & horof(Right(a, pow), 1)
        End If
    Case Is > 33
        text = ""
    End Select
If (Application.Max(InStr(x, "."), InStr(x, ChrW(1643))) > 0) Then
    text = text & ChrW(32) & ChrW(1605) & ChrW(1605) & ChrW(1740) & ChrW(1586) & ChrW(32) & horof(b, 1)
    Select Case Len(b)
    Case 1
        text = text & ChrW(32) & ChrW(1583) & ChrW(1607) & ChrW(1605)
    Case 2
        text = text & ChrW(32) & ChrW(1589) & ChrW(1583) & ChrW(1605)
    Case 3
        text = text & ChrW(32) & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605)
    Case 4
        text = text & ChrW(32) & ChrW(1583) & ChrW(1607) & ChrW(8204) & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605)
    Case 5
        text = text & ChrW(32) & ChrW(1589) & ChrW(1583) & ChrW(8204) & ChrW(1607) & ChrW(1586) & ChrW(1575) & ChrW(1585) & ChrW(1605)
    Case 6
        text = text & ChrW(32) & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606) & ChrW(1608) & ChrW(1605)
    Case 7
        text = text & ChrW(32) & ChrW(1583) & ChrW(1607) & ChrW(8204) & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606) & ChrW(1608) & ChrW(1605)
    Case 8
        text = text & ChrW(32) & ChrW(1589) & ChrW(1583) & ChrW(8204) & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1608) & ChrW(1606) & ChrW(1608) & ChrW(1605)
    Case 9
        text = text & ChrW(32) & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1605)
    Case 10
        text = text & ChrW(32) & ChrW(1583) & ChrW(1607) & ChrW(8204) & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1605)
    Case 11
        text = text & ChrW(32) & ChrW(1589) & ChrW(1583) & ChrW(8204) & ChrW(1605) & ChrW(1740) & ChrW(1604) & ChrW(1740) & ChrW(1575) & ChrW(1585) & ChrW(1583) & ChrW(1605)
    Case 12
        text = text & ChrW(32)
    Case 13
        text = text & ChrW(32)
    Case 14
        text = text & ChrW(32)
    Case 15
        text = text & ChrW(32)
    Case 16
        text = text & ChrW(32)
    Case 17
        text = text & ChrW(32)
    Case 18
        text = text & ChrW(32)
    Case 19
        text = text & ChrW(32)
    Case 20
        text = text & ChrW(32)
    Case 21
        text = text & ChrW(32)
    Case 22
        text = text & ChrW(32)
    Case 23
        text = text & ChrW(32)
    Case 24
        text = text & ChrW(32)
    Case 25
        text = text & ChrW(32)
    Case 26
        text = text & ChrW(32)
    Case 27
        text = text & ChrW(32)
    Case 28
        text = text & ChrW(32)
    Case 29
        text = text & ChrW(32)
    Case 30
        text = text & ChrW(32)
    Case 31
        text = text & ChrW(32)
    Case 32
        text = text & ChrW(32)
    Case 33
        text = text & ChrW(32)
 End Select
End If
horof = manfi & text
End Function

