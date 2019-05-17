Sub DateToQtr()


Dim testStr, year, monthRaw, month, qtr As String
Dim monthInt As Integer

testStr = "2016/02/24"

year = Left(testStr, InStr(testStr, "/") - 1)
monthRaw = Left(testStr, InStr(testStr, "/") + 2)
month = Right(monthRaw, Len(monthRaw) - 5)

monthInt = CInt(month)

If 1 <= monthInt And monthInt <= 3 Then
qtr = 1
End If
If 4 <= monthInt And monthInt <= 6 Then
qtr = 2
End If
If 7 <= monthInt And monthInt <= 9 Then
qtr = 3
End If
If 10 <= monthInt And monthInt <= 12 Then
qtr = 4
End If


MsgBox year
MsgBox month
MsgBox qtr

End Sub
