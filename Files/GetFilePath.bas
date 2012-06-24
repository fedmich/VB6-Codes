Function GetFilePath(path As String) As String
	If right(path, 1) = "\" Then GetFilePath = path: Exit Function
	If right(path, 1) = ":" Then GetFilePath = path: Exit Function
	
	Dim WhereSlash  As Integer
	WhereSlash = InStrRev(path, "\")
	If WhereSlash > 0 Then
		GetFilePath = left(path, WhereSlash)
	End If
End Function