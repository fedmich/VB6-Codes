Function RemoveFilePath(path As String) As String
	If Right$(path, 1) = ":" Then RemoveFilePath = "": Exit Function
	If Right$(path, 1) = "\" Then RemoveFilePath = "": Exit Function
	
	Dim WhereSlash  As Integer
	WhereSlash = InStrRev(path, "\")
	If WhereSlash > 0 Then
		RemoveFilePath = Right$(path, Len(path) - WhereSlash)
	Else
		RemoveFilePath = path
	End If
End Function