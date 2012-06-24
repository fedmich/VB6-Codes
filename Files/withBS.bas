Function withBS(path As String) As String
    If Right$(path, 1) <> "\" Then
        withBS = path & "\"
    Else
        withBS = path
    End If
End Function