Public Function OpenAccessMDB(file$) As Boolean
    If Not FileExists(file) Then
        MsgBox "DB file: " & file & vbNewLine & " is missing", vbExclamation, "Database not found"
        Exit Function
    End If
    
    Dim cmd$
    cmd = Files_3( _
        "c:\Program Files\Microsoft Office\OFFICE11\MSACCESS.EXE", _
        "C:\Program Files (x86)\Microsoft Office\OFFICE11\MSACCESS.EXE", _
        "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.EXE" _
        )
    If cmd = "" Then
        MsgBox "Access is not installed on this computer", vbExclamation
        Exit Function
    End If
    cmd = cmd & " """ & file & """"
    
    Shell cmd, vbNormalFocus
    OpenAccessMDB = True
End Function