Attribute VB_Name = "Module1"
Public Function FileExists(FileName As String) As Boolean
FileExists = (Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
End Function

Public Function encrypt(Message As String) As String
    Randomize
    On Error GoTo errorcheck
    Dim tempmessage As String
    Dim basea As Integer
    Dim tempbasea As String
    Message = Reverse_String(Message)
    tempmessage = CStr(Message)
    basea = Int(Rnd * 75) + 25
    If basea < 0 Then
        tempbasea = CStr(basea)
        tempbasea = Right(tempbasea, Len(tempbasea) - 1)
        basea = CInt(tempbasea)
    End If
    basea = basea / 2
    encrypt = CStr(basea) + ";"
    For X = 1 To Len(tempmessage)
        encrypt = encrypt + CStr(Asc(Left(tempmessage, X)) - basea) + ";"
        basea = basea + 1
        tempmessage = Right(tempmessage, Len(tempmessage) - 1)
    Next X
errorcheck:
End Function

Public Function decrypt(code As String) As String
    On Error GoTo errorcheck
    Dim basea As Integer
    Dim tempcode As String
    Do Until Left(code, 1) = ";"
        tempcode = tempcode + Left(code, 1)
        code = Right(code, Len(code) - 1)
    Loop
    basea = CInt(tempcode)
    tempcode = ""
    code = Right(code, Len(code) - 1)
    Do Until code = ""
        Do Until Left(code, 1) = ";"
            tempcode = tempcode + Left(code, 1)
            code = Right(code, Len(code) - 1)
        Loop
        decrypt = decrypt + Chr(CLng(tempcode) + basea)
        code = Right(code, Len(code) - 1)
        tempcode = ""
        basea = basea + 1
    Loop
    decrypt = Reverse_String(decrypt)
errorcheck:
End Function

Public Function Reverse_String(Message As String) As String
    For X = 1 To Len(Message)
        Reverse_String = Reverse_String + Left(Right(Message, X), 1)
    Next X
End Function

