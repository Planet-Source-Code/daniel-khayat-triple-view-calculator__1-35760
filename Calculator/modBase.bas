Attribute VB_Name = "modBase"

Public Sub DecToBin(n As Long)
On Error GoTo errorHandler
     If n = 0 Then
        Exit Sub
    Else
        DecToBin (Int(n / 2))
        Base.Text2.Text = Base.Text2.Text & (n Mod 2)
    End If
    ClearDisplay = True
    mustClear = True
    Exit Sub
errorHandler:
  Text1.Text = "ERROR"
 MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error"  'the vbCrlf constant inserts a line break between the literal string and the error's description
 ClearDisplay = True
End Sub

Public Function BinToDec(n As String) As Long
    
    For i = Len(n) - 1 To 0 Step -1
        BinToDec = BinToDec + (2 ^ (Len(n) - i - 1)) * Val(Mid$(n, i + 1, 1))
    Next i

End Function


Public Function OctToDec(n As String) As Long
    
    For i = Len(n) - 1 To 0 Step -1
        OctToDec = OctToDec + (8 ^ (Len(n) - i - 1)) * Val(Mid$(n, i + 1, 1))
    Next i

End Function

Public Function HexToDec(n As String) As Long
    
    Dim value As Long
    
    For i = Len(n) - 1 To 0 Step -1
        If IsNumeric(Mid$(n, i + 1, 1)) Then
            value = Val((Mid$(n, i + 1, 1)))
        Else
            value = Asc(Mid$(n, i + 1, 1)) - 55 'ascii of A =65 and A in Hex = 10 then we have to make 65-55=10
        End If
        HexToDec = HexToDec + (16 ^ (Len(n) - i - 1)) * value
    Next i

End Function



