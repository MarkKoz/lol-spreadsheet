Attribute VB_Name = "Module1"
Const strNum As String = "1234567890"
Const allowedCharsNeg As String = "-1234567890"
Const allowedCharsK As String = ".k1234567890"

Public Sub NumericInput(str As String, chars As String)
' Only allows numerical values

    Dim inp As String: inp = Form.txt_Length_M.Text
    Dim subStr As String
    Dim index As Integer: index = Form.chars.SelStart
    
    If index <> 0 Then
        subStr = Mid(txt_Length_M.Text, index, 1)
        If InStr(1, allowedChars, subStr, vbTextCompare) = 0 Then
            inp = Replace(inp, subStr, "")
            txt_Length_M.Text = inp
            If index <> 0 Then
                txt_Length_M.SelStart = (index - 1)
            End If
        End If
    End If
End Sub
Sub showUserForm()
    UserForm1.Show
End Sub
