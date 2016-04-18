<%
'-------------------------------------------------------------
' Written by:  Clint LaFever
' Email:       lafeverc@hotmail.com
' Web:         http://vbasic.iscool.net
'-------------------------------------------------------------

Function EncryptText(strText,ByVal strPwd)
    Dim i , c 
    Dim strBuff
    If strPwd <> "" And strText <> "" Then
        strPwd = UCase(strPwd)
        If Len(strPwd) Then
            For i = 1 To Len(strText)
                c = Asc(Mid(strText, i, 1))
                c = c + Asc(Mid(strPwd, (i Mod Len(strPwd)) + 1, 1))
                strBuff = strBuff & Chr(c And &HFF)
            Next
        Else
            strBuff = strText
        End If
        EncryptText = strBuff
    Else
        EncryptText = ""
    End If
End Function

Function DecryptText(strText,ByVal strPwd)
    Dim i , c 
    Dim strBuff 
    If strPwd <> "" And strText <> "" Then
        strPwd = UCase(strPwd)
        If Len(strPwd) Then
            For i = 1 To Len(strText)
                c = Asc(Mid(strText, i, 1))
                c = c - Asc(Mid(strPwd, (i Mod Len(strPwd)) + 1, 1))
                strBuff = strBuff & Chr(c And &HFF)
            Next
        Else
            strBuff = strText
        End If
        DecryptText = strBuff
    Else
        DecryptText = ""
    End If
End Function
%>