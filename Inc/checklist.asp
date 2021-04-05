<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!-- #include File="../conn.asp" -->
<!-- #include File="function.asp" -->
<%
Response.ContentType = "text/Html; charset=gb2312"
Response.Expires = -1
Dim Received, rootNode, itext, ilnum, ichannelid, itype, iname
Set Received = CreateObject("Microsoft.XMLDOM")
Received.async = False
Received.Load Request

Set rootNode = Received.getElementsByTagName("root")
If rootNode.length > 0 Then
    itext = rootNode(0).selectSingleNode("text").Text
    If itext <> "" Then
        itext = ReplaceBadChar(itext)
    Else
        Set Received = Nothing
        Response.End
    End If

    iname = rootNode(0).selectSingleNode("inputname").Text

    If itext = "ChongFuUserCheck" Then
        If iname <> "" Then
            iname = ReplaceBadChar(iname)
            Call usercheck()
        Else
            Response.write "0"
        End If
    Else
        ilnum = rootNode(0).selectSingleNode("lnum").Text
        ichannelid = rootNode(0).selectSingleNode("channelid").Text
        itype = rootNode(0).selectSingleNode("type").Text
        If ilnum = "" Or ilnum < 1 Then
            ilnum = 10
        Else
            ilnum = CLng(ilnum)
        End If
        If ichannelid = "" Or ichannelid < 1 Then
            ichannelid = 0
        Else
            ichannelid = CLng(ichannelid)
        End If
        Call outitem()
    End If
End If
Set Received = Nothing

Sub outitem()
    Dim rtext, qsql
    Select Case itype
    Case "satitle"
        qsql = "select top " & ilnum & " Title,UpdateTime from PE_Article where Title like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " and Deleted=False order by UpdateTime desc"
    Case "satitle2"
        qsql = "select top " & ilnum & " PhotoName,UpdateTime from PE_Photo where PhotoName like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " and Deleted=False order by UpdateTime desc"
    Case "satitle3"
        qsql = "select top " & ilnum & " SoftName,UpdateTime from PE_Soft where SoftName like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " and Deleted=False order by UpdateTime desc"
    Case "satitle4"
        qsql = "select top " & ilnum & " ProductName,UpdateTime from PE_Product where ProductName like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " and Deleted=False order by UpdateTime desc"
    Case "skey"
        qsql = "select top " & ilnum & " KeyText,LastUseTime from PE_NewKeys where KeyText like '" & itext & "%'"
        If ichannelid > 0 Then qsql = qsql & " and ChannelID=" & ichannelid
        qsql = qsql & " order by LastUseTime desc"
    Case "sauthor", "sauthor1"
        qsql = "select top " & ilnum & " AuthorName,LastUseTime from PE_Author where AuthorName like '" & itext & "%' and ChannelID=0"
        If ichannelid > 0 Then qsql = qsql & " or ChannelID=" & ichannelid
        qsql = qsql & " and Passed=True order by LastUseTime desc"
    Case "scopyfrom", "scopyfrom1"
        qsql = "select top " & ilnum & " SourceName,LastUseTime from PE_CopyFrom where SourceName like '" & itext & "%' and ChannelID=0"
        If ichannelid > 0 Then qsql = qsql & " or ChannelID=" & ichannelid
        qsql = qsql & " and Passed=True order by LastUseTime desc"
    End Select
    If qsql <> "" Then
        Set rtext = Conn.Execute(qsql)
        Do While Not rtext.EOF
            Response.write "<li style=""cursor:hand;"" onclick=""addinput('" & iname & "','" & rtext(0) & "');"">" & rtext(0) & "</li>"
            rtext.movenext
        Loop
        Set rtext = Nothing
    End If
End Sub

Sub usercheck()
    Dim rtext,UserNameLimit,UserNameMax,UserName_RegDisabled
    If InStr(iname, "=") > 0 Or InStr(iname, "%") > 0 Or InStr(iname, Chr(32)) > 0 Or InStr(iname, "?") > 0 Or InStr(iname, "&") > 0 Or InStr(iname, ";") > 0 Or InStr(iname, ",") > 0 Or InStr(iname, "'") > 0 Or InStr(iname, ",") > 0 Or InStr(iname, Chr(34)) > 0 Or InStr(iname, Chr(9)) > 0 Or InStr(iname, "") > 0 Or InStr(iname, "$") > 0 Or InStr(iname, "*") Or InStr(iname, "|") Or InStr(iname, """") > 0 Then
        Response.write "2" '含有非法字符
    Else
        Set rtext = Conn.Execute("select top 1 UserNameLimit,UserNameMax,UserName_RegDisabled from PE_Config")
        If Not (rtext.bof And rtext.EOF) Then
            UserNameLimit = rtext("UserNameLimit")
            UserNameMax = rtext("UserNameMax")
            UserName_RegDisabled = rtext("UserName_RegDisabled")
        Else
            UserNameLimit = 4
            UserNameMax = 20
        End If
        rtext.close

        If strLength(iname) > 20 Or strLength(iname) < UserNameLimit Then
            Response.write "3" '长度不合
        Else
            If FoundInArr(UserName_RegDisabled, iname, "|") = True Then
                Response.write "4" '禁止注册
            Else
                Set rtext = Conn.Execute("select top 1 UserName from PE_User where UserName='" & iname & "'")
                If rtext.bof and rtext.EOF Then
                    Response.write "0"
                Else
                    Response.write "1" '重复
                End If
                Set rtext = Nothing
            End If
        End If
    End If
End Sub
%>