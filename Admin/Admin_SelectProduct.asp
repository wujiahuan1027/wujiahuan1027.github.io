<%@language=vbscript codepage=936 %>
<%
Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<html>
<head>
<title>选择商品</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Admin_STYLE.CSS" rel="stylesheet" type="text/css">
<base target="_self">
</head>
<body>
<%
Dim TotalPut, ProductType, CurrentPage, MaxPerPage, KeyWord
Const ImgWidth = 130
Const ImgHeight = 90
If Request("page") <> "" Then
    CurrentPage = CInt(Request("page"))
Else
    CurrentPage = 1
End If
ProductType = PE_Clng(Trim(Request("ProductType")))
MaxPerPage = PE_Clng(Trim(Request("MaxPerPage")))
KeyWord = ReplaceBadChar(Trim(Request("KeyWord")))
If MaxPerPage <= 0 Then MaxPerPage = 20

Dim sqlChannel, rsChannel, ChannelUrl, UploadDir
sqlChannel = "select * from PE_Channel  where ChannelID=1000"
Set rsChannel = Conn.Execute(sqlChannel)
If Not (rsChannel.bof And rsChannel.EOF) Then
    If rsChannel("Disabled") = True Then
        Response.Write "<br><li>此频道已经被管理员禁用！</li>"
        Response.End
    End If
    ChannelUrl = strInstallDir & rsChannel("ChannelDir")
    UploadDir = rsChannel("UploadDir")
End If
rsChannel.Close
Set rsChannel = Nothing

Dim strFileName, strProductName
strFileName = "Admin_SelectProduct.asp"

Response.Write "<form method='post' name='myform' action=''>" & vbCrLf
%>
<table border='0' align='center' cellpadding='2' cellspacing='0' class='border'>
  <tr height='22' class='title'>
    <td>查找从属商品</td><td align=right><input name='KeyWord' type='text' size='20' value=>&nbsp;&nbsp;<input type='submit' value='查找'></td>
  </tr>
</table>
<%
Response.Write "<table cellpadding='0' cellspacing='5' border='0' align='center'><tr valign='top'>"
Dim sqlPresent, rsPresent, i, TitleStr, strPic, strLink, trs
sqlPresent = "select P.ProductID,P.ProductNum,P.ProductName,P.ProductType,P.Price,Price_Original,P.Price_Market,P.Price_Member,BeginDate,EndDate,P.UpdateTime,P.ProductThumb from PE_Product P"
sqlPresent = sqlPresent & " where P.Deleted=" & PE_False & " and P.EnableSale=" & PE_True
If KeyWord <> "" Then sqlPresent = sqlPresent & " and P.ProductName like '%" & KeyWord & "%'"
If ProductType = 4 Then
    sqlPresent = sqlPresent & " and P.ProductType=4 order by P.ProductID desc"
    strProductName = "促销礼品"
Else
    sqlPresent = sqlPresent & " and P.ProductType<>4 order by P.ProductID desc"
    strProductName = "商品"
End If
Set rsPresent = Server.CreateObject("ADODB.Recordset")
rsPresent.open sqlPresent, Conn, 1, 1
If rsPresent.bof And rsPresent.EOF Then
    TotalPut = 0
    Response.Write "<td align='center'><img class='pic5' src='" & strInstallDir & "images/nopic.gif' width='" & ImgWidth & "' height='" & ImgHeight & "' border='0'><br>没有任何" & strProductName & "</td>"
Else
    i = 0
    TotalPut = rsPresent.recordcount
    If TotalPut > 0 Then
        If CurrentPage < 1 Then
            CurrentPage = 1
        End If
        If (CurrentPage - 1) * MaxPerPage > TotalPut Then
            If (TotalPut Mod MaxPerPage) = 0 Then
                CurrentPage = TotalPut \ MaxPerPage
            Else
                CurrentPage = TotalPut \ MaxPerPage + 1
            End If
        End If
        If CurrentPage > 1 Then
            If (CurrentPage - 1) * MaxPerPage < TotalPut Then
                rsPresent.Move (CurrentPage - 1) * MaxPerPage
            Else
                CurrentPage = 1
            End If
        End If
    End If
    Do While Not rsPresent.EOF
        strLink = "<a href='#' onclick=""window.returnValue='" & rsPresent("ProductID") & "," & rsPresent("ProductName") & "," & rsPresent("ProductNum") & "';window.close();"">"
        Response.Write "<td align='center'>" & strLink & GetProductThumb(rsPresent("ProductThumb"), ImgWidth, ImgHeight) & "<br>" & rsPresent("ProductName") & "</a>"
        Response.Write "<div align='left'>原价：<font class='price'><STRIKE>￥" & rsPresent("Price_Original") & "</STRIKE></font>"
        Response.Write "<br>现价：<font class='price'>￥" & rsPresent("Price") & "</font>"
        Response.Write "</div></td>"
        rsPresent.movenext
        i = i + 1
        If i >= MaxPerPage Then Exit Do
        If ((i Mod 6 = 0) And (Not rsPresent.EOF)) Then Response.Write "</tr><tr valign='top'>"
    Loop
End If
Response.Write "</tr></table>" & vbCrLf
rsPresent.Close
Set rsPresent = Nothing
Response.Write ShowSourcePage(strFileName, TotalPut, MaxPerPage, CurrentPage, True, True, "件商品", True)

Function GetProductThumb(ProductThumb, iWidth, iHeight)
    Dim strProductThumb, FileType
    If ProductThumb = "" Then
        strProductThumb = strProductThumb & "<img src='" & strInstallDir & "images/nopic.gif' "
        If iWidth > 0 Then strProductThumb = strProductThumb & " width='" & iWidth & "'"
        If iHeight > 0 Then strProductThumb = strProductThumb & " height='" & iHeight & "'"
        strProductThumb = strProductThumb & " border='0'>"
    Else
        FileType = Right(LCase(ProductThumb), 3)
        If IsNumeric(Left(ProductThumb, 6)) Then
            ProductThumb = ChannelUrl & "/" & UploadDir & "/" & ProductThumb
        End If
        If FileType = "swf" Then
            strProductThumb = strProductThumb & "<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width='" & iWidth & "'"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height='" & iHeight & "'"
            strProductThumb = strProductThumb & "><param name='movie' value='" & ProductThumb & "'><param name='quality' value='high'><embed src='" & ProductThumb & "' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width='" & iWidth & "'"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height='" & iHeight & "'"
            strProductThumb = strProductThumb & "></embed></object>"
        ElseIf FileType = "jpg" Or FileType = "bmp" Or FileType = "png" Or FileType = "gif" Then
            strProductThumb = strProductThumb & "<img class='pic3' src='" & ProductThumb & "' "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width='" & iWidth & "'"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height='" & iHeight & "'"
            strProductThumb = strProductThumb & " border='0'>"
        Else
            strProductThumb = strProductThumb & "<img class='pic3' src='" & strInstallDir & "images/nopic.gif' "
            If iWidth > 0 Then strProductThumb = strProductThumb & " width='" & iWidth & "'"
            If iHeight > 0 Then strProductThumb = strProductThumb & " height='" & iHeight & "'"
            strProductThumb = strProductThumb & " border='0'>"
        End If
    End If
    GetProductThumb = strProductThumb
End Function

Public Function ShowSourcePage(sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit, ShowMaxPerPage)
    Dim TotalPage, strTemp, strUrl, i

    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowSourcePage = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
        
    strTemp = "<div class=""show_page"">"
    If ShowTotal = True Then
        strTemp = strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
    End If
    If ShowMaxPerPage = True Then
        strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"
    Else
        strUrl = JoinChar(sfilename)
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "首页 上一页&nbsp;"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>上一页</a>&nbsp;"
    End If

    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "下一页 尾页"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>下一页</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>尾页</a>"
    End If
    strTemp = strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>页 "
    If ShowMaxPerPage = True Then
        strTemp = strTemp & "&nbsp;<input type='text' name='MaxPerPage' size='3' maxlength='4' value='" & MaxPerPage & "' onKeyPress=""if (event.keyCode==13) this.location='" & JoinChar(sfilename) & "page=" & CurrentPage & "&MaxPerPage=" & "'+this.value;"">" & strUnit & "/页"
    Else
        strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & strUnit & "/页"
    End If
    If ShowAllPages = True Then
        strTemp = strTemp & "&nbsp;&nbsp;转到第<input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) this.location='" & strUrl & "page=" & "'+this.value;"" onmousewheel=""if ((parseInt(this.value) + parseInt(event.wheelDelta/120))>0&&(parseInt(this.value) + parseInt(event.wheelDelta/120))<=" & TotalPage & ") this.value=parseInt(this.value) + parseInt(event.wheelDelta/120);"">页"
    End If
    strTemp = strTemp & "</div>"
    ShowSourcePage = strTemp
End Function
%>
</form>
</body>
</html>

