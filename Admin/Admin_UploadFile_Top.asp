<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit
Response.buffer = True
Const PurviewLevel = 0
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
%>

<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
Dim tUploadDir, ChannelID, ChannelName
'打开数据库
Call OpenConn

'获取频道相关数据
tUploadDir = Trim(Request("UploadDir"))
ChannelID = PE_CLng(Trim(Request("ChannelID")))
If ChannelID > 0 Then
    Dim sqlChannel, rsChannel
    sqlChannel = "select * from PE_Channel where ChannelID=" & ChannelID & " order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    If rsChannel.BOF And rsChannel.EOF Then
        Response.Write "<li>找不到指定的频道！"
    Else
        If rsChannel("Disabled") = True Then
            Response.Write "此频道已经被禁用！"
        End If
        ChannelName = rsChannel("ChannelName")
    End If
    rsChannel.Close
    Set rsChannel = Nothing
Else
    If tUploadDir = "UploadAdPic" Then
        ChannelName = "网站广告"
    End If
End If
Call CloseConn

Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<title>顶部管理导航菜单</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "" & vbCrLf
Response.Write "<body background='Images/admin_top_bg.gif' leftmargin='0' topmargin='0'>" & vbCrLf
Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
Response.Write "  <tr class='topbg'> " & vbCrLf
Response.Write "    <td height='22' colspan='10'><table width='100%'><tr class='topbg'><td align='center'><b>" & ChannelName & "管理----上传文件管理" & "</b></td><td width='60' align='right'><a href='http://go.powereasy.net/go.aspx?UrlID=10011' target='_blank'><img src='images/help.gif' border='0'></a></td></tr></table></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr class='tdbg'>" & vbCrLf
Response.Write "    <td height='30'><strong>管理说明：</strong>&nbsp;请打开左边上传目录导航，以便您更快捷的管理上传目录中的文件。</td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf
%>
