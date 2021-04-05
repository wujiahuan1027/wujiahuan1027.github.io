<!--#include file="User_CommonCode.asp"-->
<!--#include file="../inc/function.asp"-->
<%
If CheckUserLogined() = False Then
    Call CloseConn
    Response.Redirect "User_Login.asp"
End If
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteName & " >> 会员中心"%></title>
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="top.asp"-->
<table width="756" border="0" align="center" cellpadding="0" cellspacing="0" class="user_border">
  <tr>
    <td valign="top">
      <table width="100%" border="0" cellpadding="5" cellspacing="0" class="user_box">
        <tr>
          <td class="user_righttitle"><img src="Images/point2.gif" align="absmiddle"><%Response.Write "您现在的位置：<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>会员中心</a> >> 我发表的评论"%></td>
        </tr>
        <tr>
          <td height="200" valign='top'>
            <% Call PE_Execute("PE_CMS6", "User_Comment") %>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>
<% Call CloseConn %>
