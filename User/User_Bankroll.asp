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
          <td class="user_righttitle"><img src="Images/point2.gif" align="absmiddle">
            <%
            Response.Write "您现在的位置：<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>会员中心</a> >> 资金明细查询 >> "
            Dim ShowType
            ShowType = PE_CLng(Trim(Request("ShowType")))
            Select Case ShowType
            Case 0
                Response.Write "所有明细记录"
            Case 1
                Response.Write "所有收入记录"
            Case 2
                Response.Write "所有支出记录"
            End Select
            %>
          </td>
        </tr>
        <tr>
          <td height="200" valign='top'>
            <br>
            <p align='center'>
            <a href='User_Bankroll.asp'><img src='images/detail_all.jpg' border='0' title='所有明细记录'></a>&nbsp;&nbsp;&nbsp;&nbsp;
            <a href='User_Bankroll.asp?ShowType=1'><img src='images/detail_income.jpg' border='0' title='所有收入记录'></a>&nbsp;&nbsp;&nbsp;&nbsp;
            <a href='User_Bankroll.asp?ShowType=2'><img src='images/detail_payout.jpg' border='0' title='所有支出记录'></a>
            </p>
            <% Call PE_Execute("PE_EShop6", "User_Bankroll") %>
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
