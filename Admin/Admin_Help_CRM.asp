<%@language=vbscript codepage=936 %>
<%
Option Explicit
response.buffer = True
Const PurviewLevel = 0
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<%
Call CloseConn
%>
<html>
<head>
<title><%=SiteName & "--后台管理首页"%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Admin_Style.css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="392" rowspan="2"><img src="Images/adminmain01.gif" width="392" height="126"></td>
    <td height="114" valign="top" background="Images/adminmain0line2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="20"></td>
      </tr>
      <tr>
        <td><%=AdminName%>您好，今天是
          <script language="JavaScript" type="text/JavaScript" src="../js/date.js"></script></td>
      </tr>
      <tr>
        <td height="8"><img src="Images/adminmain0line.gif" width="283" height="1" /></td>
      </tr>
      <tr>
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">客户关系管理</font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="9" valign="bottom" background="Images/adminmain03.gif"><img src="Images/adminmain02.gif" width="23" height="12"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Client.asp" target=main>客户管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Contacter.asp" target=main>联系人管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;您可以查看客户的详细信息，对客户进行查找、修改、删除、添加等操作。查看某个客户时，可以很方便的针对该客户进行客户信息修改、添加联系人、添加服务记录、添加投诉记录、添加银行汇款、添加其它收入、添加支出信息等操作。<br>
    快捷菜单：<a href='Admin_Client.asp?Action=AddClient&ClientType=0' target='main'><font color="#FF0000"><u>添加企业客户</u></font></a> | <a href='Admin_Client.asp?Action=AddClient&ClientType=1' target='main'><font color="#FF0000"><u>添加个人客户</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;针对客户管理中添加的客户，您可以很方便的添加属于某个客户的联系人。对已经添加好的联系人，您也可以进行查看、修改、删除等操作。<br>
    快捷菜单：<a href='Admin_Contacter.asp' target='main'><font color="#FF0000"><u>联系人管理首页</u></font></a> | <a href='Admin_Contacter.asp?Action=AddContacter' target='main'><font color="#FF0000"><u>添加联系人</u></font></a>
</td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<table width="100%" height="10"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Service.asp" target=main>服务管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Complain.asp" target=main>投诉管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;利用本功能，您可以很方便的添加客户服务记录，每条客户服务信息都可以详细记录服务主题、对应客户、客户方联系人、服务人员、服务类型等等。针对每一条客户服务记录，还可以输入这条记录的回访信息，让您能更准确地掌握客户的反馈。<br>
    快捷菜单：<a href='Admin_Service.asp' target='main'><font color="#FF0000"><u>客户服务管理首页</u></font></a> | <a href='Admin_Service.asp?Action=Add' target='main'><font color="#FF0000"><u>添加服务记录</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;利用本功能，您可以记录每个客户的投诉信息。点击投诉主题，您可以对该投诉信息进行修改投诉记录、添加处理记录/修改处理记录，对已处理的投诉，还可以添加回访记录。<br>
    快捷菜单：<a href='Admin_Complain.asp' target='main'><font color="#FF0000"><u>客户投诉管理首页</u></font></a> | <a href='Admin_Complain.asp?Action=Add' target='main'><font color="#FF0000"><u>添加投诉记录</u></font></a></td>
    <td width="21">&nbsp;</td>
  </tr>
</table>
<br>
<table cellpadding="2" cellspacing="1" border="0" width="100%" class="border" align=center>
  <tr align="center">
    <td height=25 class="topbg"><span class="Glow">Copyright 2003-2006 &copy; <%=SiteName%> All Rights Reserved.</span>
  </tr>
</table>
</body>
</html>
