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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">室场登记管理</font></td>
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
    <td width="20">&nbsp;</td>
    <td>　　欢迎您进入<%=SiteName%>室场登记管理模块！学校有若干专用教室（语音室、多媒体教室、电脑教室等），学生、教师使用时要上网登记一下，可以让学校知道某位教师、学生，何时，使用了什么专用教室，便于学校学期结束考核。室场登记管理模块一方面用于学校公共设备有关信息管理，另外一方面用于对使用者（学生或教师）从登记到确认登记，从使用到确认使用流程的管理以及使用后详细记录的查询。</td>
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
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_ClassroomSort.asp" target="main">设备管理</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_ManageRecord.asp" target=main>使用登记管理</a></td>
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
    <td width="400" valign="top">　　您可以添加和管理室场，可以指定室场类别或室场名称。支持无限个室场类别。每个类别下可以有无限个室场。室场类别、室场及设备之间有继承性，删除室场记得先删除室场下的设备，删除室场类别记得先删除室场及其里面的设备。<br>
    　　快捷菜单：<a href="Admin_ClassroomSort.asp" target=main><font color="#FF0000"><u>设备管理</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　已登记的室场设备，可以是多人同时预约登记要使用同一台设备，管理员只确认其中一个人有使用权（在确认其中一人的同时，其他同一时间预约登记使用该设备的人的登记信息将被删除）；已确认的室场设备，也可能出现该登记者有其他原因没有在该时间使用该设备的情况，管理室场登记的管理员可以确认下该人在该时间下已经使用该设备。<br>
　　快捷菜单：<a href="Admin_ManageRecord.asp" target=main><font color="#FF0000"><u>使用登记管理</u></font></a></td>
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
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_Equipment.asp" target="main">设备管理</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_SearchHistory.asp" target="main">使用记录查询</a></td>
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
    <td width="400" valign="top">　　在这里您可以添加和管理各室场的设备。每个室场中可以有无限个设备，提供设备名称、设备价值、设备状态、设备简介等信息、方便设备管理。<br>
    　　快捷菜单：<a href="Admin_Equipment.asp" target=main><font color="#FF0000"><u>设备管理</u></font></a> </td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　对已经登记并且使用的的记录可以查看列表以及每条使用记录的使用详情，包括使用者的登记ip，登记时间等等。<br>
    　　快捷菜单：<a href="Admin_SearchHistory.asp"><font color="#FF0000"><u>使用记录查询</u></font></a></td>
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
    <td width="100" align="center">&nbsp;</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">使用登记说明</td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="1" colspan="2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top"><br>    </td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　会员在用户控制面板登记某天某节课某场所使用某设备，预约登记者只需要指定查询条件：某天某节课某场所，即可以查询出该室场所有设备。设备的“登记状态”为“未登记或已登记”的可以登记使用。支持重复登记，“登记状态”有：未登记、已登记、已确认、已使用。点击“我要用”即显示登记结果详情。</td>
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
