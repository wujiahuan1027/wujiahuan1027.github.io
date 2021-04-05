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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">学生学籍管理</font></td>
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
    <td>　　欢迎您进入<%=SiteName%>学生学籍管理模块！本系统主要是对学生信息（如学号、姓名、性别、家庭住址等）及班级信息进行管理，也可以对学生成绩进行录入、修改、删除、查询、打印等操作。您可以对考试信息进行修改、添加删除等管理操作。</td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="../sdms/InfoManage.asp" target=main>学生信息管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class='Class' href="../sdms/TestManage.asp" target=main>考试管理</A></td>
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
    <td width="400" valign="top">　　本模块主要是对学生信息（如学号、姓名、性别、家庭住址等）进行管理。本模块又分为<A href="Admin_Help_Manage.asp#" title="指录入学生的学号、姓名、年级、班级、性别、民族、籍贯、出生日期、联系电话、家庭地址和家长姓名等信息。"><u>录入学生信息</u></A>、<A
href="Admin_Help_Manage.asp#" title="可从学号、姓名或班级三种查询方法任选其一进行查询学生信息。支持模糊查询。此模块必须先登录后才能使用。"><u>查询学生信息</u></A>、<A href="Admin_Help_Manage.asp#" title="首先使用查询功能查询出需要修改/删除的记录，然后进行修改/删除操作。"><u>修改/删除学生信息</u></A>、<A
href="Admin_Help_Manage.asp#" title="将查询结果以我们常见的成绩表形式打印出来，并可以自定义打印格式。"><u>打印学生信息</u></A>、<A href="Admin_Help_Manage.asp#" title="添加/删除班级，结果会直接影响“按班级查询”方式。"><u>班级管理</u></A>和<A
href="Admin_Help_Manage.asp#" title="此模块包括添加/删除/禁用年级，学期末年级自动升级。"><u>年级管理</u></A>六个子模块。<br>
    　　快捷菜单：<A href="../sdms/InfoManage.asp" target=main><font color="#FF0000"><u>学生信息管理</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　在这里您可以添加新的考试信息和考试科目，并可以对已有的考试信息和考试科目进行修改/删除等操作。<br>
　　快捷菜单：<A href="../sdms/TestManage.asp" target=main><font color="#FF0000"><u>考试管理</u></font></A></td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="../sdms/ScoreManage.asp" target=main>学生成绩管理</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center""></td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td height="1" colspan="2" class="topbg2"></td>
    <td colspan="2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">　　本模块主要是对学生成绩进行管理。本模块又分为<A href="Admin_Help_Manage.asp#" title="录入某次考试中一科或多科的成绩。提供两种录入方式：单条记录方式和列表方式。单条记录方式比较简单，但不容易出错。列表方式提供了批量录入的功能，可以一次性录入所有成绩，但容易出错。"><u>录入学生成绩</u></A>、<A href="Admin_Help_Manage.asp#" title="提供两种方式查询学生成绩：按考试查询和按科目查询。按考试查询是指首先选择要查询成绩的考试名称，然后查询这一次考试中的某一科或多科的成绩，这种方式适合某一次考试的横向比较；按科目查询是指首先选择要查询成绩的某一科目，然后查询这一科的某次或多次考试的成绩，这种方式适合科任老师对学生几次考试的成绩进行纵向比较。"><u>查询学生成绩</u></A>、<A href="Admin_Help_Manage.asp#" title="修改/删除某次考试中某个学生的成绩。"><u>修改/删除学生成绩</u></A>、<A href="Admin_Help_Manage.asp#" title="将查询结果以我们常见的成绩表形式打印出来."><u>打印学生成绩</u></A>、<A href="Admin_Help_Manage.asp#" title="自动计算各科总分并按照总分进行全级排名与班级排名，然后把排名结果显示出
来。"><u>计算总分与排名</u></A>、<A href="Admin_Help_Manage.asp#" title="对目标分进行录入/修改/删除等操作。这一功能操作与成绩管理基本相似。因为现在不允许对学生排名，但又要对学生进行评价，所以采用了目标分管理的方法，根据学生的实际情况给每个学生制定了一个目标分，然后进行达标/不达标的评价方法。"><u>目标分管理</u></A>六个子模块。其中，学生成绩查询不需登录即可使用，其他模块则需要先登录后才能使用。<br>
    　　快捷菜单：<A href="../sdms/ScoreManage.asp" target=main><font color="#FF0000"><u>学生成绩管理</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">&nbsp;</td>
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
