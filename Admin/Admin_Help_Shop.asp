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
        <td><img src="Images/img_u.gif" align="absmiddle">您现在进行的是<font color="#FF0000">商城日常操作管理</font></td>
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
    <td>　　欢迎您进入<%=SiteName%>商城管理模块！本系统具有强大的订单管理、资金管理、销售管理和便捷的在线支付功能等功能。能满足不同层次的商城需求。您可以查看销售产品的统计情况，进行资金管理信息，设置各种销售方案，便捷地处理用户的订单等操作。体会无比强大的商城管理吧：</td>
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
    <td width="100" align="center" class="topbg">订单管理</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">记录管理</td>
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
    <td width="400" valign="top">　　您可以查看用户订单的详细信息，便捷地处理订单流程，对用户所下订单进行<a href="#" title="指根据与客户联系，确认有关订单信息是否正确，录入与客户谈定的实际销售价及服务期限等。"><u>修改</u></a>、<a href="#" title="指确认订单正确无误后，将此订单锁定以防客户修改订单内容。订单一旦被确认后，客户就不可再修改或删除，管理员也不能修改订单，只能删除还没有付款的订单。如果订单是由代理商代理的，则确认后将直接从代理商资金中扣除相应费用（实际销售价格），并立即自动发货。"><u>确认</u></a>、<a href="#" title="指确认客户从银行汇的款项到账后，为客户录入汇款信息，然后支付此订单费用。"><u>银行汇款支付</u></a>、<a href="#" title="指从客户的资金余额中直接扣除款项用于支付此订单费用。"><u>从余额中扣款支付</u></a>、<a href="#" title="指录入此订单的发货信息，并加客户的论坛用户名添加到论坛的认证用户名单中。"><u>发货</u></a>、<a href="#" title="指当以上步骤都完成后，客户也已经付清此订单的所有款项后，将此订单结清。结清后，此订单的所有相关信息都不可再修改。"><u>结清订单</u></a>和<a href="#" title="指将某个
客户的订单过户给另一客户?但相关的资金明细情况仍为原客户? "><u>订单过户</u></a>等操作。您也可以在订单处理顶部利用查询功能按不同条件进行快捷查询。<br>"
    　　快捷菜单：<a href="Admin_Order.asp?SearchType=1" target=main><font color="#FF0000"><u>处理今天的订单</u></font></a> | <a href="Admin_Order.asp" target=main><font color="#FF0000"><u>处理所有订单</u></font></a>    </td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　您可以查询商品订单的发货记录，并可以查询订单功能中对商品订单过户的记录。单击订单编号与客户姓名，您也可以查询到订单与用户的详细情况。<br>
　　快捷菜单：<a href="Admin_Deliver.asp" target=main><font color="#FF0000"><u>发退货记录</u></font></a> | <a href="Admin_Transfer.asp" target=main><font color="#FF0000"><u>订单过户记录</u></font></a> </td>
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
    <td width="100" align="center" class="topbg">销售管理</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_PresentProject.asp" target="main">促销方案管理</a></td>
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
    <td width="400" valign="top">　　在这里您可以查看商品销售的明细情况和统计/排行。您可以在客户管理顶部利用查询功能按不同条件进行快捷查询，也可按销售数量、销售金额进行查看，并可以将销售的明细情况输出至EXCEL，以方便您进行打印。<br>
    　　快捷菜单：<a href="Admin_OrderItem.asp" target=main><font color="#FF0000"><u>销售明细情况</u></font></a> | <a href="Admin_SaleCount.asp" target=main><font color="#FF0000"><u>销售统计/排行</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　在促销方案管理中，您可以添加与管理商品的促销方案。您可以设置方案的名称、有效期、<a href="#" title="指促销商品购物总金额的有效价格范围。请注意：不同促销方案的价格区间必须不同，否则会产生冲突。"><u>价格区间</u></a>与促销内容。可对未提交的在线支付进行取消或删除操作。<br>
    　　快捷菜单：<a href="Admin_PresentProject.asp" target=main><font color="#FF0000"><u>促销方案管理</u></font></a></td>
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
    <td width="100" align="right" class="topbg">资金管理</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">付款送货管理</td>
    <td width="300">&nbsp;</td>
    <td width="21" rowspan="2">&nbsp;</td>
  </tr>
  <tr class="topbg2">
    <td height="1" colspan="2"></td>
    <td colspan="2" class="topbg2"></td>
  </tr>
</table>
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20">&nbsp;</td>
    <td width="400" valign="top">　　在线支付记录可以查看在线支付形式的支付平台、交易时间、汇款金额、实际转账金额、交易状态、银行信息等详细信息。可对未提交的在线支付进行取消或删除操作。也可以进行客户姓名、交易时间、交易方式、币种、收入金额、支出金额、摘要、银行名称和备注/说明等详细的资金明细查询。<br>
    　　快捷菜单：<a href="Admin_Payment.asp" target=main><font color="#FF0000"><u>在线支付记录管理</u></font></a> | <a href="Admin_Bankroll.asp" target=main><font color="#FF0000"><u>资金明细查询</u></font></a> | <a href="Admin_Bank.asp" target=main><font color="#FF0000"><u>银行帐户管理</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">　　您可以自定义付款方式，或修改默认的在线支付、余额支付、银行转帐、邮局汇款、货到付款、现金、支票等付款方式。也可以添加或修改如邮政平邮、快递、EMS快递等送货方式。<br>
    　　快捷菜单：<a href="Admin_PaymentType.asp" target=main><font color="#FF0000"><u>付款方式管理</u></font></a> | <a href="Admin_DeliverType.asp" target=main><font color="#FF0000"><u>送货方式管理</u></font></a></td>
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
