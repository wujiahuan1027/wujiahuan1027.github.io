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
<title><%=SiteName & "--��̨������ҳ"%></title>
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
        <td><%=AdminName%>���ã�������
          <script language="JavaScript" type="text/JavaScript" src="../js/date.js"></script></td>
      </tr>
      <tr>
        <td height="8"><img src="Images/adminmain0line.gif" width="283" height="1" /></td>
      </tr>
      <tr>
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">�̳��ճ���������</font></td>
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
    <td>������ӭ������<%=SiteName%>�̳ǹ���ģ�飡��ϵͳ����ǿ��Ķ��������ʽ�������۹���ͱ�ݵ�����֧�����ܵȹ��ܡ������㲻ͬ��ε��̳����������Բ鿴���۲�Ʒ��ͳ������������ʽ������Ϣ�����ø������۷�������ݵش����û��Ķ����Ȳ���������ޱ�ǿ����̳ǹ���ɣ�</td>
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
    <td width="100" align="center" class="topbg">��������</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">��¼����</td>
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
    <td width="400" valign="top">���������Բ鿴�û���������ϸ��Ϣ����ݵش��������̣����û����¶�������<a href="#" title="ָ������ͻ���ϵ��ȷ���йض�����Ϣ�Ƿ���ȷ��¼����ͻ�̸����ʵ�����ۼۼ��������޵ȡ�"><u>�޸�</u></a>��<a href="#" title="ָȷ�϶�����ȷ����󣬽��˶��������Է��ͻ��޸Ķ������ݡ�����һ����ȷ�Ϻ󣬿ͻ��Ͳ������޸Ļ�ɾ��������ԱҲ�����޸Ķ�����ֻ��ɾ����û�и���Ķ���������������ɴ����̴���ģ���ȷ�Ϻ�ֱ�ӴӴ������ʽ��п۳���Ӧ���ã�ʵ�����ۼ۸񣩣��������Զ�������"><u>ȷ��</u></a>��<a href="#" title="ָȷ�Ͽͻ������л�Ŀ���˺�Ϊ�ͻ�¼������Ϣ��Ȼ��֧���˶������á�"><u>���л��֧��</u></a>��<a href="#" title="ָ�ӿͻ����ʽ������ֱ�ӿ۳���������֧���˶������á�"><u>������пۿ�֧��</u></a>��<a href="#" title="ָ¼��˶����ķ�����Ϣ�����ӿͻ�����̳�û�����ӵ���̳����֤�û������С�"><u>����</u></a>��<a href="#" title="ָ�����ϲ��趼��ɺ󣬿ͻ�Ҳ�Ѿ�����˶��������п���󣬽��˶������塣����󣬴˶��������������Ϣ���������޸ġ�"><u>���嶩��</u></a>��<a href="#" title="ָ��ĳ��
�ͻ��Ķ�����������һ�ͻ�?����ص��ʽ���ϸ�����Ϊԭ�ͻ�? "><u>��������</u></a>�Ȳ�������Ҳ�����ڶ������������ò�ѯ���ܰ���ͬ�������п�ݲ�ѯ��<br>"
    ������ݲ˵���<a href="Admin_Order.asp?SearchType=1" target=main><font color="#FF0000"><u>�������Ķ���</u></font></a> | <a href="Admin_Order.asp" target=main><font color="#FF0000"><u>�������ж���</u></font></a>    </td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">���������Բ�ѯ��Ʒ�����ķ�����¼�������Բ�ѯ���������ж���Ʒ���������ļ�¼���������������ͻ���������Ҳ���Բ�ѯ���������û�����ϸ�����<br>
������ݲ˵���<a href="Admin_Deliver.asp" target=main><font color="#FF0000"><u>���˻���¼</u></font></a> | <a href="Admin_Transfer.asp" target=main><font color="#FF0000"><u>����������¼</u></font></a> </td>
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
    <td width="100" align="center" class="topbg">���۹���</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_PresentProject.asp" target="main">������������</a></td>
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
    <td width="400" valign="top">���������������Բ鿴��Ʒ���۵���ϸ�����ͳ��/���С��������ڿͻ����������ò�ѯ���ܰ���ͬ�������п�ݲ�ѯ��Ҳ�ɰ��������������۽����в鿴�������Խ����۵���ϸ��������EXCEL���Է��������д�ӡ��<br>
    ������ݲ˵���<a href="Admin_OrderItem.asp" target=main><font color="#FF0000"><u>������ϸ���</u></font></a> | <a href="Admin_SaleCount.asp" target=main><font color="#FF0000"><u>����ͳ��/����</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">�����ڴ������������У�����������������Ʒ�Ĵ������������������÷��������ơ���Ч�ڡ�<a href="#" title="ָ������Ʒ�����ܽ�����Ч�۸�Χ����ע�⣺��ͬ���������ļ۸�������벻ͬ������������ͻ��"><u>�۸�����</u></a>��������ݡ��ɶ�δ�ύ������֧������ȡ����ɾ��������<br>
    ������ݲ˵���<a href="Admin_PresentProject.asp" target=main><font color="#FF0000"><u>������������</u></font></a></td>
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
    <td width="100" align="right" class="topbg">�ʽ����</td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg">�����ͻ�����</td>
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
    <td width="400" valign="top">��������֧����¼���Բ鿴����֧����ʽ��֧��ƽ̨������ʱ�䡢����ʵ��ת�˽�����״̬��������Ϣ����ϸ��Ϣ���ɶ�δ�ύ������֧������ȡ����ɾ��������Ҳ���Խ��пͻ�����������ʱ�䡢���׷�ʽ�����֡������֧����ժҪ���������ƺͱ�ע/˵������ϸ���ʽ���ϸ��ѯ��<br>
    ������ݲ˵���<a href="Admin_Payment.asp" target=main><font color="#FF0000"><u>����֧����¼����</u></font></a> | <a href="Admin_Bankroll.asp" target=main><font color="#FF0000"><u>�ʽ���ϸ��ѯ</u></font></a> | <a href="Admin_Bank.asp" target=main><font color="#FF0000"><u>�����ʻ�����</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">�����������Զ��帶�ʽ�����޸�Ĭ�ϵ�����֧�������֧��������ת�ʡ��ʾֻ���������ֽ�֧Ʊ�ȸ��ʽ��Ҳ������ӻ��޸�������ƽ�ʡ���ݡ�EMS��ݵ��ͻ���ʽ��<br>
    ������ݲ˵���<a href="Admin_PaymentType.asp" target=main><font color="#FF0000"><u>���ʽ����</u></font></a> | <a href="Admin_DeliverType.asp" target=main><font color="#FF0000"><u>�ͻ���ʽ����</u></font></a></td>
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
