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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">ѧ��ѧ������</font></td>
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
    <td>������ӭ������<%=SiteName%>ѧ��ѧ������ģ�飡��ϵͳ��Ҫ�Ƕ�ѧ����Ϣ����ѧ�š��������Ա𡢼�ͥסַ�ȣ����༶��Ϣ���й���Ҳ���Զ�ѧ���ɼ�����¼�롢�޸ġ�ɾ������ѯ����ӡ�Ȳ����������ԶԿ�����Ϣ�����޸ġ����ɾ���ȹ��������</td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="../sdms/InfoManage.asp" target=main>ѧ����Ϣ����</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class='Class' href="../sdms/TestManage.asp" target=main>���Թ���</A></td>
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
    <td width="400" valign="top">������ģ����Ҫ�Ƕ�ѧ����Ϣ����ѧ�š��������Ա𡢼�ͥסַ�ȣ����й�����ģ���ַ�Ϊ<A href="Admin_Help_Manage.asp#" title="ָ¼��ѧ����ѧ�š��������꼶���༶���Ա����塢���ᡢ�������ڡ���ϵ�绰����ͥ��ַ�ͼҳ���������Ϣ��"><u>¼��ѧ����Ϣ</u></A>��<A
href="Admin_Help_Manage.asp#" title="�ɴ�ѧ�š�������༶���ֲ�ѯ������ѡ��һ���в�ѯѧ����Ϣ��֧��ģ����ѯ����ģ������ȵ�¼�����ʹ�á�"><u>��ѯѧ����Ϣ</u></A>��<A href="Admin_Help_Manage.asp#" title="����ʹ�ò�ѯ���ܲ�ѯ����Ҫ�޸�/ɾ���ļ�¼��Ȼ������޸�/ɾ��������"><u>�޸�/ɾ��ѧ����Ϣ</u></A>��<A
href="Admin_Help_Manage.asp#" title="����ѯ��������ǳ����ĳɼ�����ʽ��ӡ�������������Զ����ӡ��ʽ��"><u>��ӡѧ����Ϣ</u></A>��<A href="Admin_Help_Manage.asp#" title="���/ɾ���༶�������ֱ��Ӱ�조���༶��ѯ����ʽ��"><u>�༶����</u></A>��<A
href="Admin_Help_Manage.asp#" title="��ģ��������/ɾ��/�����꼶��ѧ��ĩ�꼶�Զ�������"><u>�꼶����</u></A>������ģ�顣<br>
    ������ݲ˵���<A href="../sdms/InfoManage.asp" target=main><font color="#FF0000"><u>ѧ����Ϣ����</u></font></A></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">��������������������µĿ�����Ϣ�Ϳ��Կ�Ŀ�������Զ����еĿ�����Ϣ�Ϳ��Կ�Ŀ�����޸�/ɾ���Ȳ�����<br>
������ݲ˵���<A href="../sdms/TestManage.asp" target=main><font color="#FF0000"><u>���Թ���</u></font></A></td>
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
    <td width="100" align="center" class="topbg"><A class='Class' href="../sdms/ScoreManage.asp" target=main>ѧ���ɼ�����</A></td>
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
    <td width="400" valign="top">������ģ����Ҫ�Ƕ�ѧ���ɼ����й�����ģ���ַ�Ϊ<A href="Admin_Help_Manage.asp#" title="¼��ĳ�ο�����һ�ƻ��Ƶĳɼ����ṩ����¼�뷽ʽ��������¼��ʽ���б�ʽ��������¼��ʽ�Ƚϼ򵥣��������׳����б�ʽ�ṩ������¼��Ĺ��ܣ�����һ����¼�����гɼ��������׳���"><u>¼��ѧ���ɼ�</u></A>��<A href="Admin_Help_Manage.asp#" title="�ṩ���ַ�ʽ��ѯѧ���ɼ��������Բ�ѯ�Ͱ���Ŀ��ѯ�������Բ�ѯ��ָ����ѡ��Ҫ��ѯ�ɼ��Ŀ������ƣ�Ȼ���ѯ��һ�ο����е�ĳһ�ƻ��Ƶĳɼ������ַ�ʽ�ʺ�ĳһ�ο��Եĺ���Ƚϣ�����Ŀ��ѯ��ָ����ѡ��Ҫ��ѯ�ɼ���ĳһ��Ŀ��Ȼ���ѯ��һ�Ƶ�ĳ�λ��ο��Եĳɼ������ַ�ʽ�ʺϿ�����ʦ��ѧ�����ο��Եĳɼ���������Ƚϡ�"><u>��ѯѧ���ɼ�</u></A>��<A href="Admin_Help_Manage.asp#" title="�޸�/ɾ��ĳ�ο�����ĳ��ѧ���ĳɼ���"><u>�޸�/ɾ��ѧ���ɼ�</u></A>��<A href="Admin_Help_Manage.asp#" title="����ѯ��������ǳ����ĳɼ�����ʽ��ӡ����."><u>��ӡѧ���ɼ�</u></A>��<A href="Admin_Help_Manage.asp#" title="�Զ���������ֲܷ������ֽܷ���ȫ��������༶������Ȼ������������ʾ��
����"><u>�����ܷ�������</u></A>��<A href="Admin_Help_Manage.asp#" title="��Ŀ��ֽ���¼��/�޸�/ɾ���Ȳ�������һ���ܲ�����ɼ�����������ơ���Ϊ���ڲ������ѧ������������Ҫ��ѧ���������ۣ����Բ�����Ŀ��ֹ���ķ���������ѧ����ʵ�������ÿ��ѧ���ƶ���һ��Ŀ��֣�Ȼ����д��/���������۷�����"><u>Ŀ��ֹ���</u></A>������ģ�顣���У�ѧ���ɼ���ѯ�����¼����ʹ�ã�����ģ������Ҫ�ȵ�¼�����ʹ�á�<br>
    ������ݲ˵���<A href="../sdms/ScoreManage.asp" target=main><font color="#FF0000"><u>ѧ���ɼ�����</u></font></A></td>
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
