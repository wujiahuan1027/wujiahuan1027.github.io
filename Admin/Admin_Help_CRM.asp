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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">�ͻ���ϵ����</font></td>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Client.asp" target=main>�ͻ�����</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Contacter.asp" target=main>��ϵ�˹���</A></td>
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
    <td width="400" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;�����Բ鿴�ͻ�����ϸ��Ϣ���Կͻ����в��ҡ��޸ġ�ɾ������ӵȲ������鿴ĳ���ͻ�ʱ�����Ժܷ������Ըÿͻ����пͻ���Ϣ�޸ġ������ϵ�ˡ���ӷ����¼�����Ͷ�߼�¼��������л�����������롢���֧����Ϣ�Ȳ�����<br>
    ��ݲ˵���<a href='Admin_Client.asp?Action=AddClient&ClientType=0' target='main'><font color="#FF0000"><u>�����ҵ�ͻ�</u></font></a> | <a href='Admin_Client.asp?Action=AddClient&ClientType=1' target='main'><font color="#FF0000"><u>��Ӹ��˿ͻ�</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;��Կͻ���������ӵĿͻ��������Ժܷ�����������ĳ���ͻ�����ϵ�ˡ����Ѿ���Ӻõ���ϵ�ˣ���Ҳ���Խ��в鿴���޸ġ�ɾ���Ȳ�����<br>
    ��ݲ˵���<a href='Admin_Contacter.asp' target='main'><font color="#FF0000"><u>��ϵ�˹�����ҳ</u></font></a> | <a href='Admin_Contacter.asp?Action=AddContacter' target='main'><font color="#FF0000"><u>�����ϵ��</u></font></a>
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
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Service.asp" target=main>�������</A></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><A class="Class" href="Admin_Complain.asp" target=main>Ͷ�߹���</A></td>
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
    <td width="400" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;���ñ����ܣ������Ժܷ������ӿͻ������¼��ÿ���ͻ�������Ϣ��������ϸ��¼�������⡢��Ӧ�ͻ����ͻ�����ϵ�ˡ�������Ա���������͵ȵȡ����ÿһ���ͻ������¼������������������¼�Ļط���Ϣ�������ܸ�׼ȷ�����տͻ��ķ�����<br>
    ��ݲ˵���<a href='Admin_Service.asp' target='main'><font color="#FF0000"><u>�ͻ����������ҳ</u></font></a> | <a href='Admin_Service.asp?Action=Add' target='main'><font color="#FF0000"><u>��ӷ����¼</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">&nbsp;&nbsp;&nbsp;&nbsp;���ñ����ܣ������Լ�¼ÿ���ͻ���Ͷ����Ϣ�����Ͷ�����⣬�����ԶԸ�Ͷ����Ϣ�����޸�Ͷ�߼�¼����Ӵ����¼/�޸Ĵ����¼�����Ѵ����Ͷ�ߣ���������ӻطü�¼��<br>
    ��ݲ˵���<a href='Admin_Complain.asp' target='main'><font color="#FF0000"><u>�ͻ�Ͷ�߹�����ҳ</u></font></a> | <a href='Admin_Complain.asp?Action=Add' target='main'><font color="#FF0000"><u>���Ͷ�߼�¼</u></font></a></td>
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
