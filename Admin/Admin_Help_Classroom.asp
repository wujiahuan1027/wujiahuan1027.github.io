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
        <td><img src="Images/img_u.gif" align="absmiddle">�����ڽ��е���<font color="#FF0000">�ҳ��Ǽǹ���</font></td>
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
    <td>������ӭ������<%=SiteName%>�ҳ��Ǽǹ���ģ�飡ѧУ������ר�ý��ң������ҡ���ý����ҡ����Խ��ҵȣ���ѧ������ʦʹ��ʱҪ�����Ǽ�һ�£�������ѧУ֪��ĳλ��ʦ��ѧ������ʱ��ʹ����ʲôר�ý��ң�����ѧУѧ�ڽ������ˡ��ҳ��Ǽǹ���ģ��һ��������ѧУ�����豸�й���Ϣ��������һ�������ڶ�ʹ���ߣ�ѧ�����ʦ���ӵǼǵ�ȷ�ϵǼǣ���ʹ�õ�ȷ��ʹ�����̵Ĺ����Լ�ʹ�ú���ϸ��¼�Ĳ�ѯ��</td>
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
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_ClassroomSort.asp" target="main">�豸����</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_ManageRecord.asp" target=main>ʹ�õǼǹ���</a></td>
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
    <td width="400" valign="top">������������Ӻ͹����ҳ�������ָ���ҳ������ҳ����ơ�֧�����޸��ҳ����ÿ������¿��������޸��ҳ����ҳ�����ҳ����豸֮���м̳��ԣ�ɾ���ҳ��ǵ���ɾ���ҳ��µ��豸��ɾ���ҳ����ǵ���ɾ���ҳ�����������豸��<br>
    ������ݲ˵���<a href="Admin_ClassroomSort.asp" target=main><font color="#FF0000"><u>�豸����</u></font></a></td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">�����ѵǼǵ��ҳ��豸�������Ƕ���ͬʱԤԼ�Ǽ�Ҫʹ��ͬһ̨�豸������Աֻȷ������һ������ʹ��Ȩ����ȷ������һ�˵�ͬʱ������ͬһʱ��ԤԼ�Ǽ�ʹ�ø��豸���˵ĵǼ���Ϣ����ɾ��������ȷ�ϵ��ҳ��豸��Ҳ���ܳ��ָõǼ���������ԭ��û���ڸ�ʱ��ʹ�ø��豸������������ҳ��ǼǵĹ���Ա����ȷ���¸����ڸ�ʱ�����Ѿ�ʹ�ø��豸��<br>
������ݲ˵���<a href="Admin_ManageRecord.asp" target=main><font color="#FF0000"><u>ʹ�õǼǹ���</u></font></a></td>
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
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_Equipment.asp" target="main">�豸����</a></td>
    <td width="300">&nbsp;</td>
    <td width="40" rowspan="2">&nbsp;</td>
    <td width="100" align="center" class="topbg"><a class='Class' href="Admin_SearchHistory.asp" target="main">ʹ�ü�¼��ѯ</a></td>
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
    <td width="400" valign="top">������������������Ӻ͹�����ҳ����豸��ÿ���ҳ��п��������޸��豸���ṩ�豸���ơ��豸��ֵ���豸״̬���豸������Ϣ�������豸����<br>
    ������ݲ˵���<a href="Admin_Equipment.asp" target=main><font color="#FF0000"><u>�豸����</u></font></a> </td>
    <td width="40">&nbsp;</td>
    <td width="400" valign="top">�������Ѿ��Ǽǲ���ʹ�õĵļ�¼���Բ鿴�б��Լ�ÿ��ʹ�ü�¼��ʹ�����飬����ʹ���ߵĵǼ�ip���Ǽ�ʱ��ȵȡ�<br>
    ������ݲ˵���<a href="Admin_SearchHistory.asp"><font color="#FF0000"><u>ʹ�ü�¼��ѯ</u></font></a></td>
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
    <td width="100" align="center" class="topbg">ʹ�õǼ�˵��</td>
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
    <td width="400" valign="top">������Ա���û��������Ǽ�ĳ��ĳ�ڿ�ĳ����ʹ��ĳ�豸��ԤԼ�Ǽ���ֻ��Ҫָ����ѯ������ĳ��ĳ�ڿ�ĳ�����������Բ�ѯ�����ҳ������豸���豸�ġ��Ǽ�״̬��Ϊ��δ�Ǽǻ��ѵǼǡ��Ŀ��ԵǼ�ʹ�á�֧���ظ��Ǽǣ����Ǽ�״̬���У�δ�Ǽǡ��ѵǼǡ���ȷ�ϡ���ʹ�á��������Ҫ�á�����ʾ�Ǽǽ�����顣</td>
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
