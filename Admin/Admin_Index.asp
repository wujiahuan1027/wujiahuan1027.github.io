<%@language=vbscript codepage=936 %>
<%
Option Explicit
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="Admin_ChkCode.asp"-->
<%
Dim AdminName, AdminPassword, RndPassword, AdminLoginCode
Dim sqlGetAdmin, rsGetAdmin

AdminName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminName")))
AdminPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminPassword")))
RndPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("RndPassword")))
AdminLoginCode = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminLoginCode")))
If AdminName = "" Or AdminPassword = "" Or RndPassword = "" Or (EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode) Then
    Call CloseConn
    Response.Redirect "Admin_login.asp"
End If
sqlGetAdmin = "select * from PE_Admin where AdminName='" & AdminName & "' and Password='" & AdminPassword & "'"
Set rsGetAdmin = Server.CreateObject("adodb.recordset")
rsGetAdmin.Open sqlGetAdmin, Conn, 1, 1
If rsGetAdmin.BOF And rsGetAdmin.EOF Then
    rsGetAdmin.Close
    Set rsGetAdmin = Nothing
    Call CloseConn
    Response.Redirect "Admin_login.asp"
Else
    If rsGetAdmin("EnableMultiLogin") <> True And Trim(rsGetAdmin("RndPassword")) <> RndPassword Then
        Response.write "<br><p align=center><font color='red' style='font-size:9pt'>�Բ���Ϊ��ϵͳ��ȫ����ϵͳ������������ʹ��ͬһ������Ա�ʺŽ��е�¼��</font><br><font style='font-size:9pt'>��Ϊ���������Ѿ��������ط�ʹ�ô˹���Ա�ʺŽ��е�¼�ˣ������㽫���ܼ������к�̨���������<br>�����<a href='Admin_Login.asp' target='_top'>������µ�¼</a>��</font></p>"
        rsGetAdmin.Close
        Set rsGetAdmin = Nothing
        Call CloseConn
        Response.End
    End If
End If

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=SiteName & "--��̨����"%></title>
</head>
<frameset rows="*" cols="200,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
  <frame name="left" scrolling="yes" marginwidth="0" marginheight="0" src="Admin_Index_Left.asp">
  <frameset rows="53,*" cols="*" framespacing="0" border="false" rows="35,*" frameborder="0" scrolling="yes">
    <frame name="top" scrolling="no" src="Admin_Index_Top.asp">
    <frame name="main" scrolling="auto" src="Admin_Index_Main.asp">
  </frameset>
</frameset>
<noframes>
  <body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
  <p>���������汾���ͣ�������ϵͳҪ��IE5�����ϰ汾����ʹ�ñ�ϵͳ��</p>
  </body>
</noframes>
</html>