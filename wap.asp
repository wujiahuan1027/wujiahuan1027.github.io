<%@language=vbscript codepage=65001 %>
<!--#include file="conn.asp"-->
<%
Dim ReadMe
ReadMe = Trim(Request("ReadMe"))
If ReadMe = "Yes" Then
%>
<html>
<title>WAP�����</title>
<link href='Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="160" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr valign="top"><td><img src="Images/WapBack01.gif" width="160" height="48"></td>
  </tr>
  <tr height="140">
    <td height="153" valign="middle" background="Images/WapBack02.gif">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="2" colspan="3"></td>
        </tr>
        <tr>
          <td width="30"></td>
          <td width="112" valign='top' style="word-break:break-all;Width:fixed"><font color="#FFFFFF">��ܰ��ʾ����վ�ѿ�ͨWAP�����������ֻ�֧��WAP���ܣ�����ʹ���ֻ����ʣ�<br><% =SiteUrl %>/wap.asp</font></td>
          <td width="18">&nbsp;</td>
        </tr>
        <tr>
          <td height="2" colspan="3"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr><td><img src="Images/WapBack03.gif" width="160" height="56"></td></tr>
</table>
</body>
</html>
<%
Else
    Response.ContentType = "text/vnd.wap.wml; charset=utf-8"
    Call CloseConn
    Call Main
End If

Sub Main()
    On Error Resume Next
    Dim PE_Index
    Set PE_Index = Server.CreateObject("PE_CMS6.ShowWap")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_Index.iConnStr = ConnStr
    PE_Index.iSystemDatabaseType = SystemDatabaseType
    Call PE_Index.Main
    Set PE_Index = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>