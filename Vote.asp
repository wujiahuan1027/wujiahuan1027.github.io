<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="conn.asp"-->
<%
Call CloseConn
Call Main

Sub Main()
    On Error Resume Next
    Dim PE_Site
    Set PE_Site = Server.CreateObject("PE_CMS6.Site")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_Site.iConnStr = ConnStr
    PE_Site.iSystemDatabaseType = SystemDatabaseType
    PE_Site.Init
    Call PE_Site.ShowVote
    Set PE_Site = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>