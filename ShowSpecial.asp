<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.buffer = True
%>
<!--#include file="conn.asp"-->
<%
Call CloseConn
Call Main

Sub Main()
    On Error Resume Next
    Dim PE_SiteSpecial
    Set PE_SiteSpecial = Server.CreateObject("PE_CMS6.SiteSpecial")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_SiteSpecial.iConnStr = ConnStr
    PE_SiteSpecial.iSystemDatabaseType = SystemDatabaseType
    Call PE_SiteSpecial.ShowSpecial
    Set PE_SiteSpecial = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>