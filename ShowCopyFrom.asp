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
    Dim PE_ShowSource
    Set PE_ShowSource = Server.CreateObject("PE_CMS6.ShowSource")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_ShowSource.iConnStr = ConnStr
    PE_ShowSource.iCMS_Edition = CMS_Edition
    PE_ShowSource.iSystemDatabaseType = SystemDatabaseType
    PE_ShowSource.iFileName = "ShowCopyFrom.asp"
    Call PE_ShowSource.ShowHTML
    Set PE_ShowSource = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>