<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
Server.ScriptTimeOut = 9999999
%>
<!--#include file="../Conn.asp"-->
<%
Call CloseConn
Call Main

Sub Main()
    On Error Resume Next
    Dim PE_Upload
    Set PE_Upload = Server.CreateObject("PE_Upload6.UpFile")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_Upload6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_Upload.iConnStr = ConnStr
    PE_Upload.iSystemDatabaseType = SystemDatabaseType
    Call PE_Upload.ShowUploadForm
    Set PE_Upload = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>