<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../conn.asp"-->
<!--#include file="Channel_Config.asp"-->
<%
Call CloseConn

Dim PE_Soft
Call CreateObject_Soft

Sub CreateObject_Soft()
    On Error Resume Next
    Set PE_Soft = Server.CreateObject("PE_CMS6.Soft")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Response.End
    End If
    PE_Soft.iConnStr = ConnStr
    PE_Soft.iSystemDatabaseType = SystemDatabaseType
    PE_Soft.CurrentChannelID = ChannelID
End Sub
%>