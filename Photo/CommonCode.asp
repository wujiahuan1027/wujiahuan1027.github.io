<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../conn.asp"-->
<!--#include file="Channel_Config.asp"-->
<%
Call CloseConn

Dim PE_Photo
Call CreateObject_Photo

Sub CreateObject_Photo()
    On Error Resume Next
    Set PE_Photo = Server.CreateObject("PE_CMS6.Photo")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Response.End
    End If
    PE_Photo.iConnStr = ConnStr
    PE_Photo.iSystemDatabaseType = SystemDatabaseType
    PE_Photo.CurrentChannelID = ChannelID
End Sub
%>