<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../conn.asp"-->
<!--#include file="Channel_Config.asp"-->
<%
Call CloseConn

Dim PE_Article
Call CreateObject_Article

Sub CreateObject_Article()
    On Error Resume Next
    Set PE_Article = Server.CreateObject("PE_CMS6.Article")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Response.End
    End If
    PE_Article.iConnStr = ConnStr
    PE_Article.iSystemDatabaseType = SystemDatabaseType
    PE_Article.CurrentChannelID = ChannelID
End Sub
%>