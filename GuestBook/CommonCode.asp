<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../conn.asp"-->
<%
Call CloseConn

Dim PE_GuestBook
Call CreateObject_GuestBook

Sub CreateObject_GuestBook()
    On Error Resume Next
    Set PE_GuestBook = Server.CreateObject("PE_CMS6.GuestBook")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Response.End
    End If
    PE_GuestBook.iConnStr = ConnStr
    PE_GuestBook.iSystemDatabaseType = SystemDatabaseType
End Sub
%>