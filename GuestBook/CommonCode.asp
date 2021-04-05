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
        Response.Write "对不起，你的服务器没有安装动易组件（PE_CMS6.dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        Response.End
    End If
    PE_GuestBook.iConnStr = ConnStr
    PE_GuestBook.iSystemDatabaseType = SystemDatabaseType
End Sub
%>