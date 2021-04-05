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
        Response.Write "对不起，你的服务器没有安装动易组件（PE_CMS6.dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        Response.End
    End If
    PE_Photo.iConnStr = ConnStr
    PE_Photo.iSystemDatabaseType = SystemDatabaseType
    PE_Photo.CurrentChannelID = ChannelID
End Sub
%>