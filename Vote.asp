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
        Response.Write "对不起，你的服务器没有安装动易组件（PE_CMS6.dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        Exit Sub
    End If
    PE_Site.iConnStr = ConnStr
    PE_Site.iSystemDatabaseType = SystemDatabaseType
    PE_Site.Init
    Call PE_Site.ShowVote
    Set PE_Site = Nothing
    If Err Then
        Response.Write "错 误 号：" & Err.Number & "<BR>"
        Response.Write "错误描述：" & Err.Description & "<BR>"
        Response.Write "错误来源：" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>