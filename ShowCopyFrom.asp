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
        Response.Write "对不起，你的服务器没有安装动易组件（PE_CMS6.dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        Exit Sub
    End If
    PE_ShowSource.iConnStr = ConnStr
    PE_ShowSource.iCMS_Edition = CMS_Edition
    PE_ShowSource.iSystemDatabaseType = SystemDatabaseType
    PE_ShowSource.iFileName = "ShowCopyFrom.asp"
    Call PE_ShowSource.ShowHTML
    Set PE_ShowSource = Nothing
    If Err Then
        Response.Write "错 误 号：" & Err.Number & "<BR>"
        Response.Write "错误描述：" & Err.Description & "<BR>"
        Response.Write "错误来源：" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>