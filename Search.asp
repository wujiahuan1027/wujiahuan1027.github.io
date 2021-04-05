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
    Dim PE_Search, ModuleName
    ModuleName = LCase(Trim(Request("ModuleName")))
    If ModuleName = "" Then ModuleName = "article"
    Select Case ModuleName
    Case "article"
        Set PE_Search = Server.CreateObject("PE_CMS6.Article")
    Case "soft"
        Set PE_Search = Server.CreateObject("PE_CMS6.Soft")
    Case "photo"
        Set PE_Search = Server.CreateObject("PE_CMS6.Photo")
    Case "shop"
        Set PE_Search = Server.CreateObject("PE_EShop6.Product")
    Case Else
        Set PE_Search = Server.CreateObject("PE_CMS6.Article")
    End Select
    If Err Then
        Err.Clear
        If ModuleName = "shop" Then
            Response.Write "对不起，你的服务器没有安装动易组件（PE_EShop6.dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        Else
            Response.Write "对不起，你的服务器没有安装动易组件（PE_CMS6.dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        End If
        Exit Sub
    End If
    PE_Search.iConnStr = ConnStr
    PE_Search.iSystemDatabaseType = SystemDatabaseType
    PE_Search.CurrentChannelID = 0
    Call PE_Search.ShowSearch
    Set PE_Search = Nothing
    If Err Then
        Response.Write "错 误 号：" & Err.Number & "<BR>"
        Response.Write "错误描述：" & Err.Description & "<BR>"
        Response.Write "错误来源：" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>