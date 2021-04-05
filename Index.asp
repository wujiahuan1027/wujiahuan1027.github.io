<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="conn.asp"-->
<%
If FileExt_SiteIndex < 4 Then
    Call CloseConn
    Response.redirect "Index" & GetFileExt(FileExt_SiteIndex)
Else
    Call CloseConn
    Call Main
End If

Function GetFileExt(FileExtType)
    Select Case FileExtType
    Case 0
        GetFileExt = ".html"
    Case 1
        GetFileExt = ".htm"
    Case 2
        GetFileExt = ".shtml"
    Case 3
        GetFileExt = ".shtm"
    Case 4
        GetFileExt = ".asp"
    End Select
End Function

Sub Main()
    On Error Resume Next
    Dim PE_Index
    Set PE_Index = Server.CreateObject("PE_CMS6.CreateIndex")
    If Err Then
        Err.Clear
        Response.Write "对不起，你的服务器没有安装动易组件（PE_CMS6.dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        Exit Sub
    End If
    PE_Index.iConnStr = ConnStr
    PE_Index.iSystemDatabaseType = SystemDatabaseType
    Call PE_Index.ShowHTML
    Set PE_Index = Nothing
    If Err Then
        Response.Write "错 误 号：" & Err.Number & "<BR>"
        Response.Write "错误描述：" & Err.Description & "<BR>"
        Response.Write "错误来源：" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>