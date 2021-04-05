<!--#include file="../Conn.asp"-->
<%
Response.Expires = -1
Response.ContentType = "text/xml; charset=gb2312"
Call CloseConn
Call Main

Sub Main()
    On Error Resume Next
    Dim PE_Blog
    Set PE_Blog = Server.CreateObject("PE_CMS6.ShowBlog")
    If Err Then
        Err.Clear
        Response.Write "对不起，你的服务器没有安装动易组件（PE_CMS6.dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        Exit Sub
    End If
    PE_Blog.iConnStr = ConnStr
    PE_Blog.iSystemDatabaseType = SystemDatabaseType
    PE_Blog.iShowType = "ShowMusic"
    Call PE_Blog.ShowHTML
    Set PE_Blog = Nothing
    If Err Then
        Response.Write "错 误 号：" & Err.Number & "<BR>"
        Response.Write "错误描述：" & Err.Description & "<BR>"
        Response.Write "错误来源：" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>