<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
Server.ScriptTimeOut = 9999999
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_ChkCode.asp"-->
<%
Call CheckSiteManageCode
Call CloseConn
Call PE_Execute("PE_CMS6", "Xml")

Sub PE_Execute(strDllName, strClassName)
    On Error Resume Next
    If strDllName = "" Or IsNull(strDllName) Then
        Response.Write "请指定动易组件名！"
        Exit Sub
    End If
    If strClassName = "" Or IsNull(strClassName) Then
        Response.Write "请指定动易组件提供的类名！"
        Exit Sub
    End If
    Dim PE_Admin, objName
    objName = strDllName & "." & strClassName
    Set PE_Admin = Server.CreateObject(objName)
    If Err Then
        Err.Clear
        Response.Write "对不起，你的服务器没有安装动易组件（" & strDllName & ".dll），所以不能使用动易系统。请和你的空间商联系以安装动易组件。"
        Exit Sub
    End If
    PE_Admin.iConnStr = ConnStr
    PE_Admin.iCMS_Edition = CMS_Edition
    PE_Admin.iSystemDatabaseType = SystemDatabaseType
    Call PE_Admin.main
    Set PE_Admin = Nothing
    If Err Then
        Response.Write "错 误 号：" & Err.Number & "<BR>"
        Response.Write "错误描述：" & Err.Description & "<BR>"
        Response.Write "错误来源：" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>