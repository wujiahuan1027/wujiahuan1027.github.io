<%
Const EnableSiteManageCode = True       '是否启用后台管理认证码 是： True  否： False 
Const SiteManageCode = "2006sp1"  '后台管理认证码，您可以修改成您的管理员认证码：×××××××××

'检查后台管理认证码
Sub CheckSiteManageCode()
    If EnableSiteManageCode = True And Trim(Request.Cookies(Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME") & GetInstallPath(Trim(Request.ServerVariables("SCRIPT_NAME")), 1)), "/", ""), ".", ""))("AdminLoginCode")) <> SiteManageCode Then
        Response.Redirect "Admin_login.asp"
        Response.End
    End If
End Sub

Function GetInstallPath(ByVal ScriptName, ParentLevel)
    Dim i
    GetInstallPath = "/"
    If ScriptName = "" Or IsNull(ScriptName) Then Exit Function
    If ParentLevel > 1 Then ParentLevel = 1
    If ParentLevel = 0 Then
        GetInstallPath = Left(ScriptName, InStrRev(ScriptName, "/"))
    ElseIf ParentLevel = 1 Then
        i = InStrRev(ScriptName, "/") - 1
        If i < 1 Then i = 1
        GetInstallPath = Left(ScriptName, InStrRev(ScriptName, "/", i))
    End If
    If Right(GetInstallPath, 1) <> "/" Then GetInstallPath = GetInstallPath & "/"
End Function
%>
