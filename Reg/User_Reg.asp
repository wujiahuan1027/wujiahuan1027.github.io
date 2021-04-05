<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
'网站配置的变量
Dim strHTML, PE_Site

Call Main
Call CloseConn

Sub Main()
    If EnableUserReg <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，本站暂停新用户注册服务！</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Set PE_Site = Server.CreateObject("PE_CMS6.Site")
    PE_Site.iConnStr = ConnStr
    PE_Site.iSystemDatabaseType = SystemDatabaseType
    PE_Site.CurrentChannelID = 0
    PE_Site.Init
    
    strHTML = PE_Site.GetTemplate(0, 18, 0)
    strHTML = PE_Site.ReplaceCommon(strHTML)

    Dim strPath
    strPath = "您现在的位置：&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;&gt;&gt;&nbsp;服务条款和声明"

    strHTML = Replace(strHTML, "{$PageTitle}", SiteTitle & " >> 服务条款和声明")
    strHTML = Replace(strHTML, "{$ShowPath}", strPath)

    strHTML = Replace(strHTML, "{$MenuJS}", PE_Site.GetMenuJS("", False))
    strHTML = Replace(strHTML, "{$Skin_CSS}", PE_Site.GetSkin_CSS(0))
    
    Set PE_Site = Nothing
    Response.Write strHTML
End Sub
%>