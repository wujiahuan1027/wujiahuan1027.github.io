<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
'��վ���õı���
Dim strHTML, PE_Site

Call Main
Call CloseConn

Sub Main()
    If EnableUserReg <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Բ��𣬱�վ��ͣ���û�ע�����</li>"
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
    strPath = "�����ڵ�λ�ã�&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;&gt;&gt;&nbsp;�������������"

    strHTML = Replace(strHTML, "{$PageTitle}", SiteTitle & " >> �������������")
    strHTML = Replace(strHTML, "{$ShowPath}", strPath)

    strHTML = Replace(strHTML, "{$MenuJS}", PE_Site.GetMenuJS("", False))
    strHTML = Replace(strHTML, "{$Skin_CSS}", PE_Site.GetSkin_CSS(0))
    
    Set PE_Site = Nothing
    Response.Write strHTML
End Sub
%>