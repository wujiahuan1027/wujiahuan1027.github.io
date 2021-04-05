<%@language=vbscript codepage=936 %>
<!--#include file="../Conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../inc/Md5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<%

Action = Trim(Request("action"))
Dim MemberName,APISysKey
MemberName = Request.Cookies(Site_Sn)("UserName")
APISysKey = MD5(MemberName&API_Key,16)
'Response.Cookies(Site_Sn)("UserName") = ""
Response.Cookies(Site_Sn)("UserPassword") = ""
Response.Cookies(Site_Sn)("LastPassword") = ""

If Action <> "xml" Then
    Dim strTempMsg
    If API_Enable Then
        Dim iIndex,strLogoutParams
        strLogoutParams = "?syskey=" & APISysKey & "&username=" & MemberName
        For iIndex = 0 to Ubound(arrAPIUrls)
            Dim arrAPIs
            arrAPIs = Split(arrAPIUrls(iIndex),"@@")
            strTempMsg = strTempMsg & "<script type=""text/javascript"" language=""JavaScript"" src=""" & arrAPIs(1) & strLogoutParams &""" charset=""gb2312""></script>"
        Next
    End If
    strTempMsg = "您已成功注销，期待您的再次光临!" & strTempMsg
    Call WriteSuccessMsg(strTempMsg,InstallDir & "Index.asp")
Else
    Response.Clear
    Response.ContentType = "text/xml; charset=gb2312"
    Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
    Response.Write "<body>"
    Response.Write "<user>" & MemberName & "</user>"
    Response.Write "<checkstat>err</checkstat>"
    Response.Write "<errsource/>"
    Response.Write "<checkcode>0</checkcode>"
    If API_Enable And MemberName <> "" Then
        Response.Write "<syskey>" & APISysKey & "</syskey>"
        Dim intIndex,tmpUrls
        For intIndex = 0 To Ubound(arrAPIUrls)
            tmpUrls = Split(arrAPIUrls(intIndex),"@@")
            Response.Write "<apiurl><![CDATA[" & tmpUrls(1) & "]]></apiurl>"
        Next
    Else
        Response.Write "<syskey/><apiurl/>"
    End If
    Response.Write "<savecookie/>"
    Response.Write "</body>"
End If
Call CloseConn
%>