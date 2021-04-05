<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Expires = -1
Response.Buffer = False
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/md5.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%
Dim sql, rs
Dim CookieDate
Dim UserPassword, RndPassword, CheckCode, TempUserName
Action = Trim(Request("action"))

If Action = "xmlstat" Then
    Response.ContentType = "text/xml; charset=gb2312"
ElseIf Action = "xml" Then
    Dim UserInfReceived, rootNode
    Set UserInfReceived = CreateObject("Microsoft.XMLDOM")
    UserInfReceived.async = False
    UserInfReceived.Load Request
    Set rootNode = UserInfReceived.getElementsByTagName("root")
    If rootNode.Length < 1 Then
        FoundErr = True
        ErrMsg = ErrMsg & "输入数据为空"
    Else
        UserName = ReplaceBadChar(rootNode(0).selectSingleNode("username").Text)
        UserPassword = ReplaceBadChar(rootNode(0).selectSingleNode("password").Text)
        CheckCode = LCase(ReplaceBadChar(rootNode(0).selectSingleNode("checkcode").Text))
        CookieDate = rootNode(0).selectSingleNode("cookiesdate").Text
        If CookieDate = "" Or (Not IsNumeric(CookieDate)) Then
            CookieDate = 0
        Else
            CookieDate = CLng(CookieDate)
        End If
        If UserName = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "用户名不能为空！"
        Else
            TempUserName = UserName
        End If
        If UserPassword = "" Then
            FoundErr = True
            ErrMsg = ErrMsg & "密码不能为空！"
        End If
    End If
    Set UserInfReceived = Nothing
    Response.ContentType = "text/xml; charset=gb2312"
Else
    UserName = ReplaceBadChar(Trim(Request("UserName")))
    UserPassword = ReplaceBadChar(Trim(Request("UserPassword")))
    CheckCode = LCase(ReplaceBadChar(Trim(Request("CheckCode"))))
    CookieDate = Trim(Request("CookieDate"))
    If CookieDate = "" Or (Not IsNumeric(CookieDate)) Then
        CookieDate = 0
    Else
        CookieDate = CLng(CookieDate)
    End If
    ComeUrl = Trim(Request("ComeUrl"))
    If InStr(ComeUrl, "Reg/") > 0 Then ComeUrl = strInstallDir & "User/Index.asp"
    If ComeUrl = "" Then
        ComeUrl = Request.ServerVariables("HTTP_REFERER")
    End If
    If ComeUrl = "" Then ComeUrl = strInstallDir & "Index.asp"
    If UserName = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>用户名不能为空！</li>"
    Else
        TempUserName = UserName
    End If
    If UserPassword = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>密码不能为空！</li>"
    End If
End If
Dim strTempMsg,iIndex,arrAPIs,strLoginParams
If FoundErr <> True Then
    '保存用户名
    TempUserName = UserName
    If CheckUserLogined() = False Then
        If Action = "xmlstat" Then
            FoundErr = True
            ErrMsg = ""
        Else
            '恢复可能被替换的用户名
            UserName = TempUserName
            sPE_Items(conPassword,1) = UserPassword
            UserPassword = MD5(UserPassword, 16)
            Set rs = Server.CreateObject("adodb.recordset")
            sql = "select UserID,UserName,UserPassword,LastPassword,LastLoginIP,LastLoginTime,LoginTimes from PE_User where UserName='" & UserName & "'"
            rs.Open sql, Conn, 1, 3
            If rs.bof And rs.EOF Then
                FoundErr = True
                ErrMsg = ErrMsg & "{a}用户不存在！！！{b}"
            Else
                If UserPassword <> rs(2) Then
                    Dim tempPassword
                    tempPassword = sPE_Items(conPassword,1)
                    MD5OLD = 0
                    tempPassword = MD5(tempPassword,16)
                    Md5OLD = 1
                    If tempPassword <> rs(2) Then
                        FoundErr = True
                        ErrMsg = ErrMsg & "{a}密码错误！！！{b}"
                    Else
                        rs("UserPassword") = UserPassword
                    End If
                Else
                    If EnableCheckCodeOfLogin = True Then
                        If Trim(Session("CheckCode")) = "" Then
                            FoundErr = True
                            ErrMsg = ErrMsg & "{a}验证码超时失效。{b}"
                        End If
                        If CheckCode <> Session("CheckCode") Then
                            FoundErr = True
                            ErrMsg = ErrMsg & "{a}验证码错误，请重新输入。{b}"
                        End If
                    End If
                    '加入整合接口支持
                    If Not FoundErr Then
                        If API_Enable Then
                            If createXmlDom Then
                                sPE_Items(conAction, 1) = "login"
                                sPE_Items(conUsername, 1) = UserName
                                'sPE_Items(conPassword, 1) = UserPassword
                                sPE_Items(conSavecookie, 1) = CookieDate
                                sPE_Items(conUserip, 1) = UserTrueIP
                                prepareXml True
                                SendPost
                                If FoundErr Then
                                    ErrMsg = "{a}" & ErrMsg & "{b}"
                                End If
                            Else
                                FoundErr = True
                                ErrMsg = ErrMsg & "{a}登陆服务暂时不可用。[APIError-XmlDom-Runtime]{b}"
                            End If
                        End If
                    End If
                    '完毕
                    If Not FoundErr Then
                        RndPassword = GetRndPassword(16)
                        rs("LastPassword") = RndPassword
                        rs("LastLoginIP") = UserTrueIP
                        rs("LastLoginTime") = Now()
                        rs("LoginTimes") = rs("LoginTimes") + 1
                        rs.Update
                        Select Case CookieDate
                            Case 0
                                'not save
                            Case 1
                                Response.Cookies(Site_Sn).Expires = Date + 1
                            Case 2
                                Response.Cookies(Site_Sn).Expires = Date + 31
                            Case 3
                                Response.Cookies(Site_Sn).Expires = Date + 365
                        End Select
                        Response.Cookies(Site_Sn)("UserName") = UserName
                        Response.Cookies(Site_Sn)("UserPassword") = UserPassword
                        Response.Cookies(Site_Sn)("LastPassword") = RndPassword
                        Response.Cookies(Site_Sn)("CookieDate") = CookieDate
                        Dim iNum, questurl, xmlquest
                        If Action = "xml" Then
                            Call CheckUserLogined
                            Call showuserxml
                        Else
                            If API_Enable Then
                                sPE_Items(conSyskey,1) = MD5(UserName & API_Key,16)
                                sPE_Items(conUsername,1) = UserName
                                sPE_Items(conPassword,1) = UserPassword
                                sPE_Items(conSavecookie,1) = CookieDate
                                strLoginParams = "?syskey=" & sPE_Items(conSyskey,1) & "&username=" & sPE_Items(conUsername,1) & "&password=" & sPE_Items(conPassword,1) & "&savecookie=" & sPE_Items(conSavecookie,1)
                                For iIndex = 0 To UBound(arrAPIUrls)
                                    arrAPIs = Split(arrAPIUrls(iIndex), "@@")
                                    strTempMsg = strTempMsg & "<script type=""text/javascript"" language=""JavaScript"" src=""" & arrAPIs(1) & strLoginParams & """ charset=""gb2312""></script>"
                                Next
                            End If
                            strTempMsg = "您已成功登陆，欢迎您的光临!" & strTempMsg
                            Call WriteSuccessMsg(strTempMsg, ComeUrl)
                        End If
                    End If
                End If
            End If
            rs.Close
            Set rs = Nothing
        End If
    Else
        UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
        If Action = "xmlstat" Then
            Call showuserxml
        Else
            If API_Enable Then
                sPE_Items(conSyskey,1) = MD5(UserName & API_Key,16)
                sPE_Items(conUsername,1) = UserName
                sPE_Items(conPassword,1) = UserPassword
                sPE_Items(conSavecookie,1) = CookieDate
                strLoginParams = "?syskey=" & sPE_Items(conSyskey,1) & "&username=" & sPE_Items(conUsername,1) & "&password=" & sPE_Items(conPassword,1) & "&savecookie=" & sPE_Items(conSavecookie,1)
                For iIndex = 0 To UBound(arrAPIUrls)
                    arrAPIs = Split(arrAPIUrls(iIndex), "@@")
                    strTempMsg = strTempMsg & "<script type=""text/javascript"" language=""JavaScript"" src=""" & arrAPIs(1) & strLoginParams &""" charset=""gb2312""></script>"
                Next
            End If
            strTempMsg = "您已成功登陆，欢迎您的光临!" & strTempMsg
            Call WriteSuccessMsg(strTempMsg, ComeUrl)
        End If
    End If
End If
If FoundErr = True Then
    If Action = "xml" Or Action = "xmlstat" Then
        Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
        Response.Write "<body>"
        If UserName <> "" Then
            Response.Write "<user>" & UserName & "</user>"
        Else
            UserName = Request.Cookies("UserName")
            Response.Write "<user>" & UserName & "</user>"
        End If
        Response.Write "<checkstat>err</checkstat>"
        ErrMsg = Replace(Replace(ErrMsg, "{a}", ""), "{b}", "")
        Response.Write "<errsource>" & ErrMsg & "</errsource>"
        If EnableCheckCodeOfLogin = True Then
            Response.Write "<checkcode>1</checkcode>"
        Else
            Response.Write "<checkcode>0</checkcode>"
        End If
        If API_Enable And UserName <> "" Then
            sPE_Items(conSyskey,1) = MD5(UserName&API_Key,16)
            Response.Write "<syskey>" & sPE_Items(conSyskey,1) & "</syskey>"
            Dim intIndex,tmpUrls
            For intIndex = 0 To Ubound(arrAPIUrls)
                tmpUrls = Split(arrAPIUrls(intIndex),"@@")
                Response.Write "<apiurl><![CDATA[" & tmpUrls(1) & "]]></apiurl>"
            Next
            Response.Write "<savecookie/>"
        Else
            Response.Write "<syskey/><apiurl/><savecookie/>"
        End If
        Response.Write "</body>"
    Else
        ErrMsg = Replace(Replace(ErrMsg, "{a}", "<br><li>"), "{b}", "</li>")
        Call WriteErrMsg
    End If
End If
Call CloseConn

Sub showuserxml()
    Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"
    Response.Write "<body>"
    Response.Write "<user>" & UserName & "</user>"
    Response.Write "<userid>" & UserID & "</userid>"
    Response.Write "<userpass>" & UserPassword & "</userpass>"
    Response.Write "<usertype>" & UserType & "</usertype>"
    Response.Write "<groupname>" & GroupName & "</groupname>"
    Response.Write "<grouptype>" & GroupType & "</grouptype>"
    Response.Write "<checkstat>ok</checkstat>"
    Response.Write "<balance>" & Balance & "</balance>"
    Response.Write "<exp>" & UserExp & "</exp>"
    Response.Write "<point>"
    Response.Write "    <pointname>" & PointName & "</pointname>"
    Response.Write "    <userpoint>" & UserPoint & "</userpoint>"
    Response.Write "    <unit>" & PointUnit & "</unit>"
    Response.Write "</point>"
    If ValidNum = -1 Then
        Response.Write "<day>unlimit</day>"
    Else
        Response.Write "<day>" & ValidDays & "</day>"
    End If
    If Trim(UnsignedItems & "") = "" Then
        Response.Write "<article>0</article>"
    Else
        Dim UnsignedItemNum, arrUser
        arrUser = Split(UnsignedItems, ",")
        UnsignedItemNum = UBound(arrUser) + 1
        Response.Write "<article>" & UnsignedItemNum & "</article>"
    End If
    Response.Write "<logined>" & LoginTimes & "</logined>"
    If UnreadMsg <> "" And CLng(UnreadMsg) > 0 Then
        Response.Write "<message>" & UnreadMsg & "</message>"
        Dim MessageID, rsMessage
        Set rsMessage = Conn.Execute("select id,sender,title,sendtime from PE_Message where incept='" & UserName & "'and delR=0 and flag=0 and IsSend=1")
        If rsMessage.bof And rsMessage.EOF Then
            Response.Write "<unreadmessage><stat>empty</stat></unreadmessage>"
        Else
            Response.Write "<unreadmessage><stat>full</stat>"
            Do While Not rsMessage.EOF
                Response.Write "<item>"
                Response.Write "<id>" & rsMessage("id") & "</id>"
                Response.Write "<sender>" & rsMessage("sender") & "</sender>"
                Response.Write "<title>" & rsMessage("title") & "</title>"
                Response.Write "<time>" & rsMessage("sendtime") & "</time>"
                Response.Write "</item>"
                rsMessage.movenext
            Loop
            Response.Write "</unreadmessage>"
        End If
        rsMessage.Close
        Set rsMessage = Nothing
    Else
        Response.Write "<message>0</message>"
        Response.Write "<unreadmessage><stat>empty</stat></unreadmessage>"
    End If
    If API_Enable Then
        sPE_Items(conSyskey,1) = MD5(UserName&API_Key,16)
        Response.Write "<syskey>" & sPE_Items(conSyskey,1) & "</syskey>"
        Dim intIndex,tmpUrls
        For intIndex = 0 To Ubound(arrAPIUrls)
            tmpUrls = Split(arrAPIUrls(intIndex),"@@")
            Response.Write "<apiurl><![CDATA[" & tmpUrls(1) & "]]></apiurl>"
        Next
        If CookieDate = "" Then CookieDate = 0
        Response.Write "<savecookie>" & CookieDate & "</savecookie>"
    Else
        Response.Write "<syskey/><apiurl/><savecookie/>"
    End If
    Response.Write "</body>"
End Sub

'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************

Sub WriteErrMsg()
    Dim strErr
    strErr = strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    strErr = strErr & "<link href='../Images/style.css' rel='stylesheet' type='text/css'></head><body>" & vbCrLf
    strErr = strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    strErr = strErr & "  <tr align='center'><td height='22' class='title'><strong>错误信息</strong></td></tr>" & vbCrLf
    strErr = strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>产生错误的可能原因：</b><br>" & ErrMsg & "</td></tr>" & vbCrLf
    strErr = strErr & "  <tr align='center'><td class='tdbg'><a href=""User_Login.asp?ComeUrl=" & ComeUrl & """>&lt;&lt; 返回登录页面</a></td></tr>" & vbCrLf
    strErr = strErr & "</table>" & vbCrLf
    strErr = strErr & "</body></html>" & vbCrLf
    Response.Write strErr
End Sub
%>