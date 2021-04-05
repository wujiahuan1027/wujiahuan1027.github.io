<!--#include file="Admin_ChkCode.asp"-->
<%
Dim ScriptName, TrueSiteUrl, cUrl
Dim AdminName, AdminPassword, RndPassword, AdminLoginCode, AdminPurview, PurviewPassed
Dim AdminPurview_Channel, AdminPurview_Others, AdminPurview_GuestBook
Dim rsGetAdmin, sqlGetAdmin
Dim arrPurview(25), PurviewIndex, strThisFile
Dim Channel, Name, Content, UploadDir
Dim ChannelID, sqlChannel, rsChannel, ChannelName, ChannelShortName, ChannelDir, ModuleType, ModuleName, SheetName

ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
ScriptName = Trim(Request.ServerVariables("SCRIPT_NAME"))
TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
If ComeUrl = "" Then
    Response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></p>"
    Call Insert_Pur_Log
    Response.End
Else
    cUrl = Trim("http://" & TrueSiteUrl) & ScriptName
    If LCase(Left(ComeUrl, InStrRev(ComeUrl, "/"))) <> LCase(Left(cUrl, InStrRev(cUrl, "/"))) Then
        Response.write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许从外部链接地址访问本系统的后台管理页面。</font></p>"
        Call Insert_Pur_Log
        Response.End
    End If
End If

'检查管理员是否登录
strInstallDir = GetScriptPath(ScriptName, 1)
Site_Sn = Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME") & strInstallDir), "/", ""), ".", "")
AdminName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminName")))
AdminPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminPassword")))
RndPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("RndPassword")))
AdminLoginCode = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("AdminLoginCode")))
If AdminName = "" Or AdminPassword = "" Or RndPassword = "" Or (EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode) Then
    Call Insert_Pur_Log
    Call CloseConn
    Response.redirect "Admin_login.asp"
End If

ChannelID = Trim(Request("ChannelID"))
If ChannelID = "" Then
    ChannelID = 0
Else
    ChannelID = CLng(ChannelID)
End If
If ChannelID > 0 Then
    sqlChannel = "select * from PE_Channel where ChannelID=" & ChannelID & " order by OrderID"
    Set rsChannel = Server.CreateObject("adodb.recordset")
    rsChannel.Open sqlChannel, Conn, 1, 1
    If rsChannel.BOF And rsChannel.EOF Then
        CheckChannelPurview = False
    Else
        If rsChannel("Disabled") = True Then
            Response.write "此频道已经被禁用！"
            Response.End
        End If
        ChannelName = rsChannel("ChannelName")
        ChannelShortName = rsChannel("ChannelShortName")
        ChannelDir = rsChannel("ChannelDir")
        ModuleType = rsChannel("ModuleType")
        UploadDir = rsChannel("UploadDir")
        Select Case ModuleType
        Case 1
            ModuleName = "Article"
            SheetName = "PE_Article"
        Case 2
            ModuleName = "Soft"
            SheetName = "PE_Soft"
        Case 3
            ModuleName = "Photo"
            SheetName = "PE_Photo"
        Case 5
            ModuleName = "Product"
            SheetName = "PE_Product"
        Case 6
            ModuleName = "Supply"
            SheetName = "PE_Supply"
        End Select
    End If
    rsChannel.Close
    Set rsChannel = Nothing
End If


sqlGetAdmin = "select * from PE_Admin where AdminName='" & AdminName & "' and Password='" & AdminPassword & "'"
Set rsGetAdmin = Server.CreateObject("adodb.recordset")
rsGetAdmin.Open sqlGetAdmin, Conn, 1, 1
If rsGetAdmin.BOF And rsGetAdmin.EOF Then
    Call Insert_Pur_Log
    rsGetAdmin.Close
    Set rsGetAdmin = Nothing
    Call CloseConn
    Response.redirect "Admin_login.asp"
Else
    If rsGetAdmin("EnableMultiLogin") <> True And Trim(rsGetAdmin("RndPassword")) <> RndPassword Then
        Response.write "<br><p align=center><font color='red'>对不起，为了系统安全，本系统不允许两个人使用同一个管理员帐号进行登录！</font></p><p>因为现在有人已经在其他地方使用此管理员帐号进行登录了，所以你将不能继续进行后台管理操作。</p><p>你可以<a href='Admin_Login.asp' target='_top'>点此重新登录</a>。</p>"
        Call Insert_Pur_Log
        rsGetAdmin.Close
        Set rsGetAdmin = Nothing
        Call CloseConn
        Response.End
    End If
End If
AdminPurview = rsGetAdmin("Purview")
AdminPurview_Others = rsGetAdmin("AdminPurview_Others")
AdminPurview_GuestBook = rsGetAdmin("AdminPurview_GuestBook")
If AdminPurview = 1 Then
    PurviewPassed = True
Else
    If PurviewLevel = 0 Then      '不进行权限检查
        PurviewPassed = True
    Else
        If AdminPurview > PurviewLevel Then
            PurviewPassed = False
        Else
            If ChannelID > 0 Then
                AdminPurview_Channel = rsGetAdmin("AdminPurview_" & ChannelDir)
                If AdminPurview_Channel = "" Then
                    AdminPurview_Channel = 5
                Else
                    AdminPurview_Channel = CLng(AdminPurview_Channel)
                End If
                If AdminPurview_Channel > PurviewLevel_Channel Then
                    PurviewPassed = False
                Else
                    PurviewPassed = True
                End If
            Else
                PurviewPassed = CheckPurview_Other(AdminPurview_Others, PurviewLevel_Others)
            End If
        End If
    End If
End If
If PurviewLevel > 0 Then
    rsGetAdmin.Close
    Set rsGetAdmin = Nothing
End If

If PurviewPassed = False Then
    Response.write "<br><p align=center><font color='red'>对不起，你没有此项操作的权限。</font></p>"
    Response.End
End If

Function CheckPurview_Other(AllPurviews, strPurview)
    If IsNull(AllPurviews) Or AllPurviews = "" Or strPurview = "" Then
        CheckPurview_Other = False
        Exit Function
    End If
    CheckPurview_Other = False
    If InStr(AllPurviews, ",") > 0 Then
        Dim arrPurviews, i
        arrPurviews = Split(AllPurviews, ",")
        For i = 0 To UBound(arrPurviews)
            If Trim(arrPurviews(i)) = strPurview Then
                CheckPurview_Other = True
                Exit For
            End If
        Next
    Else
        If AllPurviews = strPurview Then
            CheckPurview_Other = True
        End If
    End If
End Function



Function CheckClassMaster(AllMaster, MasterName)
    If IsNull(AllMaster) Or AllMaster = "" Or MasterName = "" Then
        CheckClassMaster = False
        Exit Function
    End If
    CheckClassMaster = False
    If InStr(AllMaster, "|") > 0 Then
        Dim arrMaster, i
        arrMaster = Split(AllMaster, "|")
        For i = 0 To UBound(arrMaster)
            If Trim(arrMaster(i)) = MasterName Then
                CheckClassMaster = True
                Exit For
            End If
        Next
    Else
        If AllMaster = MasterName Then
            CheckClassMaster = True
        End If
    End If
End Function

Function Insert_Pur_Log()
    Action = ""
    Channel = -1
    If ComeUrl = "" Then
        Content = "直接地址输入访问后台"
        Name = ""
    ElseIf LCase(Left(ComeUrl, InStrRev(ComeUrl, "/"))) <> LCase(Left(cUrl, InStrRev(cUrl, "/"))) Then
        Content = "外部链接访问后台"
        Name = ""
    ElseIf AdminName = "" Or AdminPassword = "" Or RndPassword = "" Or (EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode) Then
        Content = "管理员未登录"
        Name = ""
    ElseIf rsGetAdmin.BOF And rsGetAdmin.EOF Then
        Content = "用户名或密码错误"
        Name = AdminName
    ElseIf rsGetAdmin("EnableMultiLogin") <> True And Trim(rsGetAdmin("RndPassword")) <> RndPassword Then
        Content = "两人使用同一管理员帐号"
        Name = AdminName
    Else
        Channel = 0
        Name = AdminName
        Content = "登录成功"
    End If
    Dim sqlLog, rsLog
    sqlLog = "select top 1 * from PE_Log"
    Set rsLog = Server.CreateObject("Adodb.RecordSet")
    rsLog.Open sqlLog, Conn, 1, 3
    rsLog.addnew
    rsLog("LogType") = 1
    rsLog("ChannelID") = Channel
    rsLog("LogTime") = Now()
    rsLog("UserName") = Name
    rsLog("UserIP") = UserTrueIP
    rsLog("LogContent") = Content
    rsLog("ScriptName") = ComeUrl
    rsLog("PostString") = ""
    rsLog.Update
    rsLog.Close
    Set rsLog = Nothing
End Function
%>