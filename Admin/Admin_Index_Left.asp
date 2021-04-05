<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.buffer = True
Const PurviewLevel = 0
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<%
Response.Write "<html>" & vbCrLf
Response.Write "<head>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<title>管理导航菜单</title>" & vbCrLf
Response.Write "<script src=""../JS/prototype.js""></script>"
Response.Write "<link href='Admin_left.CSS' rel='stylesheet' type='text/css'>" & vbCrLf
Response.Write "</head>" & vbCrLf
Response.Write "<BODY leftmargin='0' topmargin='0' marginheight='0' marginwidth='0'>" & vbCrLf
Response.Write "<table width=180 border='0' align=center cellpadding=0 cellspacing=0>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=44 valign=top><img src='Images/title.gif'></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "<table cellpadding=0 cellspacing=0 width=180 align=center>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=26 class=menu_title onmouseover=""this.className='menu_title2';"" onmouseout=""this.className='menu_title';"" background='Images/title_bg_quit.gif' id='menuTitle0'>&nbsp;&nbsp;<a href='Admin_Index_Main.asp' target='main'><b><span class='glow'>管理首页</span></b></a><span class='glow'> | </span><a href='Admin_Login.asp?Action=Logout' target='_top'><b><span class='glow'>退出</span></b></a> </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=97 background='Images/title_bg_admin.gif' style='display:' id='submenu0'><div style='width:180'>" & vbCrLf
Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
Response.Write "          <tr>" & vbCrLf
Response.Write "            <td height=25>您的用户名：" & AdminName & "</td>" & vbCrLf
Response.Write "          </tr>" & vbCrLf
Response.Write "          <tr>" & vbCrLf
Response.Write "            <td height=25>您的身份："
Select Case AdminPurview
Case 1
    Response.Write "超级管理员"
Case 2
    Response.Write "<a href='Admin_ShowPurview.asp' target='main'>普通管理员</a>"
End Select
Dim Message
Set Message = Conn.Execute("select UnreadMsg from PE_User where LastPassword = '" & RndPassword & "'")
If Message.EOF And Message.Bof Then
    UnreadMsg = 0
Else
    UnreadMsg = Message(0)
End If
Set Message = Nothing
Response.Write "</td></tr><tr><td height=20>待阅短信：" & vbCrLf
If UnreadMsg <> "" And PE_CLng(UnreadMsg) > 0 Then
    Response.Write " <b><font color=red>" & UnreadMsg & "</font></b> 条"
Else
    Response.Write " <b><font color=gray>0</font></b> 条"
End If
Response.Write "            </td>" & vbCrLf
Response.Write "          </tr>" & vbCrLf
Response.Write "        </table>" & vbCrLf
Response.Write "      </div>" & vbCrLf
Response.Write "        <div  style='width:167'>" & vbCrLf
Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
Response.Write "            <tr>" & vbCrLf
Response.Write "              <td height=20></td>" & vbCrLf
Response.Write "            </tr>" & vbCrLf
Response.Write "          </table>" & vbCrLf
Response.Write "      </div></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Dim ShowCreateHTML, strActionLink, ShowAdmin_Guest, ShowAdmin_Shop
ShowAdmin_Shop = True
ShowCreateHTML = False
Dim ObjInstalled_FSO
ObjInstalled_FSO = IsObjInstalled(objName_FSO)

sqlChannel = "select * from PE_Channel where ChannelType<=1"
If (CMS_Edition <= 0 And eShop_Edition <= 0 And CRM_Edition <= 0) Or Not FoundInArr(AllModules, "Supply", ",") Then
    sqlChannel = sqlChannel & " and ModuleType<>6"
End If
sqlChannel = sqlChannel & " and ModuleType<>7 and ModuleType<>8 order by OrderID"
Set rsChannel = Server.CreateObject("adodb.recordset")
rsChannel.open sqlChannel, Conn, 1, 1
Do While Not rsChannel.EOF
    If rsChannel("ModuleType") = 4 Then
        If rsChannel("Disabled") = True Then
            ShowAdmin_Guest = False
        Else
            ShowAdmin_Guest = True
        End If
    Else
        If rsChannel("Disabled") = False Then
            ChannelID = rsChannel("ChannelID")
            ChannelName = Trim(rsChannel("ChannelName"))
            ChannelShortName = Trim(rsChannel("ChannelShortName"))
            ChannelDir = Trim(rsChannel("ChannelDir"))
            Select Case rsChannel("ModuleType")
            Case 1
                ModuleName = "Article"
            Case 2
                ModuleName = "Soft"
            Case 3
                ModuleName = "Photo"
            Case 5
                ModuleName = "Product"
            Case 6
                ModuleName = "Supply"
            End Select
            AdminPurview_Channel = rsGetAdmin("AdminPurview_" & ChannelDir)
            If IsNull(AdminPurview_Channel) Then
                AdminPurview_Channel = 5
            Else
                AdminPurview_Channel = CLng(AdminPurview_Channel)
            End If

            
            If AdminPurview = 1 Or AdminPurview_Channel <= 3 Then
                Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center><tr>" & vbCrLf
                Response.Write "<td height=28 class=menu_title onmouseover=""this.className='menu_title2'""; onmouseout=""this.className='menu_title'""; background='Images/Admin_left_" & rsChannel("ModuleType") & ".gif' id=menuTitle" & ChannelID & " onclick=""new Element.toggle('submenu" & ChannelID & "')"" style='cursor:hand;'>" & vbCrLf
                If rsChannel("ModuleType") = 6 Then
                    Response.Write "<a href='Admin_Help_Supply.asp?ChannelID=" & ChannelID & "' target=main><span class=glow>" & ChannelName & "管理</span></a>"
                Else
                    Response.Write "<a href='Admin_Help_Channel.asp?ChannelID=" & ChannelID & "' target=main><span class=glow>" & ChannelName & "管理</span></a>"
                End If
                Response.Write "</td></tr><tr><td style='display:none' align='right' id='submenu" & ChannelID & "'><div class=sec_menu style='width:165'><table cellpadding=0 cellspacing=0 align=center width=132>" & vbCrLf
                If rsChannel("ModuleType") = 5 Then
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&ProductKind=1' target=main>添加" & ChannelShortName & "（实物）</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&ProductKind=2' target=main>添加" & ChannelShortName & "（软件）</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&ProductKind=3' target=main>添加" & ChannelShortName & "（点卡）</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_Card.asp' target=main>充值卡管理</a></td></tr>" & vbCrLf
                Else
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&AddType=1' target=main>添加" & ChannelShortName & "</a></td></tr>" & vbCrLf
                    If rsChannel("ModuleType") = 1 And CMS_Edition > 2 Then
                        Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Receive&Status=9' target=main>签收" & ChannelShortName & "管理</a></td></tr>"
                    End If
                    If rsChannel("ModuleType") = 2 Then
                        Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Add&AddType=3' target=main>添加" & ChannelShortName & "（镜像模式）</a></td></tr>" & vbCrLf
                    End If
                    If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                        If rsChannel("ModuleType") = 3 Then
                            Response.Write "<tr><td height=20><a href='Admin_Photo.asp?ChannelID=" & ChannelID & "&Action=Add&AddType=3' target=main>添加" & ChannelShortName & "（批量模式）</a></td></tr>" & vbCrLf
                        End If
                    End If
                End If
                Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=My" & ModuleName & "&Status=9' target=main>我添加的" & ChannelShortName & "</a></td></tr>" & vbCrLf
                Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage' target=main>" & ChannelShortName & "管理</a>"
                If rsChannel("ModuleType") = 5 Then
                    Response.Write " | <a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Price&Status=0' target=main>价格设置</a>"
                Else
                    Response.Write " | <a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Check&Status=0' target=main>审核</a>"
                End If
                If rsChannel("UseCreateHTML") > 0 And ObjInstalled_FSO = True Then
                    ShowCreateHTML = True
                    Response.Write " | <a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&Status=3&ManageType=HTML' target=main>生成</a>" & vbCrLf
                    If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
                        strActionLink = strActionLink & "<tr height='20'><td><a href='Admin_CreateHTML.asp?ChannelID=" & ChannelID & "' target=main>" & ChannelName & "生成管理</a></td></tr>"
                    End If
                End If
                Response.Write "</td></tr>"
                If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Special&Status=9' target=main>专题" & ChannelShortName & "管理</a></td></tr>" & vbCrLf
                End If
                If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
                    If rsChannel("ModuleType") = 2 Then
                        Response.Write "<tr><td height=20><a href='Admin_Soft.asp?ChannelID=" & ChannelID & "&Action=ShowReplace' target=main>下载地址批量修改</a></td></tr>" & vbCrLf
                    End If
                    Response.Write "<tr><td height=20><a href='Admin_Comment.asp?ChannelID=" & ChannelID & "' target=main>" & ChannelShortName & "评论管理</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Manage&ManageType=Recyclebin' target=main>" & ChannelShortName & "回收站管理</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=10>=====================</td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_Channel.asp?Action=Modify&iChannelID=" & ChannelID & "' target=main>" & ChannelName & "设置</a></td></tr>" & vbCrLf
                    Response.Write "<tr><td height=20><a href='Admin_Class.asp?ChannelID=" & ChannelID & "' target=main>栏目管理</a> | <a href='Admin_Special.asp?ChannelID=" & ChannelID & "' target=main>专题管理</a></td></tr>" & vbCrLf
                    Select Case rsChannel("ModuleType")
                    Case 1, 5
                        Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=UploadFiles' target=main>上传文件管理</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&Action=Clear&UploadDir=UploadFiles' target=main>清理</a></td></tr>" & vbCrLf
                    Case 2
                        Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=UploadSoftPic' target=main>上传图片管理</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&Action=Clear&UploadDir=UploadSoftPic' target=main>清理</a></td></tr>" & vbCrLf
                        Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=UploadSoft' target=main>上传文件管理</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&Action=Clear&UploadDir=UploadSoft' target=main>清理</a></td></tr>" & vbCrLf
                    Case 3, 6
                        Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & ChannelID & "&UploadDir=UploadPhotos' target=main>上传图片管理</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & ChannelID & "&Action=Clear&UploadDir=UploadPhotos' target=main>清理</a></td></tr>" & vbCrLf
                    End Select

                    If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Menu_" & ChannelDir) = True Then
                        Response.Write "<tr><td height=20><a href='Admin_RootClass_Menu.asp?ChannelID=" & ChannelID & "' target=main>顶部菜单设置</a> | <a href='Admin_RootClass_Menu.asp?Action=ShowCreate&ChannelID=" & ChannelID & "' target=main>生成</a></td></tr>" & vbCrLf
                    End If
                    If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Template_" & ChannelDir) = True Or CheckPurview_Other(AdminPurview_Others, "JsFile_" & ChannelDir) = True Then
                        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Template_" & ChannelDir) = True Then
                            Response.Write "<tr><td height=20><a href='Admin_Template.asp?ChannelID=" & ChannelID & "' target=main>" & ChannelShortName & "模板页管理</a></td></tr>" & vbCrLf
                        End If
                        If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "JsFile_" & ChannelDir) = True And ObjInstalled_FSO = True) And (rsChannel("ModuleType") <> 6) Then
                            Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & "JS.asp?ChannelID=" & ChannelID & "' target=main>" & ChannelShortName & "JS文件管理</a></td></tr>" & vbCrLf
                        End If
                    End If
                    If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Keyword_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                        Response.Write "<tr><td height=20><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Keyword' target='main'>" & ChannelShortName & "关键字管理</a></td></tr>"
                    End If
                    If rsChannel("ModuleType") = 5 Then
                        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Producer_" & ChannelDir) = True Then
                            Response.Write "<tr><td height=20><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Producer' target='main'>厂商管理</a></td></tr>"
                        End If
                        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Trademark_" & ChannelDir) = True Then
                            Response.Write "<tr><td height=20><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Trademark' target='main'>品牌管理</a></td></tr>"
                        End If
                    Else
                        If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Author_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                            Response.Write "<tr><td height=20><a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=Author' target='main'>作者管理</a>"
                        End If
                        If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Copyfrom_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                            Response.Write " | <a href='Admin_SourceManage.asp?ChannelID=" & ChannelID & "&TypeSelect=CopyFrom' target='main'>来源管理</a></td></tr>"
                        End If
                    End If
                    If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "XML_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                        Response.Write "<tr><td height=20><a href='Admin_CreateXml.asp?Action=GreatNav&ChannelID=" & ChannelID & "' target=main>更新栏目XML数据</a></td></tr>" & vbCrLf
                    End If
                    If rsChannel("ModuleType") = 2 Then
                        Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=Other' target=main>其他管理</a></td></tr>" & vbCrLf
                        Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&Action=DownError' target=main>错误信息管理</a></td></tr>" & vbCrLf
                        Response.Write "<tr><td height=20><a href='Admin_DownServer.asp?ChannelID=" & ChannelID & "' target=main>镜像服务器管理</a></td></tr>" & vbCrLf
                    End If
                    If (AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Field_" & ChannelDir) = True) And (rsChannel("ModuleType") <> 6) Then
                        Response.Write "<tr><td height=20><a href='Admin_Field.asp?ChannelID=" & ChannelID & "' target=main>自定义字段管理</a></td></tr>" & vbCrLf
                    End If
                    If rsChannel("ModuleType") = 1 And CMS_Edition > 2 Then
                        If FoundInArr(rsChannel("arrEnabledTabs"), "Copyfee", ",") = True Then
                            Response.Write "<tr><td height=20><a href='Admin_" & ModuleName & ".asp?ChannelID=" & ChannelID & "&ManageType=PayMoney&PayStatus=False' target=main>稿费管理</a></td></tr>"
                        End If
                    End If
                End If
                Response.Write "</table></div><div  style='width:158'><table cellpadding=0 cellspacing=0 align=center width=130><tr><td height=4></td></tr></table></div></td></tr></table>" & vbCrLf
            End If
        Else
            If rsChannel("ModuleType") = 5 Then ShowAdmin_Shop = False
        End If
    End If
    rsChannel.MoveNext
Loop
rsChannel.Close
Set rsChannel = Nothing

If CMS_Edition > 0 Or eShop_Edition > 0 Or CRM_Edition > 0 Then
    Dim rsHouse, rsHouseConfig
    Set rsHouse = Conn.Execute("Select ChannelID, ChannelDir, ChannelName, Disabled from PE_Channel Where ModuleType=7")
    If FoundInArr(AllModules, "House", ",") And rsHouse("Disabled") = False Then
        ChannelDir = rsHouse("ChannelDir")
        AdminPurview_Channel = rsGetAdmin("AdminPurview_" & ChannelDir)
        If AdminPurview = 1 Or AdminPurview_Channel <= 3 Then
            Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_7.gif' id=menuTitle502 onclick=""new Element.toggle('submenu502')"" style='cursor:hand;'><a href='Admin_Help_House.asp' target='main'><span>" & rsHouse("ChannelName") & "管理</span></a></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td style='display:none' align='right' id='submenu502'><div class=sec_menu style='width:165'>" & vbCrLf
            Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Set rsHouseConfig = Conn.Execute("Select ClassID,ClassName from PE_HouseConfig")
            Do While Not rsHouseConfig.EOF
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=Add&ClassID=" & rsHouseConfig("ClassID") & "' target='main'>发布" & rsHouseConfig("ClassName") & "信息</a> | <a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=Manage&ClassID=" & rsHouseConfig("ClassID") & "' target='main'>管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
                rsHouseConfig.MoveNext
            Loop
            rsHouseConfig.Close
            Set rsHouseConfig = Nothing
            Response.Write "<tr><td height=10>=====================</td></tr>" & vbCrLf
            If AdminPurview = 1 Or arrPurview(2) = True Or CheckPurview_Other(AdminPurview_Others, "Area_" & ChannelDir) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=ManageArea' target='main'>区域管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or arrPurview(2) = True Or CheckPurview_Other(AdminPurview_Others, "ClassConfig_" & ChannelDir) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=ModifyConfig' target='main'>房产栏目配置</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "<tr><td height=20><a href='Admin_UploadFile.asp?ChannelID=" & rsHouse("ChannelID") & "&UploadDir=UploadPhotos' target=main>上传图片管理</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=Clear&UploadDir=UploadPhotos' target=main>清理</a></td></tr>" & vbCrLf
            If AdminPurview = 1 Or arrPurview(2) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_House.asp?ChannelID=" & rsHouse("ChannelID") & "&Action=Manage&ManageType=RecycleBin' target='main'>回收站管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or arrPurview(2) = True Or CheckPurview_Other(AdminPurview_Others, "Template_" & ChannelDir) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Template.asp?ChannelID=" & rsHouse("ChannelID") & "' target='main'>模板管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "        </table>" & vbCrLf
            Response.Write "      </div>" & vbCrLf
            Response.Write "        <div  style='width:158'>" & vbCrLf
            Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Response.Write "            <tr>" & vbCrLf
            Response.Write "              <td height=5></td>" & vbCrLf
            Response.Write "            </tr>" & vbCrLf
            Response.Write "          </table>" & vbCrLf
            Response.Write "      </div></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
        End If
    End If
    rsHouse.Close
    Set rsHouse = Nothing

    Dim rsJob
    Set rsJob = Conn.Execute("Select ChannelName, ChannelDir, Disabled from PE_Channel Where ModuleType=8")
    If FoundInArr(AllModules, "Job", ",") And rsJob("Disabled") = False Then
        ChannelDir = rsJob("ChannelDir")
        AdminPurview_Channel = rsGetAdmin("AdminPurview_" & ChannelDir)
        If AdminPurview = 1 Or AdminPurview_Channel <= 3 Then
            Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_8.gif' id=menuTitle607 onclick=""new Element.toggle('submenu607')"" style='cursor:hand;'><a href='Admin_Help_Job.asp' target='main'><span>" & rsJob("ChannelName") & "管理</span></a></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td style='display:none' align='right' id='submenu607'><div class=sec_menu style='width:165'>" & vbCrLf
            Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=Position' target='main'>职位信息管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or AdminPurview_Channel = 1 Or AdminPurview_Channel = 3 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=Resume' target='main'>人才信息管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "          <tr height=10>" & vbCrLf
            Response.Write "            <td>====================</td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            If AdminPurview = 1 Or AdminPurview_Channel <= 2 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=JobCategory' target='main'>工作类别管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=WorkPlace' target='main'>工作地点管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Job.asp?ChannelID=997&Action=SubCompany' target='main'>用人单位管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or AdminPurview_Channel = 1 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_UploadFile.asp?ChannelID=997&UploadDir=UploadPhotos' target=main>上传图片管理</a> | <a href='Admin_UploadFile_Clear.asp?ChannelID=997&Action=Clear&UploadDir=UploadPhotos' target=main>清理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='Admin_Template.asp?ChannelID=997' target='main'>招聘模板页管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "        </table>" & vbCrLf
            Response.Write "      </div>" & vbCrLf
            Response.Write "        <div  style='width:158'>" & vbCrLf
            Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Response.Write "            <tr>" & vbCrLf
            Response.Write "              <td height=5></td>" & vbCrLf
            Response.Write "            </tr>" & vbCrLf
            Response.Write "          </table>" & vbCrLf
            Response.Write "      </div></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
        End If
    End If
    rsJob.Close
    Set rsJob = Nothing
End If

Dim Purview_OrderManage
PurviewPassed = False
Purview_OrderManage = False
arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "Order_View")
arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "Order_Confirm")
arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "Order_Modify")
arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "Order_Del")
arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "Order_Payment")
arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "Order_Invoice")
arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "Order_Deliver")
arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "Order_Download")
arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "Order_SendCard")
arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "Order_End")
arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Order_Transfer")
arrPurview(11) = CheckPurview_Other(AdminPurview_Others, "Order_Print")
arrPurview(12) = CheckPurview_Other(AdminPurview_Others, "Order_Count")
arrPurview(13) = CheckPurview_Other(AdminPurview_Others, "Order_OrderItem")
arrPurview(14) = CheckPurview_Other(AdminPurview_Others, "Order_SaleCount")
arrPurview(15) = CheckPurview_Other(AdminPurview_Others, "Payment")
arrPurview(16) = CheckPurview_Other(AdminPurview_Others, "Bankroll")
arrPurview(17) = CheckPurview_Other(AdminPurview_Others, "Deliver")
arrPurview(18) = CheckPurview_Other(AdminPurview_Others, "Transfer")
arrPurview(19) = CheckPurview_Other(AdminPurview_Others, "PresentProject")
arrPurview(20) = CheckPurview_Other(AdminPurview_Others, "PaymentType")
arrPurview(21) = CheckPurview_Other(AdminPurview_Others, "DeliverType")
arrPurview(22) = CheckPurview_Other(AdminPurview_Others, "Bank")
arrPurview(23) = CheckPurview_Other(AdminPurview_Others, "Order_Refund")
For PurviewIndex = 0 To 12
    If arrPurview(PurviewIndex) = True Or arrPurview(23) = True Then
        PurviewPassed = True
        Purview_OrderManage = True
        Exit For
    End If
Next
For PurviewIndex = 13 To 22
    If arrPurview(PurviewIndex) = True Then
        PurviewPassed = True
        Exit For
    End If
Next
If AdminPurview = 1 Or PurviewPassed = True Then
    Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_10.gif' id=menuTitle901 onclick=""new Element.toggle('submenu901')"" style='cursor:hand;'><a href='admin_help_Shop.asp' target='main'><span class=glow>商城日常操作</span></a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td style='display:none' align='right' id='submenu901'><div class=sec_menu style='width:165'>" & vbCrLf
    Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    If ShowAdmin_Shop = True And eShop_Edition >= 0 Then
        If AdminPurview = 1 Or Purview_OrderManage = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Order.asp?SearchType=1' target='main'>处理今天的订单</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Order.asp' target='main'>处理所有订单</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr height=10>" & vbCrLf
            Response.Write "            <td>====================</td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(13) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_OrderItem.asp' target='main'>销售明细情况</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(14) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_SaleCount.asp' target='main'>销售统计/排行</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
    End If
    If AdminPurview = 1 Or arrPurview(15) = True Then
        Response.Write "          <tr height=20>" & vbCrLf
        Response.Write "            <td><a href='Admin_Payment.asp' target='main'>在线支付记录管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(16) = True Then
        Response.Write "          <tr height=20>" & vbCrLf
        Response.Write "            <td><a href='Admin_Bankroll.asp' target='main'>资金明细查询</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If ShowAdmin_Shop = True And eShop_Edition >= 0 Then
        If AdminPurview = 1 Or arrPurview(5) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Invoice.asp' target='main'>开发票记录</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(17) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Deliver.asp' target='main'>发退货记录</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(18) = True Then
            Response.Write "          <tr height=20>" & vbCrLf
            Response.Write "            <td><a href='Admin_Transfer.asp' target='main'>订单过户记录</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(19) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_PresentProject.asp' target='main'>促销方案管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(20) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_PaymentType.asp' target='main'>付款方式管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(21) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_DeliverType.asp' target='main'>送货方式管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
    End If
    If AdminPurview = 1 Or arrPurview(22) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Bank.asp' target='main'>银行帐户管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </div>" & vbCrLf
    Response.Write "        <div  style='width:167'>" & vbCrLf
    Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td height=5></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table>" & vbCrLf
    Response.Write "      </div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End If

If FoundInArr(AllModules, "CRM", ",") And CRM_Edition > 0 Then
    Dim PurviewPassed_Client, PurviewPassed_Service, PurviewPassed_Complain, PurviewPassed_Call
    PurviewPassed = False
    arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "Client_View")
    arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "Client_Add")
    arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "Client_ModifyOwn")
    arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "Client_ModifyAll")
    arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "Client_Del")
    arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "Service_View")
    arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "Service_Add")
    arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "Service_ModifyOwn")
    arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "Service_ModifyAll")
    arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "Service_Del")
    arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Complain_View")
    arrPurview(11) = CheckPurview_Other(AdminPurview_Others, "Complain_Add")
    arrPurview(12) = CheckPurview_Other(AdminPurview_Others, "Complain_ModifyOwn")
    arrPurview(13) = CheckPurview_Other(AdminPurview_Others, "Complain_ModifyAll")
    arrPurview(14) = CheckPurview_Other(AdminPurview_Others, "Complain_Del")
    arrPurview(15) = CheckPurview_Other(AdminPurview_Others, "Call_View")
    arrPurview(16) = CheckPurview_Other(AdminPurview_Others, "Call_Add")
    arrPurview(17) = CheckPurview_Other(AdminPurview_Others, "Call_ModifyOwn")
    arrPurview(18) = CheckPurview_Other(AdminPurview_Others, "Call_ModifyAll")
    arrPurview(19) = CheckPurview_Other(AdminPurview_Others, "Dictionary")
    For PurviewIndex = 0 To 4
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            PurviewPassed_Client = True
            Exit For
        End If
    Next
    For PurviewIndex = 5 To 9
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            PurviewPassed_Service = True
            Exit For
        End If
    Next
    For PurviewIndex = 10 To 14
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            PurviewPassed_Complain = True
            Exit For
        End If
    Next
    For PurviewIndex = 15 To 18
        If arrPurview(PurviewIndex) = True Then
            PurviewPassed = True
            PurviewPassed_Call = True
            Exit For
        End If
    Next

    If AdminPurview = 1 Or PurviewPassed = True Or arrPurview(19) = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_02.gif' id=menuTitle204 onclick=""new Element.toggle('submenu204')"" style='cursor:hand;'><a href='Admin_Help_CRM.asp' target='main'><span class=glow>客户关系管理</span></a></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu204'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        If AdminPurview = 1 Or PurviewPassed_Client = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href=Admin_Client.asp target=main>客户管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href=Admin_Contacter.asp target=main>联系人管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or PurviewPassed_Service = True Or PurviewPassed_Call = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href=Admin_Service.asp target=main>服务管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or PurviewPassed_Complain = True Or PurviewPassed_Call = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Complain.asp' target='main'>投诉管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or arrPurview(19) = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href=Admin_Dictionary.asp target=main>数据字典管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

If FoundInArr(AllModules, "Collection", ",") Then
    If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Collection") = True Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_05.gif' id=menuTitle210 onclick=""new Element.toggle('submenu210')"" style='cursor:hand;'><a href='Admin_Help_Collection.asp' target='main'><span class=glow>采集管理</span></a></td></tr><tr><td style='display:none' align='right' id='submenu210'><div class=sec_menu style='width:165'><table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Collection.asp?Action=Main target=main>文章采集</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_CollectionManage.asp?Action=ItemManage target=main>项目管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Filter.asp?Action=main target=main>过滤管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_CollectionHistory.asp?Action=main target=main>采集历史记录</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_CollectionManage.asp?Action=Import target=main>项目导入</a> | <a href=Admin_CollectionManage.asp?Action=Export target=main>项目导出</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Timing.asp?Action=main target=main>定时设置</a> | <a href=Admin_Timing.asp?Action=DoMainTiming target='_blank'>启动定时</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_AreaCollection.asp?Action=Main target=main>区域采集管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

If (AdminPurview = 1 And (FileExt_SiteIndex < 4 Or FileExt_SiteSpecial < 4)) Or strActionLink <> "" Then
    Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_9.gif' id=menuTitle301 onclick=""new Element.toggle('submenu301')"" style='cursor:hand;'><a href='Admin_Help_Create.asp' target='main'><span class=glow>网站生成管理</span></a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td style='display:none' align='right' id='submenu301'><div class=sec_menu style='width:165'>" & vbCrLf
    Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    If AdminPurview = 1 And FileExt_SiteIndex < 4 Then
        Response.Write "          <tr height=20>" & vbCrLf
        Response.Write "            <td><a href='Admin_CreateIndex.asp' target='main'>生成网站首页</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr height=20>" & vbCrLf
        Response.Write "            <td><a href='Admin_CreateOther.asp' target='main'>生成网站综合数据</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 And FileExt_SiteSpecial < 4 Then
        Response.Write "          <tr height=20>" & vbCrLf
        Response.Write "            <td><a href='Admin_CreateHTML.asp?Action=SiteSpecial' target='main'>全站专题生成管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    Response.Write strActionLink & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Timing.asp?Action=main target=main>定时设置</a> | <a href=Admin_Timing.asp?Action=DoMainTiming target='_blank'>启动定时</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </div>" & vbCrLf
    Response.Write "        <div  style='width:167'>" & vbCrLf
    Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td height=5></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table>" & vbCrLf
    Response.Write "      </div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End If

Dim PurviewPassed_User
PurviewPassed = False
PurviewPassed_User = False

arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "User_View")
arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "User_ModifyInfo")
arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "User_MofidyPurview")
arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "User_Lock")
arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "User_Del")
arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "User_Update")
arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "User_Money")
arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "User_Point")
arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "User_Valid")
arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "UserGroup")
arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Card")
arrPurview(11) = CheckPurview_Other(AdminPurview_Others, "ConsumeLog")
arrPurview(12) = CheckPurview_Other(AdminPurview_Others, "RechargeLog")
arrPurview(13) = CheckPurview_Other(AdminPurview_Others, "Message")
arrPurview(14) = CheckPurview_Other(AdminPurview_Others, "MailList")
For PurviewIndex = 0 To 8
    If arrPurview(PurviewIndex) = True Then
        PurviewPassed = True
        PurviewPassed_User = True
        Exit For
    End If
Next
For PurviewIndex = 9 To 13
    If arrPurview(PurviewIndex) = True Then
        PurviewPassed = True
        Exit For
    End If
Next
If AdminPurview = 1 Or PurviewPassed = True Then
    Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/admin_left_11.gif' id=menuTitle203 onclick=""new Element.toggle('submenu203')"" style='cursor:hand;'><a href='Admin_Help_User.asp' target='main'><span class=glow>用户管理</span></a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td style='display:none' align='right' id='submenu203'><div class=sec_menu style='width:165'>" & vbCrLf
    Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    If AdminPurview = 1 Or PurviewPassed_User = True = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_User.asp target=main>会员管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(9) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_UserGroup.asp' target='main'>会员组管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(10) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Card.asp' target='main'>充值卡管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(11) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_ConsumeLog.asp' target='main'>会员" & PointName & "明细</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(12) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_RechargeLog.asp' target='main'>会员有效期明细</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(13) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Message.asp' target='main'>短消息管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(14) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Maillist.asp' target='main'>邮件列表管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or PurviewPassed_User = True = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SpaceManage.asp' target='main'>聚合空间管理</a> | <a href='Admin_SpaceManage.asp?Action=Check' target='main'>审核</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SpaceManage.asp?Action=Kind' target='main'>空间分类管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Admin.asp target=main>管理员管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Template.asp?TemplateType=8&TempType=1 target=main>会员模板页管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </div>" & vbCrLf
    Response.Write "        <div  style='width:167'>" & vbCrLf
    Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td height=5></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table>" & vbCrLf
    Response.Write "      </div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End If


If ShowAdmin_Guest = True Then
    'PurviewPassed = CheckPurview_Other(AdminPurview_Others, "GuestBook")
    If AdminPurview = 1 Or AdminPurview_GuestBook <= 3 Then
        Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_4.gif' id=menuTitle202 onclick=""new Element.toggle('submenu202')"" style='cursor:hand;'><a href='Admin_Help_Guest.asp' target='main'><span class=glow>留言板管理</span></a></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "  <tr>" & vbCrLf
        Response.Write "    <td style='display:none' align='right' id='submenu202'><div class=sec_menu style='width:165'>" & vbCrLf
        Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_GuestBook.asp?Passed=False' target=main>网站留言审核</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_GuestBook.asp?Passed=All' target=main>网站留言管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        If AdminPurview = 1 Or AdminPurview_GuestBook < 3 Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_GuestBook.asp?Action=GKind' target=main>留言类别管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If

        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Template") = True Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Template.asp?ChannelID=4' target=main>留言模板页管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Or AdminPurview_GuestBook < 2 Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_GuestBook.asp?Action=CreateCode' target=main>首页嵌入代码生成</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        If AdminPurview = 1 Then
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_CreateXml.asp?Action=GreatNav&ChannelID=4' target=main>更新栏目XML数据</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
        End If
        Response.Write "        </table>" & vbCrLf
        Response.Write "      </div>" & vbCrLf
        Response.Write "        <div  style='width:167'>" & vbCrLf
        Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
        Response.Write "            <tr>" & vbCrLf
        Response.Write "              <td height=5></td>" & vbCrLf
        Response.Write "            </tr>" & vbCrLf
        Response.Write "          </table>" & vbCrLf
        Response.Write "      </div></td>" & vbCrLf
        Response.Write "  </tr>" & vbCrLf
        Response.Write "</table>" & vbCrLf
    End If
End If

If CMS_Edition > 0 Or eShop_Edition > 0 Or CRM_Edition > 0 Then

    If FoundInArr(AllModules, "Classroom", ",") Then
        If AdminPurview = 1 Or CheckPurview_Other(AdminPurview_Others, "Equipment") = True Then
            Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/admin_left_13.gif' id=menuTitle209 onclick=""new Element.toggle('submenu209')"" style='cursor:hand;'><a href='Admin_Help_Classroom.asp' target='main'><span>室场登记管理</span></a></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td style='display:none' align='right' id='submenu209'><div class=sec_menu style='width:165'>" & vbCrLf
            Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_ClassroomSort.asp' target='main'>设备管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_Equipment.asp' target='main'>设备管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_ManageRecord.asp?Action=AdminEnrol' target='main'>固定排课</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr>" & vbCrLf
            
            Response.Write "            <td height=20><a href='Admin_ManageRecord.asp?Status=All' target='main'>使用登记管理</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "          <tr>" & vbCrLf
            Response.Write "          <tr>" & vbCrLf
            Response.Write "            <td height=20><a href='Admin_SearchHistory.asp' target='main'>使用记录查询</a></td>" & vbCrLf
            Response.Write "          </tr>" & vbCrLf
            Response.Write "        </table>" & vbCrLf
            Response.Write "      </div>" & vbCrLf
            Response.Write "        <div  style='width:167'>" & vbCrLf
            Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Response.Write "            <tr>" & vbCrLf
            Response.Write "              <td height=5></td>" & vbCrLf
            Response.Write "            </tr>" & vbCrLf
            Response.Write "          </table>" & vbCrLf
            Response.Write "      </div></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
        End If
    End If

    If FoundInArr(AllModules, "Sdms", ",") Then
        PurviewPassed = False
        arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "InfoManage")
        arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "ScoreManage")
        arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "TestManage")
        For PurviewIndex = 0 To 2
            If arrPurview(PurviewIndex) = True Then
                PurviewPassed = True
                Exit For
            End If
        Next
        If AdminPurview = 1 Or PurviewPassed = True Then
            Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_12.gif' id=menuTitle402 onclick=""new Element.toggle('submenu402')"" style='cursor:hand;'><a href='Admin_Help_Manage.asp' target='main'><span>学生学籍管理</span></a></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "  <tr>" & vbCrLf
            Response.Write "    <td style='display:none' align='right' id='submenu402'><div class=sec_menu style='width:165'>" & vbCrLf
            Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            If AdminPurview = 1 Or arrPurview(0) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='../sdms/InfoManage.asp' target='main'>学生信息管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or arrPurview(1) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='../sdms/ScoreManage.asp' target='main'>学生成绩管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Or arrPurview(2) = True Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='../sdms/TestManage.asp' target='main'>考试管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            If AdminPurview = 1 Then
                Response.Write "          <tr>" & vbCrLf
                Response.Write "            <td height=20><a href='../sdms/DatabaseManage.asp' target='main'>学籍数据库管理</a></td>" & vbCrLf
                Response.Write "          </tr>" & vbCrLf
            End If
            Response.Write "        </table>" & vbCrLf
            Response.Write "      </div>" & vbCrLf
            Response.Write "        <div  style='width:167'>" & vbCrLf
            Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
            Response.Write "            <tr>" & vbCrLf
            Response.Write "              <td height=5></td>" & vbCrLf
            Response.Write "            </tr>" & vbCrLf
            Response.Write "          </table>" & vbCrLf
            Response.Write "      </div></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            Response.Write "</table>" & vbCrLf
        End If
    End If
End If

PurviewPassed = False
arrPurview(0) = CheckPurview_Other(AdminPurview_Others, "Channel")
arrPurview(1) = CheckPurview_Other(AdminPurview_Others, "AD")
arrPurview(2) = CheckPurview_Other(AdminPurview_Others, "FriendSite")
arrPurview(3) = CheckPurview_Other(AdminPurview_Others, "Announce")
arrPurview(4) = CheckPurview_Other(AdminPurview_Others, "Vote")
arrPurview(5) = CheckPurview_Other(AdminPurview_Others, "Counter")
arrPurview(6) = CheckPurview_Other(AdminPurview_Others, "Skin")
arrPurview(7) = CheckPurview_Other(AdminPurview_Others, "Label")
arrPurview(8) = CheckPurview_Other(AdminPurview_Others, "KeyLink")
arrPurview(9) = CheckPurview_Other(AdminPurview_Others, "Rtext")
arrPurview(10) = CheckPurview_Other(AdminPurview_Others, "Template")
For PurviewIndex = 0 To 10
    If arrPurview(PurviewIndex) = True Then
        PurviewPassed = True
        Exit For
    End If
Next
If AdminPurview = 1 Or PurviewPassed = True Then
    Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_01.gif' id=menuTitle201 onclick=""new Element.toggle('submenu201')"" style='cursor:hand;'><a href='Admin_Help_SiteConfig.asp' target='main'><span class=glow>系统设置</span></a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td style='display:none' align='right' id='submenu201'><div class=sec_menu style='width:165'>" & vbCrLf
    Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_SiteConfig.asp target=main>网站信息配置</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(0) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Channel.asp target=main>网站频道管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Special.asp target=main>全站专题管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(1) = True) And FoundInArr(AllModules, "Advertisement", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Advertisement.asp target=main>网站广告管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(2) = True) And FoundInArr(AllModules, "FriendSite", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_FriendSite.asp target=main>友情链接管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(3) = True) And FoundInArr(AllModules, "Announce", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Announce.asp target=main>网站公告管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(4) = True) And FoundInArr(AllModules, "Vote", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Vote.asp target=main>网站调查管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(5) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Counter.asp target=main>网站统计分析</a> | <a href=Admin_Counter.asp?Action=ShowConfig target=main>配置</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(6) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href=Admin_Skin.asp target=main>网站风格管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(10) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Template.asp?ChannelID=0' target='main'>网站通用模板页管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_TemplateProject.asp?ChannelID=0' target='main'>网站模板方案管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SourceManage.asp?ChannelID=0&TypeSelect=Keyword' target='main'>网站通用关键字管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Or arrPurview(7) = True Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Label.asp' target='main'>自定义标签管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Page.asp' target='main'>自定义页面管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_Log.asp' target='main'>网站日志管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(8) = True) And FoundInArr(AllModules, "KeyLink", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SourceManage.asp?TypeSelect=KeyLink' target='main'>站内链接管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If (AdminPurview = 1 Or arrPurview(9) = True) And FoundInArr(AllModules, "Rtext", ",") Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_SourceManage.asp?TypeSelect=Rtext' target='main'>字符过滤管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    If AdminPurview = 1 Then
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_CompareFilesOnline.asp' target='main'>在线比较网站文件</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
        Response.Write "          <tr>" & vbCrLf
        Response.Write "            <td height=20><a href='Admin_City.asp' target='main'>邮政编码管理</a></td>" & vbCrLf
        Response.Write "          </tr>" & vbCrLf
    End If
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </div>" & vbCrLf
    Response.Write "        <div style='width:167'>" & vbCrLf
    Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td height=5></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table>" & vbCrLf
    Response.Write "      </div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End If


If AdminPurview = 1 Then
    Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_03.gif' id=menuTitle206 onclick=""new Element.toggle('submenu206')"" style='cursor:hand;'><a href='Admin_Help_Database.asp' target='main'><span class=glow>数据库管理</span></a></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "  <tr>" & vbCrLf
    Response.Write "    <td style='display:none' align='right' id='submenu206'><div class=sec_menu style='width:165'>" & vbCrLf
    Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=Backup target=main>备份数据库</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=Restore target=main>恢复数据库</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=Compact target=main>压缩数据库</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=Init target=main>系统初始化</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "          <tr>" & vbCrLf
    Response.Write "            <td height=20><a href=Admin_Database.asp?Action=SpaceSize target=main>系统空间占用</a></td>" & vbCrLf
    Response.Write "          </tr>" & vbCrLf
    Response.Write "        </table>" & vbCrLf
    Response.Write "      </div>" & vbCrLf
    Response.Write "        <div  style='width:167'>" & vbCrLf
    Response.Write "          <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
    Response.Write "            <tr>" & vbCrLf
    Response.Write "              <td height=5></td>" & vbCrLf
    Response.Write "            </tr>" & vbCrLf
    Response.Write "          </table>" & vbCrLf
    Response.Write "      </div></td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
End If

Response.Write "<table cellpadding=0 cellspacing=0 width=167 align=center>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td height=28 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background='Images/Admin_left_04.gif' id=menuTitle208><span>系统信息</span> </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align='right'><div class=sec_menu style='width:165'>" & vbCrLf
Response.Write "        <table cellpadding=0 cellspacing=0 align=center width=130>" & vbCrLf
Response.Write "          <tr>" & vbCrLf
Response.Write "            <td height=20><br>" & vbCrLf
Response.Write "              版权所有：&nbsp;<a href='http://www.powereasy.net/' target=_blank>动易网络</a><br>" & vbCrLf
Response.Write "              设计制作：&nbsp;<a href='http://www.powereasy.net/' target=_blank>动易网络</a><BR>" & vbCrLf
Response.Write "              技术支持：&nbsp;<a href='http://bbs.powereasy.net/' target=_blank>动易论坛</a><br>" & vbCrLf
Response.Write "              --------------------<br>" & vbCrLf
Response.Write "              联 系 QQ：&nbsp;137070625<br>" & vbCrLf
Response.Write "              <br>" & vbCrLf
Response.Write "            </td>" & vbCrLf
Response.Write "          </tr>" & vbCrLf
Response.Write "        </table>" & vbCrLf
Response.Write "    </div></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
Response.Write "</body>" & vbCrLf
Response.Write "</html>" & vbCrLf

rsGetAdmin.Close
Set rsGetAdmin = Nothing
Call CloseConn
%>
