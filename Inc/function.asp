<%
Dim Action, FoundErr, ErrMsg, ComeUrl
Dim strInstallDir
Dim Site_Sn   '����ϵͳʶ����
'�����û���صı���
Dim UserLogined, GroupID, GroupName, GroupType, Discount_Member, IsOffer, LoginTimes, RegTime, JoinTime, LastLoginTime, LastLoginIP
Dim UserID, ClientID, CompanyID, ContacterID, UserType, UserName, email, Balance, UserPoint, UserExp, ValidNum, ValidDays, SpecialPermission, UserSetting, ChargeType
Dim UnsignedItems, UnreadMsg, arrClass_Input, arrClass_View
Dim DefaultTemplateProjectName

If Request("ComeUrl") = "" Then
    ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
Else
    ComeUrl = Trim(Request("ComeUrl"))
End If
Action = Trim(Request("Action"))
FoundErr = False
ErrMsg = ""
If Right(InstallDir, 1) <> "/" Then
    strInstallDir = InstallDir & "/"
Else
    strInstallDir = InstallDir
End If
Site_Sn = Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME") & InstallDir), "/", ""), ".", "")



'**************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ���
'**************************************************
Function gotTopic(ByVal str, ByVal strlen)
    If str = "" Then
        gotTopic = ""
        Exit Function
    End If
    Dim l, t, c, i, strTemp
    str = Replace(Replace(Replace(Replace(str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
    l = Len(str)
    t = 0
    strTemp = str
    strlen = CLng(strlen)
    For i = 1 To l
        c = Abs(Asc(Mid(str, i, 1)))
        If c > 255 Then
            t = t + 2
        Else
            t = t + 1
        End If
        If t >= strlen Then
            strTemp = Left(str, i)
            Exit For
        End If
    Next
    If strTemp <> str Then
        strTemp = strTemp & "��"
    End If
    gotTopic = Replace(Replace(Replace(Replace(strTemp, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
End Function

'**************************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'**************************************************
Function JoinChar(ByVal strUrl)
    If strUrl = "" Then
        JoinChar = ""
        Exit Function
    End If
    If InStr(strUrl, "?") < Len(strUrl) Then
        If InStr(strUrl, "?") > 1 Then
            If InStr(strUrl, "&") < Len(strUrl) Then
                JoinChar = strUrl & "&"
            Else
                JoinChar = strUrl
            End If
        Else
            JoinChar = strUrl & "?"
        End If
    Else
        JoinChar = strUrl
    End If
End Function

'**************************************************
'��������ShowPage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sFileName  ----���ӵ�ַ
'        TotalNumber ----������
'        MaxPerPage  ----ÿҳ����
'        CurrentPage ----��ǰҳ
'        ShowTotal   ----�Ƿ���ʾ������
'        ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת��
'        strUnit     ----������λ
'        ShowMaxPerPage  ----�Ƿ���ʾÿҳ��Ϣ��ѡ���
'����ֵ������һҳ ��һҳ������Ϣ��HTML����
'**************************************************
Function ShowPage(sfilename, totalnumber, MaxPerPage, CurrentPage, ShowTotal, ShowAllPages, strUnit, ShowMaxPerPage)
    Dim TotalPage, strTemp, strUrl, i

    If totalnumber = 0 Or MaxPerPage = 0 Or IsNull(MaxPerPage) Then
        ShowPage = ""
        Exit Function
    End If
    If totalnumber Mod MaxPerPage = 0 Then
        TotalPage = totalnumber \ MaxPerPage
    Else
        TotalPage = totalnumber \ MaxPerPage + 1
    End If
    If CurrentPage > TotalPage Then CurrentPage = TotalPage
        
    strTemp = "<table align='center'><tr><td>"
    If ShowTotal = True Then
        strTemp = strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
    End If
    If ShowMaxPerPage = True Then
        strUrl = JoinChar(sfilename) & "MaxPerPage=" & MaxPerPage & "&"
    Else
        strUrl = JoinChar(sfilename)
    End If
    If CurrentPage = 1 Then
        strTemp = strTemp & "��ҳ ��һҳ&nbsp;"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage - 1) & "'>��һҳ</a>&nbsp;"
    End If

    If CurrentPage >= TotalPage Then
        strTemp = strTemp & "��һҳ βҳ"
    Else
        strTemp = strTemp & "<a href='" & strUrl & "page=" & (CurrentPage + 1) & "'>��һҳ</a>&nbsp;"
        strTemp = strTemp & "<a href='" & strUrl & "page=" & TotalPage & "'>βҳ</a>"
    End If
    strTemp = strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & TotalPage & "</strong>ҳ "
    If ShowMaxPerPage = True Then
        strTemp = strTemp & "&nbsp;<input type='text' name='MaxPerPage' size='3' maxlength='4' value='" & MaxPerPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & JoinChar(sfilename) & "page=" & CurrentPage & "&MaxPerPage=" & "'+this.value;"">" & strUnit & "/ҳ"
    Else
        strTemp = strTemp & "&nbsp;<b>" & MaxPerPage & "</b>" & strUnit & "/ҳ"
    End If
    If ShowAllPages = True Then
        If TotalPage > 20 Then
            strTemp = strTemp & "&nbsp;&nbsp;ת����<input type='text' name='page' size='3' maxlength='5' value='" & CurrentPage & "' onKeyPress=""if (event.keyCode==13) window.location='" & strUrl & "page=" & "'+this.value;"">ҳ"
        Else
            strTemp = strTemp & "&nbsp;ת����<select name='page' size='1' onchange=""javascript:window.location='" & strUrl & "page=" & "'+this.options[this.selectedIndex].value;"">"
            For i = 1 To TotalPage
               strTemp = strTemp & "<option value='" & i & "'"
               If PE_CLng(CurrentPage) = PE_CLng(i) Then strTemp = strTemp & " selected "
               strTemp = strTemp & ">��" & i & "ҳ</option>"
            Next
            strTemp = strTemp & "</select>"
        End If
    End If
    strTemp = strTemp & "</td></tr></table>"
    ShowPage = strTemp
End Function

'**************************************************
'��������IsObjInstalled
'��  �ã��������Ƿ��Ѿ���װ
'��  ����strClassString ----�����
'����ֵ��True  ----�Ѿ���װ
'       False ----û�а�װ
'**************************************************
Function IsObjInstalled(strClassString)
    On Error Resume Next
    IsObjInstalled = False
    Err = 0
    Dim xTestObj
    Set xTestObj = Server.CreateObject(strClassString)
    If 0 = Err Then IsObjInstalled = True
    Set xTestObj = Nothing
    Err = 0
End Function

'**************************************************
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
Function strLength(str)
    On Error Resume Next
    Dim WINNT_CHINESE
    WINNT_CHINESE = (Len("�й�") = 2)
    If WINNT_CHINESE Then
        Dim l, t, c
        Dim i
        l = Len(str)
        t = l
        For i = 1 To l
            c = Asc(Mid(str, i, 1))
            If c < 0 Then c = c + 65536
            If c > 255 Then
                t = t + 1
            End If
        Next
        strLength = t
    Else
        strLength = Len(str)
    End If
    If Err.Number <> 0 Then Err.Clear
End Function


'**************************************************
'������FoundInArr
'��  �ã����һ������������Ԫ���Ƿ����ָ���ַ���
'��  ����strArr     ----�洢�������ݵ��ִ�
'       strToFind    ----Ҫ���ҵ��ַ���
'       strSplit    ----����ķָ���
'����ֵ��True,False
'**************************************************
Function FoundInArr(strArr, strToFind, strSplit)
    Dim arrTemp, i
    FoundInArr = False
    If InStr(strArr, strSplit) > 0 Then
        arrTemp = Split(strArr, strSplit)
        For i = 0 To UBound(arrTemp)
        If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
            FoundInArr = True
            Exit For
        End If
        Next
    Else
        If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then
        FoundInArr = True
        End If
    End If
End Function

'**************************************************
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'**************************************************
Sub WriteErrMsg(sErrMsg, sComeUrl)
    Dim strErr
    strErr = strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    strErr = strErr & "<link href='" & strInstallDir & AdminDir & "/Admin_Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
    strErr = strErr & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    strErr = strErr & "  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>" & vbCrLf
    strErr = strErr & "  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b>" & sErrMsg & "</td></tr>" & vbCrLf
    strErr = strErr & "  <tr align='center' class='tdbg'><td>"
    If sComeUrl <> "" Then
        strErr = strErr & "<a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a>"
    Else
        strErr = strErr & "<a href='javascript:window.close();'>���رա�</a>"
    End If
    strErr = strErr & "</td></tr>" & vbCrLf
    strErr = strErr & "</table>" & vbCrLf
    strErr = strErr & "</body></html>" & vbCrLf
    Response.Write strErr
End Sub

'**************************************************
'��������WriteSuccessMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'**************************************************
Sub WriteSuccessMsg(sSuccessMsg, sComeUrl)
    Dim strSuccess
    strSuccess = strSuccess & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    strSuccess = strSuccess & "<link href='" & strInstallDir & AdminDir & "/Admin_Style.css' rel='stylesheet' type='text/css'></head><body><br><br>" & vbCrLf
    strSuccess = strSuccess & "<table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>" & vbCrLf
    strSuccess = strSuccess & "  <tr align='center' class='title'><td height='22'><strong>��ϲ�㣡</strong></td></tr>" & vbCrLf
    strSuccess = strSuccess & "  <tr class='tdbg'><td height='100' valign='top'><br>" & sSuccessMsg & "</td></tr>" & vbCrLf
    strSuccess = strSuccess & "  <tr align='center' class='tdbg'><td>"
    If sComeUrl <> "" Then
        strSuccess = strSuccess & "<a href='" & sComeUrl & "'>&lt;&lt; ������һҳ</a>"
    Else
        strSuccess = strSuccess & "<a href='javascript:window.close();'>���رա�</a>"
    End If
    strSuccess = strSuccess & "</td></tr>" & vbCrLf
    strSuccess = strSuccess & "</table>" & vbCrLf
    strSuccess = strSuccess & "</body></html>" & vbCrLf
    Response.Write strSuccess
End Sub

'**************************************************
'��������ReplaceBadChar
'��  �ã����˷Ƿ���SQL�ַ�
'��  ����strChar-----Ҫ���˵��ַ�
'����ֵ�����˺���ַ�
'**************************************************
Function ReplaceBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceBadChar = tempChar
End Function

Function PE_CLng(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng = CLng(str1)
    Else
        PE_CLng = 0
    End If
End Function

Function PE_CDbl(ByVal str1)
    If IsNumeric(str1) Then
        PE_CDbl = CDbl(str1)
    Else
        PE_CDbl = 0
    End If
End Function

Function PE_CDate(ByVal str1)
    If IsDate(str1) Then
        PE_CDate = CDate(str1)
    Else
        PE_CDate = Date
    End If
End Function

'**************************************************
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
'**************************************************
Function IsValidEmail(email)
    Dim names, name, i, c
    IsValidEmail = True
    names = Split(email, "@")
    If UBound(names) <> 1 Then
       IsValidEmail = False
       Exit Function
    End If
    For Each name In names
        If Len(name) <= 0 Then
        IsValidEmail = False
        Exit Function
        End If
        For i = 1 To Len(name)
        c = LCase(Mid(name, i, 1))
        If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
           IsValidEmail = False
           Exit Function
         End If
       Next
       If Left(name, 1) = "." Or Right(name, 1) = "." Then
          IsValidEmail = False
          Exit Function
       End If
    Next
    If InStr(names(1), ".") <= 0 Then
        IsValidEmail = False
       Exit Function
    End If
    i = Len(names(1)) - InStrRev(names(1), ".")
    If i <> 2 And i <> 3 And i <> 4 Then
       IsValidEmail = False
       Exit Function
    End If
    If InStr(email, "..") > 0 Then
       IsValidEmail = False
    End If
End Function



'�õ�������ĳ��Ԫ�ص�ֵ
Public Function GetArrItem(ByVal arrTemp, ByVal ItemIndex)
    If Not IsArray(arrTemp) Then
        GetArrItem = ""
        Exit Function
    End If
    ItemIndex = PE_CLng(ItemIndex)
    If ItemIndex < 0 Or ItemIndex > UBound(arrTemp) Then
        GetArrItem = ""
        Exit Function
    End If
    Dim strTemp
    strTemp = arrTemp(ItemIndex)
    If InStr(strTemp, "|") > 0 Then
        GetArrItem = Left(strTemp, InStr(strTemp, "|") - 1)
    Else
        GetArrItem = strTemp
    End If
End Function

'�������������б���Ŀ
Public Function Array2Option(ByVal arrTemp, ByVal ID)
    Dim strOption, i, arrValue
    strOption = "<option value='-1'> </option>"
    ID = PE_CLng(ID)
    For i = 0 To UBound(arrTemp)
        arrValue = Split(arrTemp(i), "|")
        If CLng(arrValue(1)) = 1 Then
            If ID > -1 Then
                If i = ID Then
                    strOption = strOption & "<option value='" & i & "' selected>" & arrValue(0) & "</option>"
                Else
                    strOption = strOption & "<option value='" & i & "'>" & arrValue(0) & "</option>"
                End If
            Else
                If CLng(arrValue(2)) = 1 Then
                    strOption = strOption & "<option value='" & i & "' selected>" & arrValue(0) & "</option>"
                Else
                    strOption = strOption & "<option value='" & i & "'>" & arrValue(0) & "</option>"
                End If
            End If
        End If
    Next
    Array2Option = strOption
End Function

Function GetRndPassword(PasswordLen)
    Dim Ran, i, strPassword
    strPassword = ""
    For i = 1 To PasswordLen
        Randomize
        Ran = CInt(Rnd * 2)
        Randomize
        If Ran = 0 Then
            Ran = CInt(Rnd * 25) + 97
            strPassword = strPassword & UCase(Chr(Ran))
        ElseIf Ran = 1 Then
            Ran = CInt(Rnd * 9)
            strPassword = strPassword & Ran
        ElseIf Ran = 2 Then
            Ran = CInt(Rnd * 25) + 97
            strPassword = strPassword & Chr(Ran)
        End If
    Next
    GetRndPassword = strPassword
End Function

Function GetScriptPath(ByVal ScriptName, ParentLevel)
    Dim i
    GetScriptPath = "/"
    If ScriptName = "" Or IsNull(ScriptName) Then Exit Function
    If ParentLevel > 1 Then ParentLevel = 1
    If ParentLevel = 0 Then
        GetScriptPath = Left(ScriptName, InStrRev(ScriptName, "/"))
    ElseIf ParentLevel = 1 Then
        i = InStrRev(ScriptName, "/") - 1
        If i < 1 Then i = 1
        GetScriptPath = Left(ScriptName, InStrRev(ScriptName, "/", i))
    End If
    If Right(GetScriptPath, 1) <> "/" Then GetScriptPath = GetScriptPath & "/"
End Function

'�жϵ�ǰ�������Ƿ��Ѿ���¼�����ѵ�¼�����ȡ���ݲ�����Ҫ��ֵ
Function CheckUserLogined()
    Dim UserPassword, LastPassword
    Dim rsUser, sqlUser
    Dim rsConfig

    UserName = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserName")))
    UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
    LastPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("LastPassword")))
    UserID = 0
    ClientID = 0
    CompanyID = 0
    ContacterID = 0
    UserType = 0
    GroupID = 0
    GroupType = 0
    GroupName = "�ο�"
    Discount_Member = 100
    Balance = 0
    UserPoint = 0
    UserExp = 0
    IsOffer = "��"
    
    If (UserName = "" Or UserPassword = "" Or LastPassword = "") Then
        CheckUserLogined = False
        Exit Function
    End If

    sqlUser = "SELECT U.*,G.GroupName,G.GroupType,G.GroupSetting,G.arrClass_Input as G_arrClass_Input,G.arrClass_View as G_arrClass_View FROM PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID WHERE"
    sqlUser = sqlUser & " U.UserName='" & UserName & "' AND U.UserPassword='" & UserPassword & "' AND U.LastPassword='" & LastPassword & "' and IsLocked=" & PE_False & ""
    Set rsUser = Conn.Execute(sqlUser)
    If rsUser.EOF And rsUser.BOF Then
        CheckUserLogined = False
    Else
        CheckUserLogined = True
        UserID = rsUser("UserID")
        ClientID = rsUser("ClientID")
        CompanyID = rsUser("CompanyID")
        ContacterID = rsUser("ContacterID")
        UserType = rsUser("UserType")
        UserName = rsUser("UserName")
        UserPassword = rsUser("UserPassword")
        LastPassword = rsUser("LastPassword")
        email = rsUser("Email")
        Balance = PE_CDbl(rsUser("Balance"))
        UserPoint = PE_CLng(rsUser("UserPoint"))
        UserExp = PE_CLng(rsUser("UserExp"))
        ValidNum = rsUser("ValidNum")
        LoginTimes = rsUser("LoginTimes")
        ValidDays = ChkValidDays(rsUser("ValidNum"), rsUser("ValidUnit"), rsUser("BeginTime"))
        GroupID = rsUser("GroupID")
        GroupName = rsUser("GroupName")
        GroupType = rsUser("GroupType")
        SpecialPermission = rsUser("SpecialPermission")
        If SpecialPermission = True Then
            UserSetting = Split(rsUser("UserSetting") & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
            arrClass_Input = rsUser("arrClass_Input")
            arrClass_View = rsUser("arrClass_View")
        Else
            UserSetting = Split(rsUser("GroupSetting") & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
            arrClass_Input = rsUser("G_arrClass_Input")
            arrClass_View = rsUser("G_arrClass_View")
        End If
        Discount_Member = PE_CDbl(UserSetting(11))
        If PE_CLng(UserSetting(12)) = 1 Then
            IsOffer = "��"
        Else
            IsOffer = "��"
        End If
        ChargeType = PE_CLng(UserSetting(14))
        UnsignedItems = rsUser("UnsignedItems")
        UnreadMsg = PE_CLng(rsUser("UnreadMsg"))
        RegTime = rsUser("RegTime")
        JoinTime = rsUser("JoinTime")
        LoginTimes = rsUser("LoginTimes")
        LastLoginTime = rsUser("LastLoginTime")
        LastLoginIP = rsUser("LastLoginIP")

        If PresentExpPerLogin > 0 Then
            If DateDiff("D", rsUser("LastPresentTime"), Now()) > 0 Then
                Conn.Execute ("update PE_User set UserExp=UserExp+" & PresentExpPerLogin & ",LastPresentTime=" & PE_Now & " where UserID=" & UserID & "")
            End If
        End If
        If PE_CLng(Session("UserID")) = 0 Then
            Conn.Execute ("update PE_User set LastLoginIP='" & UserTrueIP & "',LastLoginTime=" & PE_Now & ",LoginTimes=LoginTimes+1 where UserID=" & UserID & "")
            Session("UserID") = UserID
        End If
    End If
    Set rsUser = Nothing
    DefaultTemplateProjectName = GetDefaultTemplateProjectName()

End Function

Function GetDefaultTemplateProjectName()
    Dim rsProject, strProjectName
    Set rsProject = Conn.Execute("select TemplateProjectName from PE_TemplateProject where IsDefault=" & PE_True)
    If Not rsProject.EOF Then
        strProjectName = rsProject("TemplateProjectName")
    Else
        strProjectName = "����2006��������"
    End If
    Set rsProject = Nothing
    GetDefaultTemplateProjectName = strProjectName
End Function

Function GetClientName(ClientID)
    If ClientID <= 0 Then
        GetClientName = ""
        Exit Function
    End If
    Dim rsClient
    Set rsClient = Conn.Execute("select ClientName from PE_Client where ClientID=" & ClientID & "")
    If rsClient.BOF And rsClient.EOF Then
        GetClientName = ""
    Else
        GetClientName = rsClient(0)
    End If
    rsClient.Close
    Set rsClient = Nothing
End Function


Function GetGroupName(iGroupID)
    Dim rsGroup
    Set rsGroup = Conn.Execute("select GroupName from PE_UserGroup where GroupID=" & iGroupID & "")
    If rsGroup.BOF And rsGroup.EOF Then
        GetGroupName = "δ֪"
    Else
        GetGroupName = rsGroup(0)
    End If
    Set rsGroup = Nothing
End Function

Function CheckBadChar(strChar)
    Dim strBadChar, arrBadChar, i
    strBadChar = "@@,+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & ""
    arrBadChar = Split(strBadChar, ",")
    If strChar = "" Then
        CheckBadChar = False
    Else
        For i = 0 To UBound(arrBadChar)
            If InStr(strChar, arrBadChar(i)) > 0 Then
                CheckBadChar = False
                Exit Function
            End If
        Next
    End If
    CheckBadChar = True
End Function

Function ReplaceUrlBadChar(strChar)
    If strChar = "" Or IsNull(strChar) Then
        ReplaceUrlBadChar = ""
        Exit Function
    End If
    Dim strBadChar, arrBadChar, tempChar, i
    strBadChar = "+,',--,(,),<,>,[,],{,},\,;," & Chr(34) & "," & Chr(0) & ""
    arrBadChar = Split(strBadChar, ",")
    tempChar = strChar
    For i = 0 To UBound(arrBadChar)
        tempChar = Replace(tempChar, arrBadChar(i), "")
    Next
    tempChar = Replace(tempChar, "@@", "@")
    ReplaceUrlBadChar = tempChar
End Function

Function GetNewID(SheetName, FieldName)
    Dim mrs
    Set mrs = Conn.Execute("select max(" & FieldName & ") from " & SheetName & "")
    If IsNull(mrs(0)) Then
        GetNewID = 1
    Else
        GetNewID = mrs(0) + 1
    End If
    Set mrs = Nothing
End Function

Function GetArrFromDictionary(strTableName, strFieldName)
    Dim rsDictionary
    Set rsDictionary = Conn.Execute("select FieldValue from PE_Dictionary where TableName='" & strTableName & "' and FieldName='" & strFieldName & "'")
    If rsDictionary.BOF And rsDictionary.EOF Then
        GetArrFromDictionary = ""
    Else
        GetArrFromDictionary = rsDictionary(0)
    End If
    Set rsDictionary = Nothing
End Function

Function ChkValidDays(iValidNum, iValidUnit, iBeginTime)
    If (iValidNum = "" Or IsNumeric(iValidNum) = False Or iValidUnit = "" Or IsNumeric(iValidUnit) = False Or iBeginTime = "" Or IsDate(iBeginTime) = False) Then
        ChkValidDays = 0
        Exit Function
    End If
    Dim tmpDate, arrInterval
    arrInterval = Array("h", "D", "m", "yyyy")
    If iValidNum = -1 Then
        ChkValidDays = 99999
    Else
        tmpDate = DateAdd(arrInterval(iValidUnit), iValidNum, iBeginTime)
        ChkValidDays = DateDiff("D", Date, tmpDate)
    End If
End Function
'**************************************************
'��������PE_ServerHTMLEncode
'��  �ã���ʾHTML����
'��  ����Content ---- Ҫ���HTML���ַ���
'����ֵ���������ַ���
'**************************************************
Function PE_ServerHTMLEncode(ByVal Content)
    If IsNull(Content) Then
        PE_ServerHTMLEncode = ""
    Else
        PE_ServerHTMLEncode = Server.HTMLEncode(Content)
    End If
End Function
'**************************************************
'��������nohtml
'��  �ã�����html Ԫ��
'��  ����str ---- Ҫ�����ַ�
'����ֵ��û��html ���ַ�
'**************************************************
Public Function nohtml(ByVal str)
    If IsNull(str) Or Trim(str) = "" Then
        nohtml = ""
        Exit Function
    End If
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "(\<.[^\<]*\>)"
    str = re.Replace(str, " ")
    re.Pattern = "(\<\/[^\<]*\>)"
    str = re.Replace(str, " ")
    Set re = Nothing
    
    str = Replace(str, "'", "")
    str = Replace(str, Chr(34), "")
    nohtml = str
End Function
'=================================================
'��������ReplaceBadUrl
'��  �ã����˷Ƿ���Url��ַ����
'=================================================
Private Function ReplaceBadUrl(ByVal strContent)
    Dim rsConfig, regEx
    Set rsConfig = Conn.Execute("select InstallDir,AdminDir from PE_config")
    If rsConfig.BOF And rsConfig.EOF Then
    Else
        Set regEx = New RegExp
        regEx.IgnoreCase = True
        regEx.Global = True
        regEx.Pattern = "\" & rsConfig("InstallDir") & "user\/(.*?).asp"
        strContent = regEx.Replace(strContent, "")
        regEx.Pattern = "\" & rsConfig("InstallDir") & rsConfig("AdminDir") & "\/(.*?).asp"
        strContent = regEx.Replace(strContent, "")
        Set regEx = Nothing
    End If
    Set rsConfig = Nothing
    ReplaceBadUrl = strContent
End Function

%>