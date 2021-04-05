<%
Dim Conn, ConnStr, db, PE_True, PE_False, PE_Now
Dim SqlDatabaseName, SqlPassword, SqlUsername, SqlHostIP
Dim SiteName, SiteTitle, SiteUrl, InstallDir, LogoUrl, WebmasterName, WebmasterEmail, SiteKey
Dim AdminDir, ShowSiteChannel, objName_FSO, FileExt_SiteIndex, FileExt_SiteSpecial
Dim PresentExpPerLogin

Dim EnableUserReg, RegFields_MustFill, EnableCheckCodeOfLogin
Dim RssCodeType
Dim LockIP, LockIPType
Dim UserTrueIP
Dim AllModules, PointName, PointUnit

Const CMS_Edition = 0       '0--�ռ���  1--��׼��  2--רҵ��  3--��ҵ��
Const eShop_Edition = -1    '0--�ռ���  1--��׼��  2--רҵ��  3--��ҵ��
Const CRM_Edition = 0       '0--�ռ���  1--��׼��  2--רҵ��  3--��ҵ��
Const SystemDatabaseType = "ACCESS"     'ϵͳ���ݿ����ͣ�"SQL"ΪMS SQL2000���ݿ⣬"ACCESS"ΪMS ACCESS 2000���ݿ⣬��Ѱ�ֻ��ʹ��ACCESS���ݿ�


'�����ACCESS���ݿ⣬�������޸ĺ���������ݿ���ļ���
db = "\database\PowerEasy2006.mdb"      'ACCESS���ݿ���ļ�������ʹ���������վ��Ŀ¼�ĵľ���·��
                                        '����ǰ�װ����վ��Ŀ¼��ֱ���޸��ļ������ɡ�����ǰ�װ����վĳһĿ¼�£�����ǰ����ϴ�Ŀ¼��
                                        '���磬ϵͳ��װ�ڡ�http://www.powereasy.net/PE2006/��Ŀ¼�£�PE2006Ϊ��װĿ¼����������Ӧ���޸�Ϊ��db="\PE2006\database\PowerEasy2006.mdb"

'�����SQL���ݿ⣬�������޸ĺ��������ݿ�ѡ��
SqlUsername = "PowerEasy"           'SQL���ݿ��û���
SqlPassword = "PowerEasy*9988"          'SQL���ݿ��û�����
SqlDatabaseName = "PowerEasy2006"       'SQL���ݿ���
SqlHostIP = "127.0.0.1"                 'SQL����IP��ַ�����ؿ��á�127.0.0.1����(local)�����Ǳ���������ʵIP��

Call OpenConn
Call GetSiteConfig
Call IsIPlock

Sub OpenConn()
    On Error Resume Next
    If SystemDatabaseType = "SQL" Then
        ConnStr = "Provider = Sqloledb; User ID = " & SqlUsername & "; Password = " & SqlPassword & "; Initial Catalog = " & SqlDatabaseName & "; Data Source = " & SqlHostIP & ";"
    Else
        ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
    End If
    Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.open ConnStr
    If Err Then
        Err.Clear
        Set Conn = Nothing
        Response.Write "���ݿ����ӳ�������Conn.asp�ļ��е����ݿ�������á�"
        Response.End
    End If
    If SystemDatabaseType = "SQL" Then
        PE_True = "1"
        PE_False = "0"
        PE_Now = "getdate()"
    Else
        PE_True = "True"
        PE_False = "False"
        PE_Now = "Now()"
    End If
End Sub

Sub CloseConn()
    On Error Resume Next
    If IsObject(Conn) Then
        Conn.Close
        Set Conn = Nothing
    End If
End Sub

Sub GetSiteConfig()
    Dim rsConfig
    Set rsConfig = Conn.Execute("select * from PE_Config")
    If rsConfig.BOF And rsConfig.EOF Then
        rsConfig.Close
        Set rsConfig = Nothing
        Response.Write "��վ�������ݶ�ʧ��ϵͳ�޷��������У�"
        Response.End
    Else
        SiteName = rsConfig("SiteName")
        SiteTitle = rsConfig("SiteTitle")
        SiteUrl = rsConfig("SiteUrl")
        InstallDir = rsConfig("InstallDir")
        LogoUrl = rsConfig("LogoUrl")
        WebmasterName = rsConfig("WebmasterName")
        WebmasterEmail = rsConfig("WebmasterEmail")
        SiteKey = rsConfig("SiteKey")

        AdminDir = rsConfig("AdminDir")
        ShowSiteChannel = rsConfig("ShowSiteChannel")
        objName_FSO = rsConfig("objName_FSO")
        FileExt_SiteIndex = rsConfig("FileExt_SiteIndex")
        FileExt_SiteSpecial = rsConfig("FileExt_SiteSpecial")

        EnableUserReg = rsConfig("EnableUserReg")
        RegFields_MustFill = rsConfig("RegFields_MustFill")
        AllModules = rsConfig("Modules")
        PointName = rsConfig("PointName")
        PointUnit = rsConfig("PointUnit")
        RssCodeType = rsConfig("RssCodeType")
        LockIP = rsConfig("LockIP")
        LockIPType = rsConfig("LockIPType")
        EnableCheckCodeOfLogin = rsConfig("EnableCheckCodeOfLogin")

        PresentExpPerLogin = rsConfig("PresentExpPerLogin")
    End If
    rsConfig.Close
    Set rsConfig = Nothing
    Application("SiteKey") = SiteKey
    Application("objName_FSO") = objName_FSO
End Sub

Sub IsIPlock()
    UserTrueIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If UserTrueIP = "" Then UserTrueIP = Request.ServerVariables("REMOTE_ADDR")
    If session("IPlock") = "" Then
        session("IPlock") = ChecKIPlock(LockIPType, LockIP, UserTrueIP)
    End If
    If session("IPlock") = True Then
        Response.Write "�Բ�������IP��" & UserTrueIP & "����ϵͳ�޶��������Ժ�վ����ϵ��"
        Response.End
    End If
End Sub

Function EncodeIP(Sip)
    Dim strIP
    strIP = Split(Sip, ".")
    If UBound(strIP) < 3 Then
        EncodeIP = 0
        Exit Function
    End If
    If IsNumeric(strIP(0)) = 0 Or IsNumeric(strIP(1)) = 0 Or IsNumeric(strIP(2)) = 0 Or IsNumeric(strIP(3)) = 0 Then
        Sip = 0
    Else
        Sip = CInt(strIP(0)) * 256 * 256 * 256 + CInt(strIP(1)) * 256 * 256 + CInt(strIP(2)) * 256 + CInt(strIP(3)) - 1
    End If
    EncodeIP = Sip
End Function

'�������Ķ˵���Է��ʺͺ������Ķ˵㽫��������ʡ�
Function ChecKIPlock(ByVal sLockType, ByVal sLockList, ByVal sUserIP)
    Dim IPlock, rsLockIP
    Dim arrLockIPW, arrLockIPB, arrLockIPWCut, arrLockIPBCut
    IPlock = False
    ChecKIPlock = IPlock
    Dim i, sKillIP
    If sLockType = "" Or IsNull(sLockType) Then Exit Function
    If sLockList = "" Or IsNull(sLockList) Then Exit Function
    If sUserIP = "" Or IsNull(sUserIP) Then Exit Function
    sUserIP = CDbl(EncodeIP(sUserIP))
    rsLockIP = Split(sLockList, "|||")
    If sLockType = 4 Then
        arrLockIPB = Split(Trim(rsLockIP(1)), "$$$")
        For i = 0 To UBound(arrLockIPB)
            If arrLockIPB(i) <> "" Then
                arrLockIPBCut = Split(Trim(arrLockIPB(i)), "----")
                IPlock = True
                If CDbl(arrLockIPBCut(0)) > sUserIP Or sUserIP > CDbl(arrLockIPBCut(1)) Then IPlock = False
                If IPlock Then Exit For
            End If
        Next
        If IPlock = True Then
            arrLockIPW = Split(Trim(rsLockIP(0)), "$$$")
            For i = 0 To UBound(arrLockIPW)
                If arrLockIPW(i) <> "" Then
                    arrLockIPWCut = Split(Trim(arrLockIPW(i)), "----")
                    IPlock = True
                    If CDbl(arrLockIPWCut(0)) <= sUserIP And sUserIP <= CDbl(arrLockIPWCut(1)) Then IPlock = False
                    If IPlock Then Exit For
                End If
            Next
        End If
    Else
        If sLockType = 1 Or sLockType = 3 Then
            arrLockIPW = Split(Trim(rsLockIP(0)), "$$$")
            For i = 0 To UBound(arrLockIPW)
                If arrLockIPW(i) <> "" Then
                    arrLockIPWCut = Split(Trim(arrLockIPW(i)), "----")
                    IPlock = True
                    If CDbl(arrLockIPWCut(0)) <= sUserIP And sUserIP <= CDbl(arrLockIPWCut(1)) Then IPlock = False
                    If IPlock Then Exit For
                End If
            Next
        End If
        If IPlock = False And (sLockType = 2 Or sLockType = 3) Then
            arrLockIPB = Split(Trim(rsLockIP(1)), "$$$")
            For i = 0 To UBound(arrLockIPB)
                If arrLockIPB(i) <> "" Then
                    arrLockIPBCut = Split(Trim(arrLockIPB(i)), "----")
                    IPlock = True
                    If CDbl(arrLockIPBCut(0)) > sUserIP Or sUserIP > CDbl(arrLockIPBCut(1)) Then IPlock = False
                    If IPlock Then Exit For
                End If
            Next
        End If
    End If
    ChecKIPlock = IPlock
End Function
%>
