<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = False
Const PurviewLevel = 1
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
Server.ScriptTimeOut = 99999999
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/Function.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->
<%
Dim strtmp, hf, fso, ObjInstalled_FSO, MaxPerPage, MaxPageCol, OutNum, XmlMaxPerPage, XmlOutNum, frequency, Priority, ArtPage, SoftPage, PhotoPage, ProductPage
Dim EnableRss, UOffset, Action2

Action2 = Trim(Request("Action2"))

Dim rsConfig
Set rsConfig = Conn.Execute("select EnableRss from PE_Config")
If rsConfig.bof And rsConfig.EOF Then
    rsConfig.Close
    Set rsConfig = Nothing
    Response.Write "��վ�������ݶ�ʧ��ϵͳ�޷��������У�"
    Response.End
Else
    EnableRss = rsConfig("EnableRss")
End If
rsConfig.Close
Set rsConfig = Nothing
If Right(SiteUrl, 1) <> "/" Then SiteUrl = SiteUrl & "/"
%>
<html><head><title>������վ�ۺ�����</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href='Admin_Style.css' rel='stylesheet' type='text/css'>
</head>
<body leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>
<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>
  <tr class='topbg'>
    <td height='22' colspan='2' align='center'><strong>������վ�ۺ�����</strong></td>
  </tr>
  <tr class='tdbg'>
    <td width='70' height='30'><strong>����˵����</strong></td>
    <td>���ɲ����Ƚ�����ϵͳ��Դ����ʱ��ÿ������ʱ���뾡������Ҫ���ɵ��ļ�����</td>
  </tr>
</table>
<br>
<%
If Action2 = "" Then
%>
<table width='100%' border='0' align='center' cellpadding='3' cellspacing='1' class='border'>
    <tr><td class='title'>RSS���ɲ���</td></tr>
    <tr><td class='tdbg'>
        <table width='530' border='0' align='center' cellpadding='0' cellspacing='0'>
            <form name='formrss' method='post' action='Admin_CreateOther.asp'>
            <tr><td height='40'>
                ������վ��ҳ�ģңӣ�ҳ�棬�������ãңӣӻ���վ��ҳΪ��̬���ӣи�ʽʱ����������Ч��<br>
                <input name='Action2' type='hidden' id='Action2' value='CreateRss'>
                <input name='submit' type='submit' id='submit' value='��ʼ����>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>XML���ɲ���</td></tr>
    <tr><td class='tdbg'>
        <table width='530' border='0' align='center' cellpadding='0' cellspacing='0'>
            <form name='formxml' method='post' action='Admin_CreateOther.asp'>
            <tr><td height='40'>
                ���ɣأͣ����ݽ���ҳ����ַΪ<a href='<% =SiteUrl %>xml/xml.xml' target='_blank'><% =SiteUrl %>xml/xml.xml</a><br>
                <input name='Channel' type='checkbox' id='Channel' value='True' checked>����Ƶ������<br>
                <input name='Guest' type='checkbox' id='Guest' value='True' checked>������������<br>
                <input name='Author' type='checkbox' id='Author' value='True' checked>������������<br>
                <input name='User' type='checkbox' id='User' value='True' checked>���ɻ�Ա����<br>
                <input name='Site' type='checkbox' id='Site' value='True' checked>��������վ��<br>
                <input name='Announce' type='checkbox' id='Announce' value='True' checked>���ɹ����б�<br>
                <input name='Action2' type='hidden' id='Action2' value='CreateXml'>
                <input name='submit' type='submit' id='submit' value='��ʼ����>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>HTML��ͼ���ɲ���</td></tr>
    <tr><td align='center' class='tdbg'>
        <table width='530' border='0' cellspacing='0' cellpadding='0'>
            <form name='formap' method='post' action='Admin_CreateOther.asp'>
            <tr><td>
                ����HTML��ʽ��ȫվ��ͼҳ�档<br>
                ���������<input name='OutNum' id='OutNum' value='500' size=8 maxlength='5'>&nbsp;<font color=#888888>�ȣԣ̵ͣ�ͼ���������</font><br>
                ÿҳ������<input name='MaxPerPage' id='MaxPerPage' value='100' size=8 maxlength='3'>&nbsp;<font color=#888888>ÿҳ������������ܴ��ڣ�����</font><br>
                ��ҳ������<input name='MaxPageCol' id='MaxPageCol' value='27' size=8 maxlength='2'>&nbsp;<font color=#888888>��ͼ��ҳ����ÿ����ʾ��</font><br>
                <input name='Action2' type='hidden' id='Action2' value='CreateMap'>
                <input name='submit' type='submit' id='submit' value='��ʼ����>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
    <tr><td class='title'>XML��ͼ���ɲ���</td></tr>
    <tr><td align='center' class='tdbg'>
        <table width='530' border='0' cellspacing='0' cellpadding='0'><a href='http://www.google.com/webmasters/sitemaps/login' target='_blank'><img src="images/GoogleSiteMaplogo.gif" border=0></a>���ɷ���GOOGLE�淶��XML��ʽ��ͼҳ��
            <form name='formxmlmap' method='post' action='Admin_CreateOther.asp'>
            <tr><td>
                ���������<input name='XmlOutNum' id='XmlOutNum' value='500' size=10 maxlength='5'>&nbsp;<font color=#888888>�أ̵ͣ�ͼ���������</font><br>
                ÿҳ������<input name='XmlMaxPerPage' id='XmlMaxPerPage' value='100' size=10 maxlength='4'>&nbsp;<font color=#888888>ÿҳ������,GOOGLE�淶Ҫ�󲻵ô��ڣ�������</font><br>
                &nbsp;&nbsp;ʱ��ƫ��<input name='UOffset' id='UOffset' value='08' size=10 maxlength='2'>&nbsp;<font color=#888888>Ĭ���й���½Ϊ��</font><br>
                &nbsp;&nbsp;����Ƶ��<SELECT name=frequency> <OPTION value=always>��ʱ����</OPTION> <OPTION value=hourly>ÿ С ʱ</OPTION> <OPTION value=daily>ÿ�����</OPTION> <OPTION value=weekly>ÿ�ܸ���</OPTION> <OPTION value=monthly selected>ÿ�¸���</OPTION> <OPTION value=yearly>ÿ�����</OPTION> <OPTION value=never>�Ӳ�����</OPTION></SELECT>&nbsp;<font color=#888888>����վ�����ݸ����������ѡ��</font><br>
                &nbsp;&nbsp;Ȩ&nbsp;&nbsp;&nbsp;&nbsp;��<input name='Priority' id='Priority' value='0.5' size=10 maxlength='3'>&nbsp;<font color=#888888>0-1.0֮��,�Ƽ�ʹ��Ĭ��ֵ</font><br>
                <input name='Action2' type='hidden' id='Action2' value='CreateXmlMap'>
                <input name='submit' type='submit' id='submit' value='��ʼ����>>'>
            </td></tr>
            </form>
        </table>
    </td></tr>
</table>
<%
Else
    Select Case Action2
    Case "CreateRss"
        If EnableRss = True Then
            Call GetRssIndex_file
            Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a>"
        Else
            Response.Write "<br><br><b>���Ѿ�������RSS����,ҳ��δ����..........<a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a></b>"
        End If
    Case "CreateXml"
        Call PE_CreateXml
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a>"
    Case "CreateMap"
        ObjInstalled_FSO = IsObjInstalled(objName_FSO)
        If ObjInstalled_FSO = True Then Set fso = Server.CreateObject(objName_FSO)
        
        OutNum = Trim(Request("OutNum"))
        If OutNum = "" Or Not IsNumeric(OutNum) Then
            OutNum = 500
        Else
            OutNum = Int(OutNum)
        End If
        MaxPerPage = Int(Trim(Request("MaxPerPage")))
        If MaxPerPage = "" Or Not IsNumeric(MaxPerPage) Then
            MaxPerPage = 100
        Else
            MaxPerPage = Int(MaxPerPage)
        End If
        MaxPageCol = Int(Trim(Request("MaxPageCol")))
        If MaxPageCol = "" Or Not IsNumeric(MaxPageCol) Then
            MaxPageCol = 27
        Else
            MaxPageCol = Int(MaxPageCol)
        End If


        Response.Write "<br><br><b>��������������Mapҳ��.........."
        Call OutArticleMap
        Response.Write "</b>"

        Response.Write "<br><br><b>�������������Mapҳ��.........."
        Call OutSoftMap
        Response.Write "</b>"

        Response.Write "<br><br><b>��������ͼƬ��Mapҳ��.........."
        Call OutPhotoMap
        Response.Write "</b>"

        If CMS_Edition > 0 Then
            Response.Write "<br><br><b>����������Ʒ��Mapҳ��.........."
            Call OutProductMap
            Response.Write "</b>"
        End If
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a>"
    Case "CreateXmlMap"
        ObjInstalled_FSO = IsObjInstalled(objName_FSO)
        If ObjInstalled_FSO = True Then Set fso = Server.CreateObject(objName_FSO)
        XmlOutNum = Trim(Request("XmlOutNum"))
        If XmlOutNum = "" Or Not IsNumeric(XmlOutNum) Then
            XmlOutNum = 500
        Else
            XmlOutNum = Int(XmlOutNum)
        End If
        XmlMaxPerPage = Trim(Request("XmlMaxPerPage"))
        If XmlMaxPerPage = "" Or Not IsNumeric(XmlMaxPerPage) Then
            XmlMaxPerPage = 27
        Else
            XmlMaxPerPage = Int(XmlMaxPerPage)
        End If
        UOffset = Trim(Request("UOffset"))
        If UOffset = "" Or Not IsNumeric(UOffset) Then
            UOffset = 8
        Else
            UOffset = Int(UOffset)
        End If
        frequency = Trim(Request("frequency"))
        If frequency = "" Then frequency = "Monthly"
        Priority = Trim(Request("Priority"))
        If Priority = "" Then Priority = "0.5"
        
        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼ����ҳ��.........."
        Call OutXmlMap(1)
        Response.Write "</b>"

        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼ���ҳ��.........."
        Call OutXmlMap(2)
        Response.Write "</b>"

        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼͼƬҳ��.........."
        Call OutXmlMap(3)
        Response.Write "</b>"
    
        If CMS_Edition > 0 Then
            Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼ��Ʒҳ��.........."
            Call OutXmlMap(5)
            Response.Write "</b>"
        End If

        Response.Write "<br><br><b>��������GOOGLE�淶XML��ͼ����ҳ��.........."
        Call OutXmlIndexMap
        Response.Write "</b>"
        Response.Write "<br><br><a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a>"
    Case Else
        Response.Write "<br><br><b>��������..........<a href='Admin_CreateOther.asp'>&lt;&lt; �������ɹ���</a></b>"
    End Select
    Set hf = Nothing
End If
%>
</body>
</html>
<!-- Powered by: PowerEasy 2006 -->
<%
Sub GetRssIndex_file()
    On Error Resume Next
    Dim PE_Rss
    Set PE_Rss = Server.CreateObject("PE_CMS6.ShowRss")
    PE_Rss.iConnStr = ConnStr
    PE_Rss.iSystemDatabaseType = SystemDatabaseType
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    Call PE_Rss.GetRssIndex_file
    Set PE_Rss = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub

Sub PE_CreateXml()
    On Error Resume Next
    Dim PE_Xml
    Set PE_Xml = Server.CreateObject("PE_CMS6.Xml")
    PE_Xml.iConnStr = ConnStr
    PE_Xml.iCMS_Edition = CMS_Edition
    PE_Xml.iSystemDatabaseType = SystemDatabaseType
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    Call PE_Xml.main
    Set PE_Xml = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub

Sub OutArticleMap()
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, iAuthor
    Dim oldChannelID: oldChannelID = 0

    sqlArticle = "select top " & OutNum & " A.ArticleID,A.ChannelID,A.ClassID,A.Title,A.Author,A.UpdateTime,A.Elite,A.Status,A.InfoPoint,A.Deleted,A.LinkUrl,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Article A inner join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.ArticleID Desc"
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "��������!�ݲ�����ҳ��!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalPage = totalPut \ MaxPerPage
        Else
            totalPage = totalPut \ MaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(11)
            ParentDir = rsArticle(12)
            ClassPurview = rsArticle(13)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = rsChannel("StructureType")
                    If CMS_Edition < 1 Then StructureType = 0
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                End If
                rsChannel.Close
            End If

            iAuthor = rsArticle(4)
            If UseCreateHTML > 0 And ClassPurview = 0 And (rsArticle(8) = 0 Or CMS_Edition < 1) Then
                strHTML = strHTML & "<li><a href='" & strInstallDir & iChannelDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(5)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(5), rsArticle(0)) & GetFileExt(FileExt_Item) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            Else
                strHTML = strHTML & "<li><a href='" & strInstallDir & iChannelDir & "/ShowArticle.asp?ArticleID=" & rsArticle(0) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            End If
            i = i + 1

            If i > MaxPerPage Then
                Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "SiteMap/Article" & CurrentPage & ".htm"), 2, True)
                strtmp = "<html>" & vbCrLf
                strtmp = strtmp & "<head>" & vbCrLf
                strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
                strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
                strtmp = strtmp & "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
                strtmp = strtmp & "</head>" & vbCrLf
                strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ȫվ�������� >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>��ҳ:"
                For j = 1 To totalPage
                    If CurrentPage = j Then
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " [" & j & "]<br>"
                        Else
                            strtmp = strtmp & " [" & j & "] "
                        End If
                    Else
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Article" & j & ".htm'>" & j & "</a><br>"
                        Else
                            strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Article" & j & ".htm'>" & j & "</a> "
                        End If
                    End If
                Next
                strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
                strtmp = strtmp & "</html>" & vbCrLf
                hf.Write strtmp
                hf.Close
                Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "SiteMap/Article" & CurrentPage & ".htm' target='_blank'>" & strInstallDir & "SiteMap/Article" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "SiteMap/Article" & CurrentPage & ".htm"), 2, True)
        strtmp = "<html>" & vbCrLf
        strtmp = strtmp & "<head>" & vbCrLf
        strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
        strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
        strtmp = strtmp & "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
        strtmp = strtmp & "</head>" & vbCrLf
        strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ȫվ�������� >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>��ҳ:"
        For j = 1 To totalPage
            If CurrentPage = j Then
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " [" & j & "]<br>"
                Else
                    strtmp = strtmp & " [" & j & "] "
                End If
            Else
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Article" & j & ".htm'>" & j & "</a><br>"
                Else
                    strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Article" & j & ".htm'>" & j & "</a> "
                End If
            End If
        Next
        strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
        strtmp = strtmp & "</html>" & vbCrLf
        hf.Write strtmp
        hf.Close
        Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "SiteMap/Article" & CurrentPage & ".htm' target='_blank'>" & strInstallDir & "SiteMap/Article" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutSoftMap()
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, iAuthor
    Dim oldChannelID: oldChannelID = 0

    sqlArticle = "select top " & OutNum & " A.SoftID,A.ChannelID,A.ClassID,A.SoftName,A.Author,A.UpdateTime,A.Elite,A.Status,A.Deleted,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Soft A inner join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.SoftID Desc"
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "��������!�ݲ�����ҳ��!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalPage = totalPut \ MaxPerPage
        Else
            totalPage = totalPut \ MaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(10)
            ParentDir = rsArticle(11)
            ClassPurview = rsArticle(12)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = rsChannel("StructureType")
                    If CMS_Edition < 1 Then StructureType = 0
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                End If
                rsChannel.Close
            End If
        
            iAuthor = rsArticle(4)

            If UseCreateHTML > 0 And ClassPurview = 0 And (rsArticle(9) = 0 Or CMS_Edition < 1) Then
                strHTML = strHTML & "<li><a href='" & strInstallDir & iChannelDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(5)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(5), rsArticle(0)) & GetFileExt(FileExt_Item) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            Else
                strHTML = strHTML & "<li><a href='" & strInstallDir & iChannelDir & "/ShowSoft.asp?SoftID=" & rsArticle(0) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            End If

            i = i + 1
            If i > MaxPerPage Then
                Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "SiteMap/Soft" & CurrentPage & ".htm"), 2, True)
                strtmp = "<html>" & vbCrLf
                strtmp = strtmp & "<head>" & vbCrLf
                strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
                strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
                strtmp = strtmp & "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
                strtmp = strtmp & "</head>" & vbCrLf
                strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>��ҳ:"
                For j = 1 To totalPage
                    If CurrentPage = j Then
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " [" & j & "]<br>"
                        Else
                            strtmp = strtmp & " [" & j & "] "
                        End If
                    Else
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Soft" & j & ".htm'>" & j & "</a><br>"
                        Else
                            strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Soft" & j & ".htm'>" & j & "</a> "
                        End If
                    End If
                Next
                strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
                strtmp = strtmp & "</html>" & vbCrLf
                hf.Write strtmp
                hf.Close
                Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "SiteMap/Soft" & CurrentPage & ".htm' target='_blank'>" & strInstallDir & "SiteMap/Soft" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "SiteMap/Soft" & CurrentPage & ".htm"), 2, True)
        strtmp = "<html>" & vbCrLf
        strtmp = strtmp & "<head>" & vbCrLf
        strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
        strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
        strtmp = strtmp & "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
        strtmp = strtmp & "</head>" & vbCrLf
        strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>��ҳ:"
        For j = 1 To totalPage
            If CurrentPage = j Then
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " [" & j & "]<br>"
                Else
                    strtmp = strtmp & " [" & j & "] "
                End If
            Else
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Soft" & j & ".htm'>" & j & "</a><br>"
                Else
                    strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Soft" & j & ".htm'>" & j & "</a> "
                End If
            End If
        Next
        strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
        strtmp = strtmp & "</html>" & vbCrLf
        hf.Write strtmp
        hf.Close
        Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "SiteMap/Soft" & CurrentPage & ".htm' target='_blank'>" & strInstallDir & "SiteMap/Soft" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
    
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutPhotoMap()
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, iAuthor
    Dim oldChannelID: oldChannelID = 0

    sqlArticle = "select top " & OutNum & " A.PhotoID,A.ChannelID,A.ClassID,A.PhotoName,A.Author,A.UpdateTime,A.Status,A.Deleted,A.InfoPoint,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Photo A inner join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.PhotoID Desc"
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "��������!�ݲ�����ҳ��!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalPage = totalPut \ MaxPerPage
        Else
            totalPage = totalPut \ MaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(9)
            ParentDir = rsArticle(10)
            ClassPurview = rsArticle(11)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = rsChannel("StructureType")
                    If CMS_Edition < 1 Then StructureType = 0
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                End If
                rsChannel.Close
            End If
    
            iAuthor = rsArticle(4)

            If UseCreateHTML > 0 And ClassPurview = 0 And (rsArticle(8) = 0 Or CMS_Edition < 1) Then
                strHTML = strHTML & "<li><a href='" & strInstallDir & iChannelDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(5)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(5), rsArticle(0)) & GetFileExt(FileExt_Item) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            Else
                strHTML = strHTML & "<li><a href='" & strInstallDir & iChannelDir & "/ShowPhoto.asp?PhotoID=" & rsArticle(0) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            End If
            i = i + 1
            If i > MaxPerPage Then
                Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "SiteMap/Photo" & CurrentPage & ".htm"), 2, True)
                strtmp = "<html>" & vbCrLf
                strtmp = strtmp & "<head>" & vbCrLf
                strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
                strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
                strtmp = strtmp & "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
                strtmp = strtmp & "</head>" & vbCrLf
                strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>��ҳ:"
                For j = 1 To totalPage
                    If CurrentPage = j Then
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " [" & j & "]<br>"
                        Else
                            strtmp = strtmp & " [" & j & "] "
                        End If
                    Else
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Photo" & j & ".htm'>" & j & "</a><br>"
                        Else
                            strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Photo" & j & ".htm'>" & j & "</a> "
                        End If
                    End If
                Next
                strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
                strtmp = strtmp & "</html>" & vbCrLf
                hf.Write strtmp
                hf.Close
                Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "SiteMap/Photo" & CurrentPage & ".htm' target='_blank'>" & strInstallDir & "SiteMap/Photo" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "SiteMap/Photo" & CurrentPage & ".htm"), 2, True)
        strtmp = "<html>" & vbCrLf
        strtmp = strtmp & "<head>" & vbCrLf
        strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
        strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
        strtmp = strtmp & "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
        strtmp = strtmp & "</head>" & vbCrLf
        strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>��ҳ:"
        For j = 1 To totalPage
            If CurrentPage = j Then
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " [" & j & "]<br>"
                Else
                    strtmp = strtmp & " [" & j & "] "
                End If
            Else
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Photo" & j & ".htm'>" & j & "</a><br>"
                Else
                    strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Photo" & j & ".htm'>" & j & "</a> "
                End If
            End If
        Next
        strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
        strtmp = strtmp & "</html>" & vbCrLf
        hf.Write strtmp
        hf.Close
        Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "SiteMap/Photo" & CurrentPage & ".htm' target='_blank'>" & strInstallDir & "SiteMap/Photo" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutProductMap()
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, iAuthor
    Dim oldChannelID: oldChannelID = 0

    sqlArticle = "select top " & OutNum & " A.ProductID,A.ChannelID,A.ClassID,A.ProductName,A.ProducerName,A.UpdateTime,A.EnableSale,A.Deleted,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Product A inner join PE_Class C on A.ClassID=C.ClassID Where A.Deleted=" & PE_False & " and A.EnableSale=" & PE_True & " order by A.ProductID Desc"
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "��������!�ݲ�����ҳ��!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod MaxPerPage) = 0 Then
            totalPage = totalPut \ MaxPerPage
        Else
            totalPage = totalPut \ MaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(8)
            ParentDir = rsArticle(9)
            ClassPurview = rsArticle(10)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = rsChannel("StructureType")
                    If CMS_Edition < 1 Then StructureType = 0
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                End If
                rsChannel.Close
            End If
        
            iAuthor = rsArticle(4)

            If UseCreateHTML > 0 Then
                strHTML = strHTML & "<li><a href='" & strInstallDir & iChannelDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(5)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(5), rsArticle(0)) & GetFileExt(FileExt_Item) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            Else
                strHTML = strHTML & "<li><a href='" & strInstallDir & iChannelDir & "/ShowProduct.asp?ProductID=" & rsArticle(0) & "'>" & rsArticle(3) & "</a> - [" & iAuthor & "]</li>" & vbCrLf
            End If
            i = i + 1
            If i > MaxPerPage Then
                Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "SiteMap/Product" & CurrentPage & ".htm"), 2, True)
                strtmp = "<html>" & vbCrLf
                strtmp = strtmp & "<head>" & vbCrLf
                strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
                strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
                strtmp = strtmp & "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
                strtmp = strtmp & "</head>" & vbCrLf
                strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
                strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
                strtmp = strtmp & strHTML & "<br><br>��ҳ:"
                For j = 1 To totalPage
                    If CurrentPage = j Then
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " [" & j & "]<br>"
                        Else
                            strtmp = strtmp & " [" & j & "] "
                        End If
                    Else
                        If (j Mod MaxPageCol) = 0 Then
                            strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Product" & j & ".htm'>" & j & "</a><br>"
                        Else
                            strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Product" & j & ".htm'>" & j & "</a> "
                        End If
                    End If
                Next
                strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
                strtmp = strtmp & "</html>" & vbCrLf
                hf.Write strtmp
                hf.Close
                Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "SiteMap/Product" & CurrentPage & ".htm' target='_blank'>" & strInstallDir & "SiteMap/Product" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "SiteMap/Product" & CurrentPage & ".htm"), 2, True)
        strtmp = "<html>" & vbCrLf
        strtmp = strtmp & "<head>" & vbCrLf
        strtmp = strtmp & "<title>" & SiteName & "-SiteMap</title>" & vbCrLf
        strtmp = strtmp & "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
        strtmp = strtmp & "<link href='" & strInstallDir & "Skin/DefaultSkin.css' rel='stylesheet' type='text/css'>" & vbCrLf
        strtmp = strtmp & "</head>" & vbCrLf
        strtmp = strtmp & "<body><table width='760' align='center'><tr><td>" & vbCrLf
        strtmp = strtmp & "<a href='" & SiteUrl & "'>" & SiteName & "</a> >> ��վ��ͼ >> ��" & CurrentPage & "ҳ:<br>" & vbCrLf
        strtmp = strtmp & strHTML & "<br><br>��ҳ:"
        For j = 1 To totalPage
            If CurrentPage = j Then
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " [" & j & "]<br>"
                Else
                    strtmp = strtmp & " [" & j & "] "
                End If
            Else
                If (j Mod MaxPageCol) = 0 Then
                    strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Product" & j & ".htm'>" & j & "</a><br>"
                Else
                    strtmp = strtmp & " <a href='" & strInstallDir & "SiteMap/Product" & j & ".htm'>" & j & "</a> "
                End If
            End If
        Next
        strtmp = strtmp & "</td></tr></table></body>" & vbCrLf
        strtmp = strtmp & "</html>" & vbCrLf
        hf.Write strtmp
        hf.Close
        Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "SiteMap/Product" & CurrentPage & ".htm' target='_blank'>" & strInstallDir & "SiteMap/Product" & CurrentPage & ".htm</a>��<font color=red>�ɹ�!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutXmlMap(OutType)
    Dim rsArticle, sqlArticle, rsChannel, strHTML, totalPut, totalPage, CurrentPage, i, j
    Dim iChannelDir, UseCreateHTML, StructureType, FileNameType, FileExt_Item, ClassDir, ParentDir, ClassPurview, AspName, OutFileName
    Dim oldChannelID: oldChannelID = 0
  
    Select Case OutType
    Case 1
        sqlArticle = "select top " & XmlOutNum & " A.ArticleID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.LinkUrl,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Article A inner join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.ArticleID Desc"
    Case 2
    sqlArticle = "select top " & XmlOutNum & " A.SoftID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Soft A inner join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.SoftID Desc"
    Case 3
    sqlArticle = "select top " & XmlOutNum & " A.PhotoID,A.ChannelID,A.ClassID,A.UpdateTime,A.Status,A.InfoPoint,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Photo A inner join PE_Class C on A.ClassID=C.ClassID Where A.Status=3 and A.Deleted=" & PE_False & " order by A.PhotoID Desc"
    Case 5
    sqlArticle = "select top " & XmlOutNum & " A.ProductID,A.ChannelID,A.ClassID,A.UpdateTime,A.EnableSale,A.Stocks,A.Deleted,A.Hits,C.ClassDir,C.ParentDir,C.ClassPurview from PE_Product A inner join PE_Class C on A.ClassID=C.ClassID Where A.Deleted=" & PE_False & " and A.EnableSale=" & PE_True & " order by A.ProductID Desc"
    End Select
    Set rsArticle = Server.CreateObject("adodb.recordset")
    rsArticle.Open sqlArticle, Conn, 1, 1
    If rsArticle.bof And rsArticle.EOF Then
        Response.Write "��������!�ݲ�����ҳ��!<br>"
    Else
        totalPut = rsArticle.recordcount
        If (totalPut Mod XmlMaxPerPage) = 0 Then
            totalPage = totalPut \ XmlMaxPerPage
        Else
            totalPage = totalPut \ XmlMaxPerPage + 1
        End If
        i = 1
        CurrentPage = 1

        Do While Not rsArticle.EOF

            ClassDir = rsArticle(8)
            ParentDir = rsArticle(9)
            ClassPurview = rsArticle(10)

            If rsArticle(1) <> oldChannelID Then
                Set rsChannel = Conn.Execute("select Top 1 ChannelID,ChannelDir,UseCreateHTML,StructureType,FileNameType,FileExt_Item from PE_Channel where ChannelID=" & rsArticle(1))
                If Not (rsChannel.bof And rsChannel.EOF) Then
                    iChannelDir = rsChannel("ChannelDir")
                    UseCreateHTML = rsChannel("UseCreateHTML")
                    StructureType = rsChannel("StructureType")
                    If CMS_Edition < 1 Then StructureType = 0
                    FileNameType = rsChannel("FileNameType")
                    FileExt_Item = rsChannel("FileExt_Item")
                End If
                rsChannel.Close
            End If
            Select Case OutType
            Case 1
                AspName = "/ShowArticle.asp?ArticleID="
                OutFileName = "sitemap_article_"
            Case 2
                AspName = "/ShowSoft.asp?SoftID="
                OutFileName = "sitemap_Soft_"
            Case 3
                AspName = "/ShowPhoto.asp?PhotoID="
                OutFileName = "sitemap_Photo_"
            Case 5
                AspName = "/ShowProduct.asp?ProductID="
                OutFileName = "sitemap_Product_"
            End Select
            strHTML = strHTML & "<url>" & vbCrLf
            If OutType < 4 Then
                If UseCreateHTML > 0 And ClassPurview = 0 And (rsArticle(5) = 0 Or CMS_Edition < 1) Then
                    strHTML = strHTML & "<loc>" & SiteUrl & iChannelDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(3)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(3), rsArticle(0)) & GetFileExt(FileExt_Item) & "</loc>" & vbCrLf
                Else
                    strHTML = strHTML & "<loc>" & SiteUrl & iChannelDir & AspName & rsArticle(0) & "</loc>" & vbCrLf
                End If
            ElseIf OutType = 5 Then
                If UseCreateHTML > 0 Then
                    strHTML = strHTML & "<loc>" & SiteUrl & iChannelDir & GetItemPath(StructureType, ParentDir, ClassDir, rsArticle(3)) & GetItemFileName(FileNameType, iChannelDir, rsArticle(3), rsArticle(0)) & GetFileExt(FileExt_Item) & "</loc>" & vbCrLf
                Else
                    strHTML = strHTML & "<loc>" & SiteUrl & iChannelDir & AspName & rsArticle(0) & "</loc>" & vbCrLf
                End If
            End If
            strHTML = strHTML & "<lastmod>" & iso8601date(rsArticle(3), UOffset) & "</lastmod>" & vbCrLf
            strHTML = strHTML & "<changefreq>" & frequency & "</changefreq>" & vbCrLf
            strHTML = strHTML & "<priority>" & Priority & "</priority>" & vbCrLf
            strHTML = strHTML & "</url>" & vbCrLf
            i = i + 1

            If i > XmlMaxPerPage Then
                Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & OutFileName & CurrentPage & ".xml"), 2, True)
                strtmp = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
                strtmp = strtmp & "<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">" & vbCrLf
                strtmp = strtmp & strHTML
                strtmp = strtmp & "</urlset>" & vbCrLf
                hf.Write strtmp
                hf.Close
                Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & strInstallDir & OutFileName & CurrentPage & ".xml</a>��<font color=red>�ɹ�!</font>"
                CurrentPage = CurrentPage + 1
                i = 1
                strHTML = ""
            End If
            oldChannelID = rsArticle(1)
            rsArticle.movenext
        Loop
        Set rsChannel = Nothing

        Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & OutFileName & CurrentPage & ".xml"), 2, True)
        strtmp = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
        strtmp = strtmp & "<urlset xmlns=""http://www.google.com/schemas/sitemap/0.84"">" & vbCrLf
        strtmp = strtmp & strHTML
        strtmp = strtmp & "</urlset>" & vbCrLf
        hf.Write strtmp
        hf.Close
        Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & OutFileName & CurrentPage & ".xml' target='_blank'>" & strInstallDir & OutFileName & CurrentPage & ".xml</a>)<font color=red>�ɹ�!</font>"
        strHTML = strHTML & "<br>" & vbCrLf
    End If
    Select Case OutType
    Case 1
        ArtPage = totalPage
    Case 2
        SoftPage = totalPage
    Case 3
        PhotoPage = totalPage
    Case 5
        ProductPage = totalPage
    End Select
    rsArticle.Close
    Set rsArticle = Nothing
End Sub

Sub OutXmlIndexMap()
    Dim strtmp, j
    Set hf = fso.OpenTextFile(Server.MapPath(strInstallDir & "sitemap_index.xml"), 2, True)
    strtmp = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    strtmp = strtmp & "<sitemapindex xmlns=""http://www.google.com/schemas/sitemap/0.84"">" & vbCrLf
    If ArtPage > 0 Then
        For j = 1 To ArtPage
            strtmp = strtmp & "<sitemap>" & vbCrLf
            strtmp = strtmp & "<loc>" & SiteUrl & "sitemap_article_" & j & ".xml</loc>" & vbCrLf
            strtmp = strtmp & "<lastmod>" & iso8601date(Now(), UOffset) & "</lastmod>" & vbCrLf
            strtmp = strtmp & "</sitemap>" & vbCrLf
        Next
    End If
    If SoftPage > 0 Then
        For j = 1 To SoftPage
            strtmp = strtmp & "<sitemap>" & vbCrLf
            strtmp = strtmp & "<loc>" & SiteUrl & "sitemap_Soft_" & j & ".xml</loc>" & vbCrLf
            strtmp = strtmp & "<lastmod>" & iso8601date(Now(), UOffset) & "</lastmod>" & vbCrLf
            strtmp = strtmp & "</sitemap>" & vbCrLf
        Next
    End If
    If PhotoPage > 0 Then
        For j = 1 To PhotoPage
            strtmp = strtmp & "<sitemap>" & vbCrLf
            strtmp = strtmp & "<loc>" & SiteUrl & "sitemap_Photo_" & j & ".xml</loc>" & vbCrLf
            strtmp = strtmp & "<lastmod>" & iso8601date(Now(), UOffset) & "</lastmod>" & vbCrLf
            strtmp = strtmp & "</sitemap>" & vbCrLf
        Next
    End If
    If ProductPage > 0 And CMS_Edition > 0 Then
        For j = 1 To ProductPage
            strtmp = strtmp & "<sitemap>" & vbCrLf
            strtmp = strtmp & "<loc>" & SiteUrl & "sitemap_Product_" & j & ".xml</loc>" & vbCrLf
            strtmp = strtmp & "<lastmod>" & iso8601date(Now(), UOffset) & "</lastmod>" & vbCrLf
            strtmp = strtmp & "</sitemap>" & vbCrLf
        Next
    End If
    strtmp = strtmp & "</sitemapindex>" & vbCrLf
    hf.Write strtmp
    hf.Close
    Response.Write "<br> ����ҳ�棨<a href='" & strInstallDir & "sitemap_index.xml' target='_blank'>" & strInstallDir & "sitemap_index.xml</a>��<font color=red>�ɹ�!</font>��&nbsp;[<a href='http://www.google.com/webmasters/sitemaps/ping?sitemap=" & SiteUrl & "sitemap_index.xml' target='_blank'>��������ύ��Google</a>]"
End Sub

'**************************************************
'��������GetItemPath
'��  �ã������Ŀ·��
'��  ����iStructureType ---- Ŀ¼�ṹ��ʽ
'        sParentDir ---- ����ĿĿ¼
'        sClassDir ---- ��ǰ��ĿĿ¼
'        UpdateTime ---- ��ĿĿ¼
'����ֵ�������Ŀ·��
'**************************************************
Public Function GetItemPath(iStructureType, sParentDir, sClassDir, UpdateTime)
    Select Case iStructureType
    Case 0      'Ƶ��/����/С��/�·�/�ļ�����Ŀ�ּ����ٰ��·ݱ��棩
        GetItemPath = sParentDir & sClassDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 1      'Ƶ��/����/С��/����/�ļ�����Ŀ�ּ����ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = sParentDir & sClassDir & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 2      'Ƶ��/����/С��/�ļ�����Ŀ�ּ������ٰ��·ݣ�
        GetItemPath = sParentDir & sClassDir & "/"
    Case 3      'Ƶ��/��Ŀ/�·�/�ļ�����Ŀƽ�����ٰ��·ݱ��棩
        GetItemPath = "/" & sClassDir & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 4      'Ƶ��/��Ŀ/����/�ļ�����Ŀƽ�����ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & sClassDir & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 5      'Ƶ��/��Ŀ/�ļ�����Ŀƽ�������ٰ��·ݣ�
        GetItemPath = "/" & sClassDir & "/"
    Case 6      'Ƶ��/�ļ���ֱ�ӷ���Ƶ��Ŀ¼�У�
        GetItemPath = "/"
    Case 7      'Ƶ��/HTML/�ļ���ֱ�ӷ���ָ���ġ�HTML���ļ����У�
        GetItemPath = "/HTML/"
    Case 8      'Ƶ��/���/�ļ���ֱ�Ӱ���ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/"
    Case 9      'Ƶ��/�·�/�ļ���ֱ�Ӱ��·ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 10     'Ƶ��/����/�ļ���ֱ�Ӱ����ڱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 11     'Ƶ��/���/�·�/�ļ����Ȱ���ݣ��ٰ��·ݱ��棬ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/"
    Case 12     'Ƶ��/���/����/�ļ����Ȱ���ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 13     'Ƶ��/�·�/����/�ļ����Ȱ��·ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    Case 14     'Ƶ��/���/�·�/����/�ļ����Ȱ���ݣ��ٰ����ڷ֣�ÿ��һ��Ŀ¼��
        GetItemPath = "/" & Year(UpdateTime) & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & Year(UpdateTime) & "-" & Right("0" & Month(UpdateTime), 2) & "-" & Right("0" & Day(UpdateTime), 2) & "/"
    End Select
    GetItemPath = Replace(GetItemPath, "//", "/")
End Function

'**************************************************
'��������GetItemFileName
'��  �ã������Ŀ����
'��  ����iFileNameType ---- �ļ���������
'        sChannelDir ---- ��ǰƵ��Ŀ¼
'        UpdateTime ---- ����ʱ��
'        iArticleID ---- ����ID
'����ֵ�������Ŀ����
'**************************************************
Public Function GetItemFileName(iFileNameType, sChannelDir, UpdateTime, iArticleID)
    Select Case iFileNameType
    Case 0
        GetItemFileName = iArticleID
    Case 1
        GetItemFileName = Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2)
    Case 2
        GetItemFileName = sChannelDir & "_" & iArticleID
    Case 3
        GetItemFileName = sChannelDir & "_" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2)
    Case 4
        GetItemFileName = Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2) & "_" & iArticleID
    Case 5
        GetItemFileName = sChannelDir & "_" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & Right("0" & Day(UpdateTime), 2) & Right("0" & Hour(UpdateTime), 2) & Right("0" & Minute(UpdateTime), 2) & Right("0" & Second(UpdateTime), 2) & "_" & iArticleID
    End Select
End Function

'**************************************************
'��������GetFileExt
'��  �ã�ȡ��Ƶ�������е���չ��
'��  ����FileExtType ---- ȡֵ����
'����ֵ��Ƶ����׺��
'**************************************************
Public Function GetFileExt(FileExtType)
    Select Case FileExtType
    Case 0
        GetFileExt = ".html"
    Case 1
        GetFileExt = ".htm"
    Case 2
        GetFileExt = ".shtml"
    Case 3
        GetFileExt = ".shtm"
    Case 4
        GetFileExt = ".asp"
    End Select
End Function

Function iso8601date(dLocal, utcOffset)
    Dim d, d1
    d = DateAdd("H", -1 * utcOffset, dLocal)
    If Len(utcOffset) < 2 Then
        d1 = "0" & utcOffset
    Else
        d1 = utcOffset
    End If
    iso8601date = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2) & "T"
    iso8601date = iso8601date & (Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2) & ":" & Right("0" & Second(d), 2))
    If utcOffset < 0 Then
        iso8601date = iso8601date & ("-" & d1 & ":00")
    Else
        iso8601date = iso8601date & ("+" & d1 & ":00")
    End If
End Function
%>