<%@language=vbscript codepage=936 %>
<%
Option Explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!-- #include File="../conn.asp" -->
<%
Dim ChannelID,sql,rs,ModuleType
Dim InsertTemplate,InsertTemplateType
ChannelID=PE_CLng(trim(request("ChannelID")))
InsertTemplate=PE_CLng(Request("insertTemplate"))
InsertTemplateType=PE_CLng(Request("InsertTemplateType"))

Public Function PE_CLng(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng = CLng(str1)
    Else
        PE_CLng = 0
    End If
End Function
Public Function PE_CStr(ByVal str1, ByVal str2)
    If IsNull(str1) Then
        PE_CStr = str2
    Else
        PE_CStr = CStr(str1)
    End If
End Function

Private Sub BacktrackEditor() 
    If InsertTemplate=1 Then
        Response.Write  "if(label!=null){" & vbCrLf
        Response.Write "    parent.insertTemplateLabel(label," & InsertTemplateType & ");" & vbCrLf
        Response.Write  "}" & vbCrLf
    Else
        Response.Write "window.returnValue = label" & vbCrLf
        Response.Write "window.close();" & vbCrLf
    End if
End Sub

sql = "select ModuleType from PE_Channel where ChannelID=" & PE_CLng(ChannelID)
Set rs = Conn.Execute(sql)
    If rs.bof And rs.EOF Then
    else
        ModuleType=rs("ModuleType")
    End If
    rs.Close
Set rs = Nothing

'**************************************************
'��������FoundInArr
'��  �ã�����������Ƿ���ָ������ֵ
'��  ����strArr ----- ���������
'        strItem  ----- �����ַ�
'        strSplit  ----- �ָ��ַ�
'����ֵ��True  ----��
'        False ----û��
'**************************************************
Public Function FoundInArr(strArr, strItem, strSplit)
    Dim arrTemp, arrTemp2, i, j
    FoundInArr = False
    If IsNull(strArr) Or IsNull(strItem) Or Trim(strArr) = "" Or Trim(strItem) = "" Then
        Exit Function
    End If
    If IsNull(strSplit) Or strSplit = "" Then
        strSplit = ","
    End If
    If InStr(Trim(strArr), strSplit) > 0 Then
        If InStr(Trim(strItem), strSplit) > 0 Then
            arrTemp = Split(strArr, strSplit)
            arrTemp2 = Split(strItem, strSplit)
            For i = 0 To UBound(arrTemp)
                For j = 0 To UBound(arrTemp2)
                    If LCase(Trim(arrTemp2(j))) <> "" And LCase(Trim(arrTemp(i))) <> "" And LCase(Trim(arrTemp2(j))) = LCase(Trim(arrTemp(i))) Then
                        FoundInArr = True
                        Exit Function
                    End If
                Next
            Next
        Else
            arrTemp = Split(strArr, strSplit)
            For i = 0 To UBound(arrTemp)
                If LCase(Trim(arrTemp(i))) = LCase(Trim(strItem)) Then
                    FoundInArr = True
                    Exit Function
                End If
            Next
        End If
    Else
        If LCase(Trim(strArr)) = LCase(Trim(strItem)) Then
            FoundInArr = True
        End If
    End If
End Function
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <meta name="Keywords" content="��������Ƽ����޹�˾��������վ����ϵͳ�����ף���������ϵͳ������ϵͳ����������վϵͳ��ϵͳ����վ��������վ����ƣ���ҳ��������������������ϵͳ���������װ�����֧�֣���װ����">
    <title>��վ����ϵͳ--��ǩ����</title>
</head>
<body leftmargin="0"  rightmargin="0"topmargin="0">

<!-- ******** �˵�Ч����ʼ ******** -->
<table width="100%"  border="0" cellspacing="0" cellpadding="4" align="center">
  <tr>
    <td align="center" bgcolor="#0066FF"><b><font color="#ffffff">��վ����ϵͳ--��ǩ����</font></b></td>
  </tr>
</table>
<table width="90%"  border="0" cellspacing="0" cellpadding="2" align="center">
  <tr>
    <td height="50" valign="top" background="images/left_tdbg_01.gif">
      <style rel=stylesheet type=text/css>
        td {
        FONT-SIZE: 9pt; COLOR: #000000; FONT-FAMILY: ����,Dotum,DotumChe,Arial;line-height: 150%; 
        }
        INPUT {
        BACKGROUND-COLOR: #ffffff; 
        BORDER-BOTTOM: #666666 1px solid;
        BORDER-LEFT: #666666 1px solid;
        BORDER-RIGHT: #666666 1px solid;
        BORDER-TOP: #666666 1px solid;
        COLOR: #666666;
        HEIGHT: 18px;
        border-color: #666666 #666666 #666666 #666666; font-size: 9pt
        }
        .favMenu {
            BACKGROUND: #ffffff; COLOR: windowtext; CURSOR: hand;line-height: 150%; 
        }
        .favMenu DIV {
            WIDTH: 100%;height: 5px;
        }
        .favMenu A {
            COLOR: windowtext; TEXT-DECORATION: none
        }
        .favMenu A:hover {
            COLOR: windowtext; TEXT-DECORATION: underline
        }
        .topFolder {
            
        }
        .topItem {

        }
        .subFolder {
            PADDING: 0px;BACKGROUND: #ffffff;
        }
        .subItem {
            PADDING: 0px;BACKGROUND: #ffffff;
        }
        .sub {
            BACKGROUND: #ffffff;DISPLAY: none; PADDING: 0px;
        }
        .sub .sub {
            BORDER: 0px;BACKGROUND: #ffffff;
        }
        .icon {
            HEIGHT: 18px; MARGIN-RIGHT: 5px; VERTICAL-ALIGN: absmiddle; WIDTH: 18px
        }
        .outer {
            BACKGROUND: #ffffff;PADDING: 0px;
        }
        .inner {
            BACKGROUND: #ffffff;PADDING: 0px;
        }
        .scrollButton {
            BACKGROUND: #ffffff; BORDER: #f6f6f6 1px solid; PADDING: 0px;
        }
        .flat {
            BACKGROUND: #ffffff; BORDER: #f6f6f6 1px solid; PADDING: 0px;
        }
    </style>
    <SCRIPT type=text/javascript>

    var selectedItem = null;
    var targetWin;

    document.onclick = handleClick;
    document.onmouseover = handleOver;
    document.onmouseout = handleOut;
    document.onmousedown = handleDown;
    document.onmouseup = handleUp;
    document.write(writeSubPadding(10));  

    function handleClick() {
        el = getReal(window.event.srcElement, "tagName", "DIV");
        
        if ((el.className == "topFolder") || (el.className == "subFolder")) {
            el.sub = eval(el.id + "Sub");
            if (el.sub.style.display == null) el.sub.style.display = "none";
            if (el.sub.style.display != "block") { 
                if (el.parentElement.openedSub != null) {
                    var opener = eval(el.parentElement.openedSub + ".opener");
                    ChangeFolderImg(opener,1)
                    hide(el.parentElement.openedSub);
                    if (opener.className == "topFolder")
                        outTopItem(opener);
                }
                el.sub.style.display = "block";
                el.sub.parentElement.openedSub = el.sub.id;
                el.sub.opener = el;
                ChangeFolderImg(el,2)
            }
            else {
                hide(el.sub.id);
                ChangeFolderImg(el,1)
            }
        }    
        if ((el.className == "subItem") || (el.className == "subFolder")) {
            if (selectedItem != null)
                restoreSubItem(selectedItem);
            highlightSubItem(el);
        }
        if ((el.className == "topItem") || (el.className == "topFolder")) {
            if (selectedItem != null)
                restoreSubItem(selectedItem);
        }
        if ((el.className == "topItem") || (el.className == "subItem")) {
            if ((el.href != null) && (el.href != "")) {
                if ((el.target == null) || (el.target == "")) {
                    if (window.opener == null) {
                        if (document.all.tags("BASE").item(0) != null)
                            window.open(el.href, document.all.tags("BASE").item(0).target);
                        else 
                            window.location = el.href;                    
                    }
                    else {
                        window.opener.location =  el.href;
                    }
                }
                else {
                    window.open(el.href, el.target);
                }
            }
        }
        var tmp  = getReal(el, "className", "favMenu");
        if (tmp.className == "favMenu") fixScroll(tmp);
    }
    function handleOver() {
        var fromEl = getReal(window.event.fromElement, "tagName", "DIV");
        var toEl = getReal(window.event.toElement, "tagName", "DIV");
        if (fromEl == toEl) return;
        el = toEl;
        if ((el.className == "topFolder") || (el.className == "topItem")) overTopItem(el);
        if ((el.className == "subFolder") || (el.className == "subItem")) overSubItem(el);
        if ((el.className == "topItem") || (el.className == "subItem")) {
            if (el.href != null) {
                if (el.oldtitle == null) el.oldtitle = el.title;
                if (el.oldtitle != "")
                    el.title = el.oldtitle + "\n" + el.href;
                else
                    el.title = el.oldtitle + el.href;
            }
        }
        if (el.className == "scrollButton") overscrollButton(el);
    }
    function handleOut() {
        var fromEl = getReal(window.event.fromElement, "tagName", "DIV");
        var toEl = getReal(window.event.toElement, "tagName", "DIV");
        if (fromEl == toEl) return;
        el = fromEl;
        if ((el.className == "topFolder") || (el.className == "topItem")) outTopItem(el);
        if ((el.className == "subFolder") || (el.className == "subItem")) outSubItem(el);
        if (el.className == "scrollButton") outscrollButton(el);
    }
    function handleDown() {
        el = getReal(window.event.srcElement, "tagName", "DIV");
        if (el.className == "scrollButton") {
            downscrollButton(el);
            var mark = Math.max(el.id.indexOf("Up"), el.id.indexOf("Down"));
            var type = el.id.substr(mark);
            var menuID = el.id.substring(0,mark);
            eval("scroll" + type + "(" + menuID + ")");
        }
    }
    function handleUp() {
        el = getReal(window.event.srcElement, "tagName", "DIV");
        if (el.className == "scrollButton") {
            upscrollButton(el);
            window.clearTimeout(scrolltimer);
        }
    }
    ////////////////////// EVERYTHING IS HANDLED ////////////////////////////
    function hide(elID) {
        var el = eval(elID);
        el.style.display = "none";
        el.parentElement.openedSub = null;
    }
    function writeSubPadding(depth) {
        var str, str2, val;
        var str = "<style type='text/css'>\n";
        for (var i=0; i < depth; i++) {
            str2 = "";
            val  = 0;
            for (var j=0; j < i; j++) {
                str2 += ".sub "
                val += 18;    //����Ŀ��߾���ֵ
            }
            str += str2 + ".subFolder {padding-left: " + val + "px;}\n";
            str += str2 + ".subItem   {padding-left: " + val + "px;}\n";
        }
        str += "</style>\n";
        return str;
    }
    function overTopItem(el) {
        with (el.style) {
            background   = "#f8f8f8";
            paddingBottom = "0px";
        }
    }
    function outTopItem(el) {
        if ((el.sub != null) && (el.parentElement.openedSub == el.sub.id)) { 
            with(el.style) {
                background = "#ffffff";
            }
        }
        else {
            with (el.style) {
                background = "#ffffff";
                padding = "0px";
            }
        }
    }
    function overSubItem(el) {
            el.style.background = "#F6F6F6";
        el.style.textDecoration = "underline";
    }
    function outSubItem(el) {
                el.style.background = "#ffffff";
        el.style.textDecoration = "none";
    }
    function highlightSubItem(el) {
        el.style.background = "#ffffff";
        el.style.color      = "#ff0000"; 
        selectedItem = el;
    }
    function restoreSubItem(el) {
        el.style.background = "#ffffff";
        el.style.color      = "menutext";
        selectedItem = null;
    }
    function overscrollButton(el) {
        overTopItem(el);
        el.style.padding = "0px";
    }
    function outscrollButton(el) {
        outTopItem(el);
        el.style.padding = "0px";
    }
    function downscrollButton(el) {
        with (el.style) {
            borderRight   = "0px solid buttonhighlight";
            borderLeft  = "0px solid buttonshadow";
            borderBottom    = "0px solid buttonhighlight";
            borderTop = "0px solid buttonshadow";
        }
    }
    function upscrollButton(el) {
        overTopItem(el);
        el.style.padding = "0px";
    }
    function getReal(el, type, value) {
        var temp = el;
        while ((temp != null) && (temp.tagName != "BODY")) {
            if (eval("temp." + type) == value) {
                el = temp;
                return el;
            }
            temp = temp.parentElement;
        }
        return el;
    }
    ////////////////////////////////////////////////////////////////////////////////////////
    // Fix the scrollbars
    var globalScrollContainer;    
    var overflowTimeout = 1;

    function fixScroll(el) {
        globalScrollContainer = el;
        window.setTimeout('changeOverflow(globalScrollContainer)', overflowTimeout);
    }
    function changeOverflow(el) {
        if (el.offsetHeight > el.parentElement.clientHeight)
            window.setTimeout('globalScrollContainer.parentElement.style.overflow = "auto";', overflowTimeout);
        else
            window.setTimeout('globalScrollContainer.parentElement.style.overflow = "hidden";', overflowTimeout);
    }
    function ChangeFolderImg(el,type) {
        var FolderImg = eval(el.id + "Img");
        if (type == 1) {
            FolderImg.src="images/foldericon1.gif"
        }
        else {
            FolderImg.src="images/foldericon2.gif"
        }
    }
    ////////////////////////////////////////////////////////////////////////////////////////
    // ��ǩ����
    //��ͨ��ǩ
    function InsertLabel(label){
    <%
      Call BacktrackEditor()
    %>
    }
    //������ǩ
    function InsertAdjs(type,fiflepath){
        var str="";
        switch(type){
        case "SwitchFont":
            str="<a name=StranLink href=''>�л������w����</a>"
            break;
        case "Adjs":
            break;
        default:
            alert("����������ã�");
            break;
       }
    <%  If InsertTemplate=1 Then %>
            str=str+"<"+"script language=\"javascript\" src=\""+fiflepath+"\"></"+"script>"
            
    <%  Else %>
            str=str+"<IMG alt='#[!"+"script language=\"javascript\" src=\""+fiflepath+"\"!][!/"+"script!]#'  src=\"editor/images/jscript.gif\" border=0 $>"
    <%
        End if
        If InsertTemplate=1 Then
            Response.Write "parent.insertTemplateLabel(str," & InsertTemplateType & ");" & vbCrLf
        Else
            Response.write "window.returnValue =str" & vbCrLf
        End if
    %>
       window.close();
    }
    //������ǩ����
    function FunctionLabel(url,width,height){
        var label = showModalDialog(url, "", "dialogWidth:"+width+"px; dialogHeight:"+height+"px; help: no; scroll:no; status: no"); 
        <%
          Call BacktrackEditor()
        %>
    }
    //�����Ա�ǩ
    function FunctionLabel2(name){
        var str,label
        switch(name){
        case "ShowTopUser":
            str=prompt("��������ʾע���û��б������.","5"); 
            label="{$"+name+"("+str+")}"
            break;
        case "��ArticleList_ChildClass��":
            str=prompt("ѭ����ʾ������Ŀ¼�б�ÿ����ʾ������","2"); 
                if (str!=null) {
            label=name+"��Cols="+str+"��{$rsClass_ClassUrl} ��Ŀ��¼������Ŀ��ַ {$rsClass_Readme} ˵�� {$rsClass_ClassName}����  ���������������Զ���ı�ǩ��/ArticleList_ChildClass��"
            }
            break;
        case "��SoftList_ChildClass��":
            str=prompt("ѭ����ʾ������Ŀ¼�б�ÿ����ʾ������","2"); 
                if (str!=null) {
            label=name+"��Cols="+str+"��{$rsClass_ClassUrl} ��Ŀ��¼������Ŀ��ַ {$rsClass_Readme} ˵�� {$rsClass_ClassName}����  ���������������Զ���ı�ǩ��/SoftList_ChildClass��"
            }
            break;
        case "��PhotoList_ChildClass��":
            str=prompt("ѭ����ʾͼƬ��Ŀ¼�б�ÿ����ʾ������","2"); 
                if (str!=null) {
            label=name+"��Cols="+str+"��{$rsClass_ClassUrl} ��Ŀ��¼������Ŀ��ַ {$rsClass_Readme} ˵�� {$rsClass_ClassName}����  ���������������Զ���ı�ǩ��/PhotoList_ChildClass��"
            }
            break;
        case "��PositionList_Content��":
            str=prompt("ѭ����ʾְλ������Ϣ�б�ÿҳ��ʾ��ְλ��","6");
                if (str!=null) {
            label = name + "��PerPageNum=" + str + "��˵�������ڴ˼����˲���Ƹ���ݱ�ǩ��������ְλ��ť��ǩ{$SaveSupply}����/PositionList_Content��"
            }
            break;
        case "DownloadUrl":
            str=prompt("һ����ʾ������","3");
                if (str!=null) {
            label = "{$"+name + "(" + str + ")}"
            }
            break;
        case "ResumeError":
            label = "\n<"+"SCRIPT LANGUAGE='JavaScript'>"
            label += "\n<!--"
            label += "\n function ResumeError() {"
            label += "\n return true;"
            label += "\n }"
            label += "\n window.onerror = ResumeError;"
            label += "\n // -->"
            label += "\n</"+"SCRIPT>"
            break;
        default:
            alert("����������ã�");
            break;
        }
        <%
          Call BacktrackEditor()
        %>
    }
    //��̬�����Ա�ǩ
    function FunctionLabel3(name){
        str=prompt("�����붯̬������ǩ����.","5"); 
        label="{$"+name+"("+str+")}"
        <%
          Call BacktrackEditor()
        %>
    }
    //����������ǩ 
    function SuperFunctionLabel (url,label,title,ModuleType,ChannelShowType,iwidth,iheight){
        var label = showModalDialog(url+"?ChannelID=<%=ChannelID%>&Action="+label+"&Title="+title+"&ModuleType="+ModuleType+"&ChannelShowType="+ChannelShowType+"&InsertTemplate=<%=InsertTemplate%>", "", "dialogWidth:"+iwidth+"px; dialogHeight:"+iheight+"px; help: no; scroll:yes; status: yes"); 
        <%
          Call BacktrackEditor()
        %>
    }      
    </SCRIPT>
    <!-- ��ҳ -->
    <DIV class=topItem>
      <IMG class=icon height=16 src="images/home.gif" style="HEIGHT: 16px">��ǩ����
    </DIV>
    <!-- ϵͳ���� -->
    <DIV class=favMenu id=aMenu>
    <!-- ͨ�ñ�ǩ -->
    <DIV class=topFolder id=web><IMG id=webImg class=icon src="images/foldericon1.gif">��վͨ�ñ�ǩ</DIV>
    <DIV class=sub id=webSub>
        <!-- ��վͨ����ͨ��ǩ -->
        <DIV class=subFolder id=subwebInsert><IMG id=subwebInsertImg class=icon src="images/foldericon1.gif"> ��վͨ�ñ�ǩ</DIV>
        <DIV class=sub id=subwebInsertSub>
            <DIV class=subItem onclick="InsertLabel('{$PageTitle}')"><IMG class=icon src="images/label.gif">��ʾ������ı�������ʾҳ��ı�����Ϣ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowChannel}')"><IMG class=icon src="images/label.gif">��ʾ����Ƶ����Ϣ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowPath}')"><IMG class=icon src="images/label.gif">��ʾ������Ϣ</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ShowVote}')"><IMG class=icon src="images/label.gif">��ʾ��վ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SiteName}')"><IMG class=icon src="images/label.gif">��ʾ��վ����</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$SiteUrl}')"><IMG class=icon src="images/label.gif">��ʾ��վ��ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$InstallDir}')"><IMG class=icon src="images/label.gif">ϵͳ��װĿ¼</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowAdminLogin}')"><IMG class=icon src="images/label.gif">��ʾ�����¼������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Copyright}')"><IMG class=icon src="images/label.gif">��ʾ��Ȩ��Ϣ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Meta_Keywords}')"><IMG class=icon src="images/label.gif">�����������Ĺؼ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Meta_Description}')"><IMG class=icon src="images/label.gif">������������˵��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowSiteCountAll}')"><IMG class=icon src="images/label.gif">��ʾ����ע���Ա</DIV>
            <DIV class=subItem onclick="InsertLabel('{$WebmasterName}')"><IMG class=icon src="images/label.gif">��ʾվ������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$WebmasterEmail}')"><IMG class=icon src="images/label.gif">��ʾվ��Email����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$MenuJS}')"><IMG class=icon src="images/label.gif">������ĿJS����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Skin_CSS}')"><IMG class=icon src="images/label.gif">���CSS</DIV>
        </DIV>
        <!-- ��վͨ�ú�����ͨ��ǩ���ٱ�ǩ -->
        <DIV class=subFolder id=subwebFunction><IMG id=subwebFunctionImg class=icon src="images/foldericon1.gif"> ��վͨ�ú�����ǩ</DIV>
        <DIV class=sub id=subwebFunctionSub>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Logo.htm','240','140')"><IMG class=icon src="images/label2.gif">��ʾ��վLOGO</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Banner.htm','240','140')"><IMG class=icon src="images/label2.gif">��ʾ��վBanner</DIV>            
            <DIV class=subItem onclick="FunctionLabel2('ShowTopUser')"><IMG class=icon src="images/label2.gif">��ʾע���û��б�</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Annouce.htm','240','140')"><IMG class=icon src="images/label2.gif">��ʾ��վ������Ϣ</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Annouce2.htm','240','210')"><IMG class=icon src="images/label2.gif">��ʾ��ϸ������Ϣ</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_FSite.htm','330','260')"><IMG class=icon src="images/label2.gif">��ʾ����������Ϣ</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_FSite2.htm','330','510')"><IMG class=icon src="images/label2.gif">��ʾ��ϸ����������Ϣ</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_ProducerList.htm','400','450')"><IMG class=icon src="images/label2.gif">��ʾ�����б�</DIV> 
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Author_ShowList.htm','400','340')"><IMG class=icon src="images/label2.gif">��ʾ�����б�</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_ShowSpecialList.htm','320','300')"><IMG class=icon src="images/label2.gif">��ʾָ��Ƶ��ר��</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_ShowBlogList.htm','400','400')"><IMG class=icon src="images/label2.gif">��ʾ��Ʒ������</DIV>
        </DIV>
    </DIV>
    <!-- Ƶ��ͨ�ñ�ǩ -->
    <DIV class=topFolder id=ChannelCommon><IMG id=ChannelCommonImg class=icon src="images/foldericon1.gif">Ƶ��ͨ�ñ�ǩ</DIV>
    <DIV class=sub id=ChannelCommonSub>
        <DIV class=subItem onclick="InsertLabel('{$ChannelName}')"><IMG class=icon src="images/label.gif">��ʾƵ������</DIV>    
        <DIV class=subItem onclick="InsertLabel('{$ChannelID}')"><IMG class=icon src="images/label.gif">�õ�Ƶ��ID</DIV>    
        <DIV class=subItem onclick="InsertLabel('{$ChannelDir}')"><IMG class=icon src="images/label.gif">�õ�Ƶ��Ŀ¼</DIV>    
        <DIV class=subItem onclick="InsertLabel('{$ChannelUrl}')"><IMG class=icon src="images/label.gif">Ƶ��Ŀ¼·��</DIV>
        <DIV class=subItem onclick="InsertLabel('{$UploadDir}')"><IMG class=icon src="images/label.gif">Ƶ���ϴ�Ŀ¼·��</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ChannelPicUrl}')"><IMG class=icon src="images/label.gif">Ƶ��ͼƬ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Meta_Keywords_Channel}')"><IMG class=icon src="images/label.gif">�����������Ĺؼ���</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Meta_Description_Channel}')"><IMG class=icon src="images/label.gif">������������˵��</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ChannelShortName}')"><IMG class=icon src="images/label.gif">��ʾƵ����</DIV>    
        <DIV class=subItem onclick="FunctionLabel('Lable/PE_ClassNavigation.htm','260','200')"><IMG class=icon src="images/label2.gif">��ʾ��Ŀ������HTML����</DIV>
    </DIV>
    <!-- Ƶ��ר��ҳ��ǩ -->
    <DIV class=topFolder id=Channel><IMG id=ChannelImg class=icon src="images/foldericon1.gif">Ƶ��ר�ñ�ǩ</DIV>
    <DIV class=sub id=ChannelSub>
        <DIV class=subItem onclick="FunctionLabel('Lable/PE_AnnWin.htm','240','140')"><IMG class=icon src="images/label2.gif">�������洰��</DIV>    
        <DIV class=subItem onclick="InsertLabel('{$ClassListUrl}')"><IMG class=icon src="images/label.gif">ģ���С����ࡱ��������</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ShowChildClass}')"><IMG class=icon src="images/label.gif">��ʾһ����Ŀ�µڶ�����Ŀ��</DIV>
        <DIV class=subItem onclick="FunctionLabel('Lable/PE_ShowChildClass.htm','330','360')"><IMG class=icon src="images/label2.gif">��ʾ��ǰ��Ŀ����һ������Ŀ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ParentDir}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ�ĸ�Ŀ¼</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ClassDir}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ��Ŀ¼</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Readme}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ��˵��</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ClassUrl}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ�����ӵ�ַ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ClassPicUrl}')"><IMG class=icon src="images/label.gif">��ĿͼƬ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Meta_Keywords_Class}')"><IMG class=icon src="images/label.gif">�����������Ĺؼ���</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Meta_Description_Class}')"><IMG class=icon src="images/label.gif">������������˵��</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ClassName}')"><IMG class=icon src="images/label.gif">��ʾ��ǰ��Ŀ������</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ClassID}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ��ID</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ShowChannelCount}')"><IMG class=icon src="images/label.gif">��ʾƵ��ͳ����Ϣ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$SpecialName}')"><IMG class=icon src="images/label.gif">��ʾ��ǰר���ר������</DIV>
        <DIV class=subItem onclick="InsertLabel('{$SpecialPicUrl}')"><IMG class=icon src="images/label.gif">��ʾר��ͼƬ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$GetAllSpecial}')"><IMG class=icon src="images/label.gif">��ʾȫ��ר��</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ShowPage}')"><IMG class=icon src="images/label.gif">��ʾ��ҳ��ǩ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ShowPage_en}')"><IMG class=icon src="images/label.gif">��ʾӢ�ķ�ҳ��ǩ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$InstallDir}{$ChannelDir}')"><IMG class=icon src="images/label.gif">Ƶ����װĿ¼</DIV>
    </DIV>
    <!-- Ƶ������ҳ��ǩ -->
    <DIV class=topFolder id=ChannelSearch><IMG id=ChannelSearchImg class=icon src="images/foldericon1.gif">Ƶ������ҳ��ǩ</DIV>
    <DIV class=sub id=ChannelSearchSub>
        <DIV class=subItem onclick="InsertLabel('{$ResultTitle}')"><IMG class=icon src="images/label.gif">��ʾ��������ʲô������Ϣ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$SearchResult}')"><IMG class=icon src="images/label.gif">�������</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Keyword}')"><IMG class=icon src="images/label.gif">�����ؼ���</DIV>
    </DIV>
    <!-- ����ҳͨ�ñ�ǩ -->
    <DIV class=topFolder id=ContentCommon><IMG id=ContentCommonImg class=icon src="images/foldericon1.gif">����ҳͨ�ñ�ǩ</DIV>
    <DIV class=sub id=ContentCommonSub>
        <DIV class=subItem onclick="InsertLabel('{$ClassID}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ��ID</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ClassName}')"><IMG class=icon src="images/label.gif">��ʾ��ǰ��Ŀ������</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ClassDir}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ��Ŀ¼</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Readme}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ��˵��</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ClassUrl}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ�����ӵ�ַ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ParentDir}')"><IMG class=icon src="images/label.gif">�õ���ǰ��Ŀ�ĸ�Ŀ¼</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Keyword}')"><IMG class=icon src="images/label.gif">�����ؼ���</DIV>
    </DIV>

<% if ModuleType=1 or ModuleType=0 then %>
    <!-- ����Ƶ����ǩ -->
     <DIV class=topFolder id=Article><IMG id=ArticleImg class=icon src="images/foldericon1.gif">���±�ǩ</DIV>
     <DIV class=sub id=ArticleSub>
        <!-- ����ͨ��Ƶ����ǩ -->
        <DIV class=subFolder id=subArticleChannelFunction><IMG id=subArticleChannelFunctionImg class=icon src="images/foldericon1.gif"> ����Ƶ����ǩ</DIV>
        <DIV class=sub id=subArticleChannelFunctionSub>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetArticleList','�����б�����ǩ',1,'GetList',800,700)"><IMG class=icon src="images/label3.gif">��ʾ���±������Ϣ</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetPicArticle','��ʾͼƬ���±�ǩ',1,'GetPic',700,500)"><IMG class=icon src="images/label3.gif">��ʾͼƬ����</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetSlidePicArticle','��ʾ�õ�Ƭ���±�ǩ',1,'GetSlide',700,500)"><IMG class=icon src="images/label3.gif">��ʾ�õ�Ƭ����</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_CustomListLabel.asp','CustomListLable','�����Զ����б��ǩ',1,'GetArticleCustom',720,700)"><IMG class=icon src="images/label3.gif">�����Զ����б��ǩ</DIV>
        </Div>
        <DIV class=subFolder id=subArticleClass><IMG id=subArticleClassImg class=icon src="images/foldericon1.gif"> ������Ŀ��ǩ</DIV>
        <DIV class=sub id=subArticleClassSub>
            <DIV class=subItem onclick="FunctionLabel2('��ArticleList_ChildClass��')"><IMG class=icon src="images/label2.gif">ѭ����ʾ������Ŀ¼�б�</DIV> 
            <DIV class=subItem onclick="InsertLabel('��ArticleList_CurrentClass��{$rsClass_ClassUrl} ��Ŀ��¼������Ŀ��ַ {$rsClass_Readme}˵�� {$rsClass_ClassName}����  ���������������Զ���ı�ǩ��/ArticleList_CurrentClass��')"><IMG class=icon src="images/label.gif">��ǰ��Ŀ�б�(ͬʱ�������¼�����Ŀ)ѭ����ǩ</DIV>
        </DIV>
         <!-- ����ͨ��Ƶ����ǩ���� -->
         <!-- ����Ƶ�����ݱ�ǩ -->
         <DIV class=subFolder id=subArticleChannelContent><IMG id=subArticleChannelContentImg class=icon src="images/foldericon1.gif"> �������ݱ�ǩ</DIV>
         <DIV class=sub id=subArticleChannelContentSub>
            <DIV class=subItem onclick="InsertLabel('{$ArticleID}')"><IMG class=icon src="images/label.gif">��ǰ���µ�I D</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ArticleProtect}')"><IMG class=icon src="images/label.gif">����Ƶ�����õõ������ƹ��ܵĴ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ArticleProperty}')"><IMG class=icon src="images/label.gif">��ʾ��ǰ���µ����ԣ����š��Ƽ����ȼ���</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$ArticleTitle}')"><IMG class=icon src="images/label.gif">��ʾ����������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ArticleTitle2}')"><IMG class=icon src="images/label.gif">��ʾ������ʾҳ��������ǰ���±�����Ϣ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ArticleInfo}')"><IMG class=icon src="images/label.gif">������ʾ�������ߡ�������Դ�������������ʱ����Ϣ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ArticleSubheading}')"><IMG class=icon src="images/label.gif">��ʾ���¸�����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Subheading}')"><IMG class=icon src="images/label.gif">�Զ����б�����</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ReadPoint}')"><IMG class=icon src="images/label.gif">�Ķ�����</DIV>            
            <DIV class=subItem onclick="InsertLabel('{$Author}')"><IMG class=icon src="images/label.gif">����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$CopyFrom}')"><IMG class=icon src="images/label.gif">������Դ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Editor}')"><IMG class=icon src="images/label.gif">���α༭</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">�����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">����ʱ����Ϣ</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ArticleIntro}')"><IMG class=icon src="images/label.gif">��ʾ���¼��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ArticleContent}')"><IMG class=icon src="images/label.gif">��ʾ���µľ��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PrevArticle}')"><IMG class=icon src="images/label.gif">��ʾ��һƪ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$NextArticle}')"><IMG class=icon src="images/label.gif">��ʾ��һƪ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ArticleEditor}')"><IMG class=icon src="images/label.gif">��ʾ����¼������α༭����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ArticleAction}')"><IMG class=icon src="images/label.gif">��ʾ���������ۡ������ߺ��ѡ�����ӡ���ġ����رմ��ڡ�</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_CorrelativeArticle.htm','240','140')"><IMG class=icon src="images/label2.gif">��ʾ�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ManualPagination}')"><IMG class=icon src="images/label.gif">�����ֶ���ҳ��ʽ��ʾ���¾��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AutoPagination}')"><IMG class=icon src="images/label.gif">�����Զ���ҳ��ʽ��ʾ���¾��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Vote}')"><IMG class=icon src="images/label.gif">��ʾ����</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_GetSubTitleHtml.htm','340','200')"><IMG class=icon src="images/label2.gif">���ķ�ҳ����</DIV>
        </DIV>
        <!-- ����Ƶ�����ݱ�ǩ���� -->
    </DIV>
    <!-- ����Ƶ����ǩ���� -->
    <%
    End if
    if  ModuleType=2 or ModuleType=0 then %>
    <!-- ����Ƶ����ǩ -->
    <DIV class=topFolder id=Soft><IMG id=SoftImg class=icon src="images/foldericon1.gif">���ر�ǩ</DIV>
    <DIV class=sub id=SoftSub>
         <!-- ����ͨ��Ƶ����ǩ -->
         <DIV class=subFolder id=subSoftChannelFunction><IMG id=subSoftChannelFunctionImg class=icon src="images/foldericon1.gif"> ����Ƶ����ǩ</DIV>
         <DIV class=sub id=subSoftChannelFunctionSub>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetSoftList','�����б�����ǩ',2,'GetList',800,700)"><IMG class=icon src="images/label3.gif">��ʾ����������Ϣ</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetPicSoft','��ʾͼƬ���ر�ǩ',2,'GetPic',700,500)"><IMG class=icon src="images/label3.gif">��ʾͼƬ����</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetSlidePicSoft','��ʾ�õ�Ƭ���ر�ǩ',2,'GetSlide',700,500)"><IMG class=icon src="images/label3.gif">��ʾ�õ�Ƭ����</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_CustomListLabel.asp','CustomListLable','�����Զ����б��ǩ',2,'GetSoftCustom',720,700)"><IMG class=icon src="images/label3.gif">�����Զ����б��ǩ</DIV>
        </DIV>
        <DIV class=subFolder id=subSoftClassFunction><IMG id=subSoftClassFunctionImg class=icon src="images/foldericon1.gif"> ������Ŀ��ǩ</DIV>
        <DIV class=sub id=subSoftClassFunctionSub>
            <DIV class=subItem onclick="FunctionLabel2('��SoftList_ChildClass��')"><IMG class=icon src="images/label2.gif"> ѭ����ʾ������Ŀ¼�б�</DIV>
            <DIV class=subItem onclick="InsertLabel('��SoftList_CurrentClass��{$rsClass_ClassUrl} ��Ŀ��¼������Ŀ��ַ {$rsClass_Readme}˵�� {$rsClass_ClassName}����  ���������������Զ���ı�ǩ��/SoftList_CurrentClass��')"><IMG class=icon src="images/label.gif">��ǰ��Ŀ�б�(ͬʱ�������ؼ�����Ŀ)ѭ����ǩ</DIV>
        </DIV>
        <!-- ����ͨ��Ƶ����ǩ���� -->
        <!-- ����Ƶ�����ݱ�ǩ -->
        <DIV class=subFolder id=subSoftChannelContent><IMG id=subSoftChannelContentImg class=icon src="images/foldericon1.gif"> �������ݱ�ǩ</DIV>
        <DIV class=sub id=subSoftChannelContentSub>
            <DIV class=subItem onclick="InsertLabel('{$SoftID}')"><IMG class=icon src="images/label.gif">���ID</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftName}')"><IMG class=icon src="images/label.gif">�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftVersion}')"><IMG class=icon src="images/label.gif">��ʾ����汾</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftSize} K')"><IMG class=icon src="images/label.gif">����ļ���С</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftSize_M}')"><IMG class=icon src="images/label.gif">��ʾ�����С����M Ϊ��λ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$DecompressPassword}')"><IMG class=icon src="images/label.gif">��ѹ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$OperatingSystem}')"><IMG class=icon src="images/label.gif">����ƽ̨</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">���ش����ܼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Author}')"><IMG class=icon src="images/label.gif">�� �� ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$DayHits}')"><IMG class=icon src="images/label.gif">���ش�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$WeekHits}')"><IMG class=icon src="images/label.gif">���ش�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$MonthHits}')"><IMG class=icon src="images/label.gif">���ش�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Stars}')"><IMG class=icon src="images/label.gif">����ȼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$CopyFrom}')"><IMG class=icon src="images/label.gif">�����Դ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftLink}')"><IMG class=icon src="images/label.gif">��ʾ�������ʾ��ַ��ע���ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftType}')"><IMG class=icon src="images/label.gif">������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftLanguage}')"><IMG class=icon src="images/label.gif">�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftProperty}')"><IMG class=icon src="images/label.gif">�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">����ʱ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Editor}')"><IMG class=icon src="images/label.gif">���������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Inputer}')"><IMG class=icon src="images/label.gif">������¼��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftPicUrl}')"><IMG class=icon src="images/label.gif">��ʾ����ͼƬ</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_SoftPic.htm','240','140')"><IMG class=icon src="images/label2.gif">��ʾ����ͼƬ��ϸ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$DemoUrl}')"><IMG class=icon src="images/label.gif">��ʾ��ʾ��ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$RegUrl}')"><IMG class=icon src="images/label.gif">��ʾע���ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftPoint}')"><IMG class=icon src="images/label.gif">��ʾ�շ���������ص���</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$CopyrightType}')"><IMG class=icon src="images/label.gif">��Ȩ��ʽ</DIV>    
            <DIV class=subItem onclick="FunctionLabel2('DownloadUrl')"><IMG class=icon src="images/label2.gif">������ص�ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftIntro}')"><IMG class=icon src="images/label.gif">������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$CorrelativeSoft}')"><IMG class=icon src="images/label.gif">������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SoftAuthor}')"><IMG class=icon src="images/label.gif">��ʾ�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorEmail}')"><IMG class=icon src="images/label.gif">��ʾ����Email</DIV>
            <DIV class=subItem onclick="InsertLabel('{$BrowseTimes}')"><IMG class=icon src="images/label.gif">��ʾ����������</DIV>
        </DIV>
        <!-- ����Ƶ�����ݱ�ǩ���� -->
    </DIV>
    <%
    End if
    If  ModuleType=3 or ModuleType=0 then %>
    <!-- ͼƬƵ����ǩ -->
     <DIV class=topFolder id=Photo><IMG id=PhotoImg class=icon src="images/foldericon1.gif">ͼƬ��ǩ</DIV>
     <DIV class=sub id=PhotoSub>
        <!-- ͼƬͨ��Ƶ����ǩ -->
        <DIV class=subFolder id=subPhotoChannelFunction><IMG id=subPhotoChannelFunctionImg class=icon src="images/foldericon1.gif"> ͼƬƵ����ǩ</DIV>
        <DIV class=sub id=subPhotoChannelFunctionSub>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetPhotoList','ͼƬ�б�����ǩ',3,'GetList',800,700)"><IMG class=icon src="images/label3.gif">��ʾͼƬ�������Ϣ</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetPicPhoto','��ʾͼƬͼ�ı�ǩ',3,'GetPic',700,550)"><IMG class=icon src="images/label3.gif">��ʾͼƬ</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetSlidePicPhoto','��ʾ�õ�ƬͼƬ��ǩ',3,'GetSlide',700,550)"><IMG class=icon src="images/label3.gif">��ʾ�õ�ƬͼƬ</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_CustomListLabel.asp','CustomListLable','ͼƬ�Զ����б��ǩ',3,'GetPhotoCustom',720,700)"><IMG class=icon src="images/label3.gif">ͼƬ�Զ����б��ǩ</DIV>
        </DIV>
        <DIV class=subFolder id=subPhotoClassFunction><IMG id=subPhotoClassFunctionImg class=icon src="images/foldericon1.gif"> ͼƬ��Ŀ��ǩ</DIV>
        <DIV class=sub id=subPhotoClassFunctionSub>
            <DIV class=subItem onclick="FunctionLabel2('��PhotoList_ChildClass��')"><IMG class=icon src="images/label2.gif">ѭ����ʾͼƬ��Ŀ¼�б�</DIV>
            <DIV class=subItem onclick="InsertLabel('��PhotoList_CurrentClass��{$rsClass_ClassUrl} ��Ŀ��¼������Ŀ��ַ {$rsClass_Readme}˵�� {$rsClass_ClassName}����  ���������������Զ���ı�ǩ��/PhotoList_CurrentClass��')"><IMG class=icon src="images/label.gif">��ǰ��Ŀ�б�(ͬʱ����ͼƬ������Ŀ)ѭ����ǩ</DIV>
        </DIV>
        <!-- ͼƬƵ��ͨ�ñ�ǩ���� -->
        <!-- ͼƬƵ�����ݱ�ǩ -->
        <DIV class=subFolder id=subPhotoChannelContent><IMG id=subPhotoChannelContentImg class=icon src="images/foldericon1.gif"> ͼƬ���ݱ�ǩ</DIV>
        <DIV class=sub id=subPhotoChannelContentSub>
            <DIV class=subItem onclick="InsertLabel('{$PhotoID}')"><IMG class=icon src="images/label.gif">ͼƬI D</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PhotoName}')"><IMG class=icon src="images/label.gif">��ʾͼƬ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">�鿴�����ܼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Author}')"><IMG class=icon src="images/label.gif">ͼƬ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$CopyFrom}')"><IMG class=icon src="images/label.gif">ͼƬ��Դ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PhotoProperty}')"><IMG class=icon src="images/label.gif">��ʾͼƬ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Stars}')"><IMG class=icon src="images/label.gif">�Ƽ��ȼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">����ʱ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Editor}')"><IMG class=icon src="images/label.gif">��ʾͼƬ�����α༭</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Inputer}')"><IMG class=icon src="images/label.gif">��ʾͼƬ¼����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PhotoPoint}')"><IMG class=icon src="images/label.gif">�շ�ͼƬ����ĵ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PhotoIntro}')"><IMG class=icon src="images/label.gif">��ʾͼƬ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PrevPhotoUrl}')"><IMG class=icon src="images/label.gif">��һ��ͼƬ�����ӵ�ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$NextPhotoUrl}')"><IMG class=icon src="images/label.gif">��һ��ͼƬ�����ӵ�ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ViewPhoto}')"><IMG class=icon src="images/label.gif">��ʾͼƬ��Flash</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_PhotoUrlList.htm','300','270')"><IMG class=icon src="images/label2.gif">��ʾͼƬ��ַ�б�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PhotoUrl}')"><IMG class=icon src="images/label.gif">ͼƬ��ַ�б��еĵ�һ����ַ</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_CorrelativePhoto.htm','240','140')"><IMG class=icon src="images/label2.gif">���ͼƬ�б�</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$PhotoSize} K')"><IMG class=icon src="images/label.gif">ͼƬ��С</DIV>
            <DIV class=subItem onclick="InsertLabel('{$DayHits}')"><IMG class=icon src="images/label.gif">�鿴��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$WeekHits}')"><IMG class=icon src="images/label.gif">�鿴��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$MonthHits}')"><IMG class=icon src="images/label.gif">�鿴��������</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$PhotoThumb}')"><IMG class=icon src="images/label.gif">��ʾͼƬ����ͼ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GetUrlArray}')"><IMG class=icon src="images/label.gif">��ȡͼƬ��ַ�ĳ�ʼ��JS</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_PhotoThumb.htm','240','140')"><IMG class=icon src="images/label2.gif">��ʾָ����С��ͼƬ����ͼ</DIV>
        </DIV>
        <!-- ͼƬƵ�����ݱ�ǩ���� -->
     </DIV>
    <%
    End if
    if  ModuleType=4 or ModuleType=0 then %>
    <!--  ����Ƶ������  -->
     <DIV class=topFolder id=Guest><IMG id=GuestImg class=icon src="images/foldericon1.gif">���Ժ���</DIV>
     <DIV class=sub id=GuestSub>
        <!-- ���԰�ͨ�ñ�ǩ -->
        <DIV class=subFolder id=subGuestCommonFunction><IMG id=subGuestCommonFunctionImg class=icon src="images/foldericon1.gif">���԰�ͨ�ñ�ǩ</DIV>
        <DIV class=sub id=subGuestCommonFunctionSub>
            <DIV class=subItem onclick="InsertLabel('{$GetGKindList}')"><IMG class=icon src="images/label.gif">��ʾ���������򵼺�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GuestBook_top}')"><IMG class=icon src="images/label.gif">��ʾ�������ܲ˵�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GuestBook_Mode}')"><IMG class=icon src="images/label.gif">��ʾ����ģʽ���ο�/ ��Աģʽ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GuestBook_See}')"><IMG class=icon src="images/label.gif">��ʾ���Բ鿴ģʽ�����԰�/ ������ģʽ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GuestBook_Appear}')"><IMG class=icon src="images/label.gif">��ʾ���Է���ģʽ����˷���/ ֱ�ӷ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowGueststyle}')"><IMG class=icon src="images/label.gif">�л�����һ�ֲ鿴��ʽ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GuestBook_Search}')"><IMG class=icon src="images/label.gif">��ʾ����������</DIV>
        </DIV>
        <!-- ���԰�ͨ�ñ�ǩ -->
        <DIV class=subFolder id=subGuestIndexFunction><IMG id=subGuestIndexFunctionImg class=icon src="images/foldericon1.gif">���԰�ͨ�ñ�ǩ</DIV>
        <DIV class=sub id=subGuestIndexFunctionSub>
            <DIV class=subItem onclick="InsertLabel('{$GuestMain}')"><IMG class=icon src="images/label.gif">��ʾ�����б�</DIV>    
        </DIV>
        <!-- �༭����ҳ��ǩ -->
        <DIV class=subFolder id=subGuestEditFunction><IMG id=subGuestEditFunctionImg class=icon src="images/foldericon1.gif">�༭����ҳ��ǩ</DIV>
        <DIV class=sub id=subGuestEditFunctionSub>
            <DIV class=subItem onclick="InsertLabel('{$WriteGuest}')"><IMG class=icon src="images/label.gif">ǩд����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowJS_Guest}')"><IMG class=icon src="images/label.gif">����Js��֤</DIV>
            <DIV class=subItem onclick="InsertLabel('{$WriteTitle}')"><IMG class=icon src="images/label.gif">��ʾ���Ա���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GetGKind_Option}')"><IMG class=icon src="images/label.gif">��ʾ�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GuestFace}')"><IMG class=icon src="images/label.gif">��ʾ��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GuestContent}')"><IMG class=icon src="images/label.gif">��ʾ��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$saveedit}')"><IMG class=icon src="images/label.gif">����Ƿ�Ϊ�༭����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ReplyId}')"><IMG class=icon src="images/label.gif">�ظ�����id</DIV>
            <DIV class=subItem onclick="InsertLabel('{$saveeditid}')"><IMG class=icon src="images/label.gif">Ҫ�༭���Ե�ID</DIV>
        </DIV>
        <!-- ���Իظ�ҳ��ǩ -->
        <DIV class=subFolder id=subGuestReplyFunction><IMG id=subGuestReplyFunctionImg class=icon src="images/foldericon1.gif">���Իظ�ҳ��ǩ</DIV>
        <DIV class=sub id=subGuestReplyFunctionSub>
            <DIV class=subItem onclick="InsertLabel('{$ReplyGuest}')"><IMG class=icon src="images/label.gif">�ظ�����������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowJS_Guest}')"><IMG class=icon src="images/label.gif">����Js��֤</DIV>
            <DIV class=subItem onclick="InsertLabel('{$WriteTitle}')"><IMG class=icon src="images/label.gif">��ʾ���Ա���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ReplyId}')"><IMG class=icon src="images/label.gif">�ظ�����id</DIV>
        </DIV>
        <!-- ��������ҳ��ǩ -->
        <DIV class=subFolder id=subGuestSearchFunction><IMG id=subGuestSearchFunctionImg class=icon src="images/foldericon1.gif">��������ҳ��ǩ</DIV>
        <DIV class=sub id=subGuestSearchFunctionSub>
            <DIV class=subItem onclick="InsertLabel('{$ResultTitle}')"><IMG class=icon src="images/label.gif">�����������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SearchResult}')"><IMG class=icon src="images/label.gif">�������</DIV>
        </DIV>
     </DIV>
    <%
    End if
    if (ModuleType=5 or ModuleType=0) And eShop_Edition > 0 then%>
    <!--  �̳�Ƶ����ǩ  -->
     <DIV class=topFolder id=Shop><IMG id=ShopImg class=icon src="images/foldericon1.gif">�̳Ǳ�ǩ</DIV>
     <DIV class=sub id=ShopSub>
        <!-- �̳�ͨ��Ƶ����ǩ -->
        <DIV class=subFolder id=subShopChannelFunction><IMG id=subShopChannelFunctionImg class=icon src="images/foldericon1.gif"> �̳�ͨ�ñ�ǩ</DIV>
        <DIV class=sub id=subShopChannelFunctionSub>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetProductList','�̳��б�����ǩ',5,'GetList',800,750)"><IMG class=icon src="images/label3.gif">��ʾ��Ʒ�������Ϣ</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetPicProduct','��ʾͼƬ�̳Ǳ�ǩ',5,'GetPic',700,600)"><IMG class=icon src="images/label3.gif">��ʾͼƬ��Ʒ</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetSlidePicProduct','��ʾ�õ�Ƭ�̳Ǳ�ǩ',5,'GetSlide',700,460)"><IMG class=icon src="images/label3.gif">��ʾ�õ�Ƭ��Ʒ</DIV>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_CustomListLabel.asp','CustomListLable','�̳��Զ����б��ǩ',5,'GetProductCustom',720,700)"><IMG class=icon src="images/label3.gif">��Ʒ�Զ����б��ǩ</DIV>
        </DIV>
        <!--  �̳�Ƶ����ҳ��ǩ -->
        <DIV class=subFolder id=subshopcontent><IMG id=subshopcontentImg class=icon src="images/foldericon1.gif"> �̳����ݱ�ǩ</DIV>
        <DIV class=sub id=subshopcontentSub>
            <DIV class=subItem onclick="InsertLabel('{$ProductID}')"><IMG class=icon src="images/label.gif">��ƷID</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ProductName}')"><IMG class=icon src="images/label.gif">��Ʒ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ProductNum}')"><IMG class=icon src="images/label.gif">��Ʒ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ProductModel}')"><IMG class=icon src="images/label.gif">��Ʒ�ͺ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ProductStandard}')"><IMG class=icon src="images/label.gif">��Ʒ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ProducerName}')"><IMG class=icon src="images/label.gif">�� �� ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PresentExp}')"><IMG class=icon src="images/label.gif">�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PresentPoint}')"><IMG class=icon src="images/label.gif">���͵�ȯ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PresentMoney}')"><IMG class=icon src="images/label.gif">�������ֽ�ȯ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PointName}')"><IMG class=icon src="images/label.gif">��ȯ������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PointUnit}')"><IMG class=icon src="images/label.gif">��ȯ�ĵ�λ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Stocks}')"><IMG class=icon src="images/label.gif">��ʾ�����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ServiceTerm}')"><IMG class=icon src="images/label.gif">�ṩ����ʱ��</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$TrademarkName}')"><IMG class=icon src="images/label.gif">Ʒ���̱�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ProductTypeName}')"><IMG class=icon src="images/label.gif">��Ʒ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Price_Your}')"><IMG class=icon src="images/label.gif">��ǰ�����ߵļ۸�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">�ϼ�ʱ��</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_PhotoThumb.htm','240','140')"><IMG class=icon src="images/label2.gif">��Ʒ����ͼ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">��Ʒ�����</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ProductProperty}')"><IMG class=icon src="images/label.gif">��ʾ��ǰ��Ʒ�����ԣ����š��Ƽ����ȼ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ProductIntro}')"><IMG class=icon src="images/label.gif">��Ʒ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$OnTop}')"><IMG class=icon src="images/label.gif">��ʾ�̶�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Hot}')"><IMG class=icon src="images/label.gif">��ʾ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Elite}')"><IMG class=icon src="images/label.gif">��ʾ�Ƽ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Stars}')"><IMG class=icon src="images/label.gif">�Ƽ��ȼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$LimitNum}')"><IMG class=icon src="images/label.gif">�޹�����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Discount}')"><IMG class=icon src="images/label.gif">�����ۿ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$BeginDate}��{$EndDate}')"><IMG class=icon src="images/label.gif">��ʾ�Ż�����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Price_Original}')"><IMG class=icon src="images/label.gif">ԭʼ���ۼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Price_Market}')"><IMG class=icon src="images/label.gif">��ʾ�г���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Price}')"><IMG class=icon src="images/label.gif">��ʾ�̳Ǽ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Price_Member}')"><IMG class=icon src="images/label.gif">��ʾ��Ա��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Unit}')"><IMG class=icon src="images/label.gif">��ʾ��Ʒ��λ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SalePromotion}')"><IMG class=icon src="images/label.gif">��ʾ��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ProductExplain}')"><IMG class=icon src="images/label.gif">��ʾ��Ʒ˵��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$CorrelativeProduct}')"><IMG class=icon src="images/label.gif">��ʾ�����Ʒ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Vote}')"><IMG class=icon src="images/label.gif">��ʾ����</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_CorrelativeProduct.htm','240','260')"><IMG class=icon src="images/label2.gif">��ʾ��ϸ�����Ʒ</DIV>
        </DIV>
        <!-- �̳�Ƶ����ҳ��ǩ����  -->
        <!-- ���ﳵ��ǩ -->
        <DIV class=subFolder id=subshopping><IMG id=subshoppingImg class=icon src="images/foldericon1.gif"> �ҵĹ��ﳵ</DIV>
        <DIV class=sub id=subshoppingSub>
            <DIV class=subItem onclick="InsertLabel('{$ShowTips_Login}')"><IMG class=icon src="images/label.gif">�û���¼��ʾ</DIV>            
            <DIV class=subItem onclick="InsertLabel('{$UserName}')"><IMG class=icon src="images/label.gif">�û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GroupName}')"><IMG class=icon src="images/label.gif">�û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Discount_Member}%')"><IMG class=icon src="images/label.gif">��Ա�ۿ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$IsOffer}')"><IMG class=icon src="images/label.gif">�����������Ż�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Balance}')"><IMG class=icon src="images/label.gif">�ʽ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UserPoint}')"><IMG class=icon src="images/label.gif">���õ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UserExp}')"><IMG class=icon src="images/label.gif">���û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowTips_CartIsEmpty}')"><IMG class=icon src="images/label.gif">���ﳵ��ʾ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowCart}')"><IMG class=icon src="images/label.gif">��ʾ���ﳵ�е���Ʒ</DIV>
        </DIV>
        <!-- ���ﳵ��ǩ���� -->
        <!-- ����̨��ǩ -->
        <DIV class=subFolder id=subshopcash><IMG id=subshopcashImg class=icon src="images/foldericon1.gif"> �ա�����̨</DIV>
        <DIV class=sub id=subshopcashSub>
            <DIV class=subItem onclick="InsertLabel('{$ShowTips_Login}')"><IMG class=icon src="images/label.gif">�û���¼��ʾ</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ShowTips_CartIsEmpty}')"><IMG class=icon src="images/label.gif">���ﳵ��ʾ</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$UserName}')"><IMG class=icon src="images/label.gif">�û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GroupName}')"><IMG class=icon src="images/label.gif">�û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Discount_Member}%')"><IMG class=icon src="images/label.gif">��Ա�ۿ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$IsOffer}')"><IMG class=icon src="images/label.gif">�����������Ż�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Balance}')"><IMG class=icon src="images/label.gif">�ʽ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UserPoint}')"><IMG class=icon src="images/label.gif">���õ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UserExp}')"><IMG class=icon src="images/label.gif">���û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ContacterName}')"><IMG class=icon src="images/label.gif">�ջ�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">�ջ��˵�ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ZipCode}')"><IMG class=icon src="images/label.gif">�ջ����ʱ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Phone}')"><IMG class=icon src="images/label.gif">�ջ��˵绰</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Email}')"><IMG class=icon src="images/label.gif">�ջ�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PaymentTypeList}')"><IMG class=icon src="images/label.gif">���ʽ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$DeliverTypeList}')"><IMG class=icon src="images/label.gif">�ͻ���ʽ</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$InvoiceInfo}')"><IMG class=icon src="images/label.gif">��Ʊ��Ϣ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Remark}')"><IMG class=icon src="images/label.gif">��ע/����</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ShowCart}')"><IMG class=icon src="images/label.gif">��ʾ���ﳵ�е���Ʒ</DIV>
        </DIV>
        <!-- ����̨��ǩ���� -->
        <!-- ����Ԥ����ǩ -->
        <DIV class=subFolder id=subshopPreview><IMG id=subshopPreviewImg class=icon src="images/foldericon1.gif"> ����Ԥ����ǩ</DIV>
        <DIV class=sub id=subshopPreviewSub>
            <DIV class=subItem onclick="InsertLabel('{$ShowTips_CartIsEmpty}')"><IMG class=icon src="images/label.gif">���ﳵ��ʾ</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$UserName}')"><IMG class=icon src="images/label.gif">�û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$GroupName}')"><IMG class=icon src="images/label.gif">�û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Discount_Member}%')"><IMG class=icon src="images/label.gif">��Ա�ۿ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$IsOffer}')"><IMG class=icon src="images/label.gif">�����������Ż�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Balance}')"><IMG class=icon src="images/label.gif">�ʽ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UserPoint}')"><IMG class=icon src="images/label.gif">���õ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UserExp}')"><IMG class=icon src="images/label.gif">���û���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$TrueName}')"><IMG class=icon src="images/label.gif">�ջ�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">�ջ��˵�ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ZipCode}')"><IMG class=icon src="images/label.gif">�ջ����ʱ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Phone}')"><IMG class=icon src="images/label.gif">�ջ��˵绰</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowPaymentTypeList}')"><IMG class=icon src="images/label.gif">���ʽ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Company}')"><IMG class=icon src="images/label.gif">��Ʊ��Ϣ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowDeliverTypeList}')"><IMG class=icon src="images/label.gif">�ͻ���ʽ</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$ShowCart}')"><IMG class=icon src="images/label.gif">��ʾ���ﳵ�е���Ʒ</DIV>
        </DIV>
        <!-- ����Ԥ����ǩ���� -->
        <!-- ��ʾ�����ɹ���ǩ -->
        <DIV class=subFolder id=subshopSucceed><IMG id=subshopSucceedImg class=icon src="images/foldericon1.gif"> ��ʾ�����ɹ���ǩ</DIV>
        <DIV class=sub id=subshopSucceedSub>
            <DIV class=subItem onclick="InsertLabel('{$OrderFormNum}')"><IMG class=icon src="images/label.gif">��������</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$TotalMoney}')"><IMG class=icon src="images/label.gif">���׽��</DIV>
        </DIV>
        <!-- ��ʾ�����ɹ���ǩ���� -->
        <!-- ��ʾ�����ɹ���ǩ -->
        <DIV class=subFolder id=subshopPayment><IMG id=subshopPaymentImg class=icon src="images/foldericon1.gif"> ����֧����ǩ</DIV>
        <DIV class=sub id=subshopPaymentSub>
            <DIV class=subItem onclick="InsertLabel('{$OrderFormNum}')"><IMG class=icon src="images/label.gif">��������</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$MoneyTotal}')"><IMG class=icon src="images/label.gif">�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$MoneyReceipt}')"><IMG class=icon src="images/label.gif">�� �� ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$MoneyNeedPay}')"><IMG class=icon src="images/label.gif">��Ҫ֧��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PaymentNum}')"><IMG class=icon src="images/label.gif">֧�����к�</DIV>        
            <DIV class=subItem onclick="InsertLabel('��{$vMoney}')"><IMG class=icon src="images/label.gif">֧�����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PayOnlineRate}')"><IMG class=icon src="images/label.gif">������</DIV>
            <DIV class=subItem onclick="InsertLabel('��{$v_amount}')"><IMG class=icon src="images/label.gif">ʵ�ʻ�����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PayOnlineProviderName}')"><IMG class=icon src="images/label.gif">����֧��ƽ̨�ṩ��</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$HiddenField}')"><IMG class=icon src="images/label.gif">֧�������ֶ�</DIV>
        </DIV>
        <!-- ��ʾ�����ɹ���ǩ���� -->
     </DIV>
    <% 
    End if
    If (ModuleType=7 or ModuleType=0) And FoundInArr(AllModules, "House", ",") Then %>
    <!--  ����Ƶ������  -->
     <DIV class=topFolder id=House><IMG id=HouseImg class=icon src="images/foldericon1.gif">������ǩ</DIV>
    <DIV class=sub id=HouseSub>
         <!-- ����ͨ��Ƶ����ǩ -->
         <DIV class=subFolder id=subHouseChannelFunction><IMG id=subHouseChannelFunctionImg class=icon src="images/foldericon1.gif"> ����Ƶ����ǩ</DIV>
         <DIV class=sub id=subHouseChannelFunctionSub>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_HouseList.htm',560,360)"><IMG class=icon src="images/label2.gif">��ʾ������ַ����Ϣ</DIV>
        </DIV>
        <!-- ����ͨ��Ƶ����ǩ���� -->
        <!-- ����Ƶ�����ݱ�ǩ -->
        <DIV class=subFolder id=subHouseChannelContent><IMG id=subHouseChannelContentImg class=icon src="images/foldericon1.gif"> �������ݱ�ǩ</DIV>
        <DIV class=sub id=subHouseChannelContentSub>
            <DIV class=subItem onclick="InsertLabel('{$HeZhuType}')"><IMG class=icon src="images/label.gif">��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseDiZhi}')"><IMG class=icon src="images/label.gif">���ݵ�ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$My}')"><IMG class=icon src="images/label.gif">�ҵļ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Chum}')"><IMG class=icon src="images/label.gif">����Ҫ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseHuXing}')"><IMG class=icon src="images/label.gif">���ݻ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseHuXing1}')"><IMG class=icon src="images/label.gif">���ⲿ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseHuXing2}')"><IMG class=icon src="images/label.gif">���ò���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseXingZhi}')"><IMG class=icon src="images/label.gif">�������ʣ��·������ַ��ȣ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseChanQuan}')"><IMG class=icon src="images/label.gif">���ݲ�Ȩ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseJianCheng}')"><IMG class=icon src="images/label.gif">��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseJianCheng1}')"><IMG class=icon src="images/label.gif">�����������ڷ�Χ��ʼ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseJianCheng2}')"><IMG class=icon src="images/label.gif">�����������ڷ�Χ���գ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseMianJi}')"><IMG class=icon src="images/label.gif">�������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseMianJi1}')"><IMG class=icon src="images/label.gif">���������Χ��ʼ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseMianJi2}')"><IMG class=icon src="images/label.gif">���������Χ���գ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseLouCeng}')"><IMG class=icon src="images/label.gif">¥��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseLeiXing}')"><IMG class=icon src="images/label.gif">��ҵ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseChaoXiang}')"><IMG class=icon src="images/label.gif">���ݳ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseShuiDian}')"><IMG class=icon src="images/label.gif">ˮ����ʩ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseSheShi}')"><IMG class=icon src="images/label.gif">������ʩ�����ݡ����⡢���ȡ�ˮ���ȣ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseZhuangXiu}')"><IMG class=icon src="images/label.gif">װ�޳̶�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseDianQi}')"><IMG class=icon src="images/label.gif">�����豸</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseWeiSheng}')"><IMG class=icon src="images/label.gif">������ʩ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseJiaJu}')"><IMG class=icon src="images/label.gif">�����Ҿ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseXinXi}')"><IMG class=icon src="images/label.gif">��Ϣ��ʩ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseGongJia}')"><IMG class=icon src="images/label.gif">��������</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$HouseHuanJing}')"><IMG class=icon src="images/label.gif">��������</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$JiaoFangStartDate}')"><IMG class=icon src="images/label.gif">��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseQiTa}')"><IMG class=icon src="images/label.gif">����˵�����磺����ͼƬ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$TotalPrice}')"><IMG class=icon src="images/label.gif">���ݼ۸����ڳ��ۣ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseZuJin}')"><IMG class=icon src="images/label.gif">����������ڳ��⣩</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HousePrice1}')"><IMG class=icon src="images/label.gif">�����۸�Χ����ͣ������󹺣�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HousePrice2}')"><IMG class=icon src="images/label.gif">�����۸�Χ����ߣ������󹺣�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseZuJin1}')"><IMG class=icon src="images/label.gif">�������Χ����ͣ��������⣩</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseZuJin2}')"><IMG class=icon src="images/label.gif">�������Χ����ߣ��������⣩</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseZhiFu}')"><IMG class=icon src="images/label.gif">֧����ʽ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HousePriceType}')"><IMG class=icon src="images/label.gif">�۸�λ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$HouseZuJinType}')"><IMG class=icon src="images/label.gif">���λ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ZuLinStartDate}')"><IMG class=icon src="images/label.gif">����ʱ�䷶Χ��ʼ��</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ZuLinEndDate}')"><IMG class=icon src="images/label.gif">����ʱ�䷶Χ��ĩ��</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$JiaoFangStartDate}')"><IMG class=icon src="images/label.gif">��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ContactPhone}')"><IMG class=icon src="images/label.gif">��ϵ�绰</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ContactName}')"><IMG class=icon src="images/label.gif">��ϵ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ContactEmail}')"><IMG class=icon src="images/label.gif">��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ContactQQ}')"><IMG class=icon src="images/label.gif">��ϵ�ѣ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Editor}')"><IMG class=icon src="images/label.gif">������Ϣ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Hits}')"><IMG class=icon src="images/label.gif">�����</DIV>
        </DIV>
        <!-- ����Ƶ�����ݱ�ǩ���� -->
    </DIV>
    <%
    End if
    If (ModuleType=8 or ModuleType=0) And FoundInArr(AllModules, "Job", ",") Then %>
    <!--  �˲���ƸƵ������  -->
     <DIV class=topFolder id=Job><IMG id=JobImg class=icon src="images/foldericon1.gif">�˲���Ƹ��ǩ</DIV>
    <DIV class=sub id=JobSub>
        <!-- �˲���Ƹͨ��Ƶ����ǩ -->
        <DIV class=subFolder id=subJobChannelFunction><IMG id=subJobChannelFunctionImg class=icon src="images/foldericon1.gif"> �˲���ƸƵ����ҳ���б���ǩ</DIV>
        <DIV class=sub id=subJobChannelFunctionSub>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetPositionList','ְλ�б�����ǩ',8,'GetPositionList',650,500)"><IMG class=icon src="images/label3.gif">��ʾ�����С����¡�������ְλ���Ƶ���Ϣ</DIV>
        </DIV>
        <DIV class=subFolder id=subJobChannelFunction2><IMG id=subJobChannelFunction2Img class=icon src="images/foldericon1.gif"> �˲���ƸƵ����ҳ�����ݣ���ǩ</DIV>
        <DIV class=sub id=subJobChannelFunction2Sub>
            <DIV class=subItem onclick="FunctionLabel2('��PositionList_Content��')"><IMG class=icon src="images/label2.gif">ѭ����ʾְλ������Ϣ</DIV>
         </DIV>
        <DIV class=subFolder id=subJobChannelFunction3><IMG id=subJobChannelFunction3Img class=icon src="images/foldericon1.gif"> �˲���ƸƵ���������ҳ��ǩ</DIV>
        <DIV class=sub id=subJobChannelFunction3Sub>
            <DIV class=subItem onclick="SuperFunctionLabel('editor_label.asp','GetSearchResult','ְλ��������б�����ǩ',8,'GetSearchResult',590,450)"><IMG class=icon src="images/label3.gif">��ʾ�����������ְλ���Ƶ���Ϣ</DIV>
        </DIV>
        <!-- �˲���Ƹͨ��Ƶ����ǩ���� -->
        <!-- �˲���ƸƵ�����ݱ�ǩ -->
        <DIV class=subFolder id=subJobChannelContent><IMG id=subJobChannelContentImg class=icon src="images/foldericon1.gif"> �˲���Ƹ���ݱ�ǩ</DIV>
        <DIV class=sub id=subJobChannelContentSub>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_CorrelativePosition.htm',560,360)"><IMG class=icon src="images/label2.gif">��ʾ���ְλ���Ƶ���Ϣ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PositionName}')"><IMG class=icon src="images/label.gif">ְλ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$WorkPlaceName}')"><IMG class=icon src="images/label.gif">�����ص�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PositionNum}')"><IMG class=icon src="images/label.gif">��Ƹ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ReleaseDate}')"><IMG class=icon src="images/label.gif">��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ValidDate}')"><IMG class=icon src="images/label.gif">��Ч��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SubCompanyName}')"><IMG class=icon src="images/label.gif">���˵�λ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Contacter}')"><IMG class=icon src="images/label.gif">��ϵ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Telephone}')"><IMG class=icon src="images/label.gif">��ϵ�绰</DIV>
            <DIV class=subItem onclick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">��ϵ��ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$E_mail}')"><IMG class=icon src="images/label.gif">��ϵE_mail</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PositionDescription')"><IMG class=icon src="images/label.gif">ְλ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$DutyRequest}')"><IMG class=icon src="images/label.gif">��ְҪ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$PositionStatus}')"><IMG class=icon src="images/label.gif">ְλ״̬</DIV>
            <DIV class=subItem onclick="InsertLabel('{$SaveSupply}')"><IMG class=icon src="images/label.gif">����ְλ��ť</DIV>
        </DIV>
        <!-- �˲���ƸƵ�����ݱ�ǩ���� -->
    </DIV>
    <%End if
    If (ModuleType=6 or ModuleType=0) And FoundInArr(AllModules, "Supply", ",") then%>
        <DIV class=subFolder id=subsupplyInfo><IMG id=subsupplyInfoImg class=icon src="images/foldericon1.gif">������Ϣҳ��ǩ</DIV>
        <DIV class=sub id=subsupplyInfoSub>
        <DIV class=subItem onclick="FunctionLabel('Lable/PE_SupplyInfo.htm','600','700')"><IMG class=icon src="images/label2.gif">������Ϣ�б��ǩ</DIV>
        <DIV class=subItem onclick="FunctionLabel('Lable/PE_SupplyLasterInfo.htm','600','350')"><IMG class=icon src="images/label2.gif">����������Ϣ�б��ǩ</DIV>
        <DIV class=subFolder id=subsupplyInfoContent><IMG id=subsupplyInfoContentImg class=icon src="images/foldericon1.gif"> ������Ϣ���ݱ�ǩ</DIV>
        <DIV class=sub id=subsupplyInfoContentSub>
            <DIV class=subItem onclick="InsertLabel('{$SupplyInfoTitle}')"><IMG class=icon src="images/label.gif">��Ϣ����</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$SupplyInfoType}')"><IMG class=icon src="images/label.gif">��Ϣ����</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$TradeType}')"><IMG class=icon src="images/label.gif">���׷�ʽ</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$UserName}')"><IMG class=icon src="images/label.gif">�� �� ��</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$UpdateTime}')"><IMG class=icon src="images/label.gif">��������</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$EndTime}')"><IMG class=icon src="images/label.gif">��Ч����</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$SupplyIntro}')"><IMG class=icon src="images/label.gif">��ϸ����</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$Province}')"><IMG class=icon src="images/label.gif">������������ʡ</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$City}')"><IMG class=icon src="images/label.gif">��������������</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">��ϵ��ַ</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$ZipCode}')"><IMG class=icon src="images/label.gif">�ʡ�����</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$Email}')"><IMG class=icon src="images/label.gif">�����ʼ�</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$CompanyName}')"><IMG class=icon src="images/label.gif">�� ˾ ��</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$Department}')"><IMG class=icon src="images/label.gif">��������</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$CompanyAddress}')"><IMG class=icon src="images/label.gif">��˾��ַ</DIV> 
            <DIV class=subItem onclick="InsertLabel('{$SupplyAction}')"><IMG class=icon src="images/label.gif">��ʾ���������ۡ������ߺ��ѡ�����ӡ���ġ����رմ��ڡ�</DIV>
        </DIV>
        <DIV class=subItem onclick="FunctionLabel('Lable/PE_SupplySearchInfo.htm','500','250')"><IMG class=icon src="images/label2.gif">������Ϣ����������ǩ</DIV>    
        <DIV class=subItem onclick="InsertLabel('{$SearchResul}')"><IMG class=icon src="images/label.gif">��ʾ�������ҳ��ǩ</DIV> 
        </DIV>
    </DIV>
    <%End if%>
     <!--  ����,��Դ,����,Ʒ��,��ǩ  -->
     <DIV class=topFolder id=Aomb><IMG id=AombImg class=icon src="images/foldericon1.gif">����,��Դ,����,Ʒ��</DIV>
     <DIV class=sub id=AombSub>
         <!-- ���� ��ǩ -->
         <DIV class=subFolder id=Author><IMG id=AuthorImg class=icon src="images/foldericon1.gif">���߱�ǩ</DIV>
         <DIV class=sub id=AuthorSub>
            <DIV class=subItem onclick="InsertLabel('{$AuthorName}')"><IMG class=icon src="images/label.gif">��������</DIV>    
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Author_Photo.htm','240','150')"><IMG class=icon src="images/label2.gif">������Ƭ</DIV>    
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Author_List.htm','240','230')"><IMG class=icon src="images/label2.gif">��ʾ�����б�</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$AuthorSex}')"><IMG class=icon src="images/label.gif">�����Ա�</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$AuthorAddTime}')"><IMG class=icon src="images/label.gif">�ļ���׼ʱ��</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$AuthorBirthDay}')"><IMG class=icon src="images/label.gif">��������</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$AuthorCompany}')"><IMG class=icon src="images/label.gif">���߹�˾</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorDepartment}')"><IMG class=icon src="images/label.gif">���߲���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorAddress}')"><IMG class=icon src="images/label.gif">���ߵ�ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorTel}')"><IMG class=icon src="images/label.gif">���ߵ绰</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorFax}')"><IMG class=icon src="images/label.gif">���ߴ���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorZipCode}')"><IMG class=icon src="images/label.gif">�����ʱ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorHomePage}')"><IMG class=icon src="images/label.gif">������ҳ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorEmail}')"><IMG class=icon src="images/label.gif">�����ʼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorQQ}')"><IMG class=icon src="images/label.gif">����QQ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorType}')"><IMG class=icon src="images/label.gif">���߷���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$AuthorIntro}')"><IMG class=icon src="images/label.gif">����˵��</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Author_ArtList.htm','350','330')"><IMG class=icon src="images/label2.gif">���������б�</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Author_ShowList.htm','400','345')"><IMG class=icon src="images/label2.gif">��ʾ�����б�</DIV>
         </DIV>
         <!-- ��Դ ��ǩ -->
         <DIV class=subFolder id=origin><IMG id=originImg class=icon src="images/foldericon1.gif">��Դ��ǩ</DIV>
         <DIV class=sub id=originSub>
            <DIV class=subItem onclick="InsertLabel('{$ShowPhoto}')"><IMG class=icon src="images/label.gif">��ԴͼƬ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowName}')"><IMG class=icon src="images/label.gif">��Դ����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowContacterName}')"><IMG class=icon src="images/label.gif">��ϵ��</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowAddress}')"><IMG class=icon src="images/label.gif">��ַ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowTel}')"><IMG class=icon src="images/label.gif">�绰</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowFax}')"><IMG class=icon src="images/label.gif">����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowZipCode}')"><IMG class=icon src="images/label.gif">�ʱ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowMail}')"><IMG class=icon src="images/label.gif">����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowHomePage}')"><IMG class=icon src="images/label.gif">��ҳ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowEmail}')"><IMG class=icon src="images/label.gif">�ʼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowQQ}')"><IMG class=icon src="images/label.gif">QQ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowType}')"><IMG class=icon src="images/label.gif">����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowMemo}')"><IMG class=icon src="images/label.gif">���</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowArticleList}')"><IMG class=icon src="images/label.gif">��ʾ�����б�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowCopyFromList}')"><IMG class=icon src="images/label.gif">��Դ�б�</DIV>  
         </DIV>
         <!-- ���̱�ǩ -->
         <DIV class=subFolder id=manufacturer><IMG id=manufacturerImg class=icon src="images/foldericon1.gif">���̱�ǩ</DIV>
         <DIV class=sub id=manufacturerSub>
            <DIV class=subItem onclick="InsertLabel('{$ShowPhoto}')"><IMG class=icon src="images/label.gif">����ͼƬ</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$ShowName}')"><IMG class=icon src="images/label.gif">����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowProducerShortName}')"><IMG class=icon src="images/label.gif">��д</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowBirthDay}')"><IMG class=icon src="images/label.gif">��������</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowAddress}')"><IMG class=icon src="images/label.gif">��ַ</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ShowTel}')"><IMG class=icon src="images/label.gif">�绰</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ShowFax}')"><IMG class=icon src="images/label.gif">����</DIV>    
            <DIV class=subItem onclick="InsertLabel('{$ShowZipCode}')"><IMG class=icon src="images/label.gif">�ʱ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowHomePage}')"><IMG class=icon src="images/label.gif">��ҳ</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowEmail}')"><IMG class=icon src="images/label.gif">�ʼ�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowType}')"><IMG class=icon src="images/label.gif">����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowtrademarkList}')"><IMG class=icon src="images/label.gif">����Ʒ��</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$ShowMemo}')"><IMG class=icon src="images/label.gif">���</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Product_List.htm','400','230')"><IMG class=icon src="images/label2.gif">��ʾ��Ʒ�б�</DIV>
         </DIV>
         <!-- Ʒ�Ʊ�ǩ -->
         <DIV class=subFolder id=brand><IMG id=brandImg class=icon src="images/foldericon1.gif">Ʒ�Ʊ�ǩ</DIV>
         <DIV class=sub id=brandSub>
            <DIV class=subItem onclick="InsertLabel('{$ShowPhoto}')"><IMG class=icon src="images/label.gif">Ʒ��ͼƬ</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$ShowName}')"><IMG class=icon src="images/label.gif">����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowType}')"><IMG class=icon src="images/label.gif">����</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowProducerName}')"><IMG class=icon src="images/label.gif">��������</DIV>        
            <DIV class=subItem onclick="InsertLabel('{$ShowMemo}')"><IMG class=icon src="images/label.gif">���</DIV>
            <DIV class=subItem onclick="FunctionLabel('Lable/PE_Product_List.htm','400','230')"><IMG class=icon src="images/label2.gif">��ʾ��Ʒ�б�</DIV>
            <DIV class=subItem onclick="InsertLabel('{$ShowtrademarkList}')"><IMG class=icon src="images/label.gif">��ʾƷ���б�</DIV>  
         </DIV>
     </DIV>
     <!-- Rss��ǩ -->
     <DIV class=topFolder id=RssItem><IMG id=RssItemImg class=icon src="images/foldericon1.gif">RSS</DIV>
     <DIV class=sub id=RssItemSub>
        <DIV class=subItem onclick="InsertLabel('{$Rss}')"><IMG class=icon src="images/label.gif">RSS��ǩ��ʾ</DIV>    
        <DIV class=subItem onclick="InsertLabel('{$RssElite}')"><IMG class=icon src="images/label.gif">RSS�Ƽ���ǩ��ʾ</DIV>    
        <DIV class=subItem onclick="InsertLabel('{$RssHot}')"><IMG class=icon src="images/label.gif">RSS�ȵ����±�ǩ��ʾ</DIV>
     </DIV>
     <DIV class=topFolder id=AnnounceItem><IMG id=AnnounceItemImg class=icon src="images/foldericon1.gif">�����ǩ</DIV>
     <DIV class=sub id=AnnounceItemSub>
        <DIV class=subItem onclick="InsertLabel('{$AnnounceList}')"><IMG class=icon src="images/label.gif">�����б�</DIV>     
     </DIV>
     <DIV class=topFolder id=FriendItem><IMG id=FriendItemImg class=icon src="images/foldericon1.gif">�������ӱ�ǩ</DIV>
     <DIV class=sub id=FriendItemSub>
        <DIV class=subItem onclick="InsertLabel('{$FriendSiteList}')"><IMG class=icon src="images/label.gif">���������б�</DIV>
     </DIV>
     <DIV class=topFolder id=VoteItem><IMG id=VoteItemImg class=icon src="images/foldericon1.gif">�����ǩ</DIV>
     <DIV class=sub id=VoteItemSub>
         <DIV class=subItem onclick="InsertLabel('[VoteItem] ������������Ҫѭ������ı�ǩ[/VoteItem] ')"><IMG class=icon src="images/label.gif">ѭ����ʾ������Ŀ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$VoteTitle}')"><IMG class=icon src="images/label.gif">��ʾ�������</DIV>
        <DIV class=subItem onclick="InsertLabel('{$TotalVote}')"><IMG class=icon src="images/label.gif">����ͶƱ����</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ItemNum}')"><IMG class=icon src="images/label.gif">����ѡ������</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ItemSelect}')"><IMG class=icon src="images/label.gif">����ѡ������</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ItemPer}')"><IMG class=icon src="images/label.gif">����ѡ����ռ�ٷֱ�</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ItemAnswer}')"><IMG class=icon src="images/label.gif">����ѡ������Ʊ��</DIV>
        <DIV class=subItem onclick="InsertLabel('{$VoteForm}')"><IMG class=icon src="images/label.gif">����ѡ������</DIV>        
        <DIV class=subItem onclick="InsertLabel('{$OtherVote}')"><IMG class=icon src="images/label.gif">�鿴����������Ŀ</DIV>
     </DIV>
     <!-- Wap��ǩ -->
     <DIV class=topFolder id=WapItem><IMG id=WapItemImg class=icon src="images/foldericon1.gif">Wap��ǩ</DIV>
     <DIV class=sub id=WapItemSub>    
        <DIV class=subItem onclick="InsertLabel('{$Wap}')"><IMG class=icon src="images/label.gif">WAP��ǩ��ʾ</DIV>    
     </DIV>
     <!-- ��Ա��ǩ -->
     <DIV class=topFolder id=associatorItem><IMG id=associatorItemImg class=icon src="images/foldericon1.gif">��Ա�����ǩ</DIV>
     <DIV class=sub id=associatorItemSub>
        <DIV class=subItem onclick="InsertLabel('{$UserFace}')"><IMG class=icon src="images/label.gif">��Աͷ��</DIV>    
        <DIV class=subItem onclick="InsertLabel('{$TrueName}')"><IMG class=icon src="images/label.gif">����</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Sex}')"><IMG class=icon src="images/label.gif">�Ա�</DIV>
        <DIV class=subItem onclick="InsertLabel('{$BirthDay}')"><IMG class=icon src="images/label.gif">����</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Company}')"><IMG class=icon src="images/label.gif">��˾</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Department}')"><IMG class=icon src="images/label.gif">����</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Address}')"><IMG class=icon src="images/label.gif">��ַ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$HomePhone}')"><IMG class=icon src="images/label.gif">�绰</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Fax}')"><IMG class=icon src="images/label.gif">����</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ZipCode}')"><IMG class=icon src="images/label.gif">�ʱ�</DIV>
        <DIV class=subItem onclick="InsertLabel('{$HomePage}')"><IMG class=icon src="images/label.gif">��ҳ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$Email}')"><IMG class=icon src="images/label.gif">�ʼ�</DIV>
        <DIV class=subItem onclick="InsertLabel('{$QQ}')"><IMG class=icon src="images/label.gif">QQ</DIV>
        <DIV class=subItem onclick="InsertLabel('{$ShowUserList}')"><IMG class=icon src="images/label.gif">��Ա�б�</DIV>
     </DIV>
     <!--  �Զ����ǩ  -->
     <%
        Set rs=Server.CreateObject("ADODB.Recordset")
        sql="select LabelID,LabelName,LabelClass,LabelType,fieldlist from PE_Label Where LabelType=0 Order by LabelClass,LabelID desc"
        rs.open sql,conn,1,1
        If not(rs.bof and rs.EOF) Then
            response.Write("<DIV class=topFolder id=Label><IMG id=LabelImg class=icon src=""images/foldericon1.gif"">�Զ��徲̬��ǩ</DIV>")
            response.Write("<DIV class=sub id=LabelSub>")
            Do while not rs.eof
                Response.write "<DIV class=subItem onclick=""InsertLabel('{$" & rs("LabelName")& "}')""><IMG class=icon src=""images/label.gif"">"
                If Trim(rs("LabelClass") & "") <> "" Then Response.write "<font color=#999999>[" & rs("LabelClass") & "]</font>"
                Response.write rs("LabelName") &"</DIV>"
                rs.movenext
            loop
            response.Write("</DIV>")
        End If
        rs.close
        sql="select LabelID,LabelName,LabelClass,LabelType,fieldlist from PE_Label Where LabelType=1 Order by LabelClass,LabelID desc"
        rs.open sql,conn,1,1
        If not(rs.bof and rs.EOF) Then
            response.Write("<DIV class=topFolder id=Label1><IMG id=Label1Img class=icon src=""images/foldericon1.gif"">�Զ��嶯̬��ǩ</DIV>")
            response.Write("<DIV class=sub id=Label1Sub>")
            Do while not rs.eof
                Response.write "<DIV class=subItem onclick=""InsertLabel('{$" & rs("LabelName")& "}')""><IMG class=icon src=""images/label.gif"">"
                If Trim(rs("LabelClass") & "") <> "" Then Response.write "<font color=#999999>[" & rs("LabelClass") & "]</font>"
                Response.write rs("LabelName") &"</DIV>"
                rs.movenext
            loop
            response.Write("</DIV>")
        End If
        rs.close
        sql="select LabelID,LabelName,LabelClass,LabelType,fieldlist from PE_Label Where LabelType=2 Order by LabelClass,LabelID desc"
        rs.open sql,conn,1,1
        If not(rs.bof and rs.EOF) Then
            response.Write("<DIV class=topFolder id=Label2><IMG id=Label2Img class=icon src=""images/foldericon1.gif"">�Զ���ɼ���ǩ</DIV>")
            response.Write("<DIV class=sub id=Label2Sub>")
            Do while not rs.eof
                Response.write "<DIV class=subItem onclick=""InsertLabel('{$" & rs("LabelName")& "}')""><IMG class=icon src=""images/label.gif"">"
                If Trim(rs("LabelClass") & "") <> "" Then Response.write "<font color=#999999>[" & rs("LabelClass") & "]</font>"
                Response.write rs("LabelName") &"</DIV>"
                rs.movenext
            loop
            response.Write("</DIV>")
        End If
        rs.close
        sql="select LabelID,LabelName,LabelClass,LabelType,fieldlist from PE_Label Where LabelType=3 Order by LabelClass,LabelID desc"
        rs.open sql,conn,1,1
        If not(rs.bof and rs.EOF) Then
            response.Write("<DIV class=topFolder id=Label3><IMG id=Label3Img class=icon src=""images/foldericon1.gif"">�Զ��庯����ǩ</DIV>")
            response.Write("<DIV class=sub id=Label3Sub>")
            Do while not rs.eof
                Response.write "<DIV class=subItem onclick=""FunctionLabel('editor_listdynafield.asp?id=" & rs("LabelID")& "','400','480')""><IMG class=icon src=""images/label.gif"">"
                If Trim(rs("LabelClass") & "") <> "" Then Response.write "<font color=#999999>[" & rs("LabelClass") & "]</font>"
                Response.write rs("LabelName") &"</DIV>"
                rs.movenext
            loop
            response.Write("</DIV>")
        End If
        rs.close
        set rs=nothing
      %>
     <!--  �Զ����ֶα�ǩ  -->
     <DIV class=topFolder id=Field><IMG id=FieldImg class=icon src="images/foldericon1.gif">�Զ����ֶα�ǩ</DIV>
     <DIV class=sub id=FieldSub>
     <%
        sql="select  LabelName,FieldName from PE_Field where ChannelID=" & PE_Clng(ChannelID) & " Order by FieldID desc"
        Set rs=Server.CreateObject("ADODB.Recordset")
            rs.open sql,conn,1,1
            if rs.bof and  rs.eof then
                response.Write("<li>����û���Զ����ֶα�ǩ,���Զ����ֶα�ǩֻ��ʾ������Ƶ��</li>")
            else
                Do while not rs.eof
                    Response.write"<DIV class=subItem onclick=""InsertLabel('" & rs("LabelName")& "')""><IMG class=icon src=""images/label.gif"">"& rs("FieldName") &"</DIV>"
                    rs.movenext
                loop
            end if 
            rs.close
        set rs=nothing
      %>
    </DIV>
    <!--  ���λ��ǩ��ʼ  -->
    <DIV class=topFolder id=AdJs><IMG id=AdJsImg class=icon src="images/foldericon1.gif">����λ��ǩ</DIV>
    <DIV class=sub id=AdJsSub>
    <%
        Dim XmlDoc
        Function XmlText(ByVal iBigNode, ByVal iSmallNode, ByVal DefChar)
            Dim LangRoot
            If IsNull(iBigNode) Or IsNull(iSmallNode) Then
                XmlText = DefChar
            Else
                Set LangRoot = XmlDoc.getElementsByTagName(iBigNode)
                If LangRoot.Length = 0 Then
                    XmlText = DefChar
                Else
                    On Error Resume Next
                    XmlText = LangRoot(0).selectSingleNode(iSmallNode).Text
                    If XmlText = "" Then XmlText = DefChar
                End If
                Set LangRoot = Nothing
            End If
        End Function
        Private Function GetZoneJSName(iZoneID, UpdateTime)
            Set XmlDoc = CreateObject("Microsoft.XMLDOM")
            XmlDoc.async = False
            XmlDoc.Load (Server.MapPath(InstallDir & "Language/Gb2312.xml"))
            GetZoneJSName = InstallDir & PE_CStr(Conn.Execute("select ADDir from PE_Config")(0), "AD") & "/" & Year(UpdateTime) & Right("0" & Month(UpdateTime), 2) & "/" & iZoneID & ".js"
            Set XmlDoc = Nothing
        End Function
        sql="select ZoneID,ZoneName,UpdateTime from PE_AdZone"
        Set rs=Server.CreateObject("ADODB.Recordset")
        rs.open sql,conn,1,1
        if rs.bof and  rs.eof then
           response.Write("<li>����û�ж������λ </li>")
        else
            Do while not rs.eof             
                Response.write"<DIV class=subItem onclick=""InsertAdjs('Adjs','" & GetZoneJSName(rs("ZoneID"), rs("UpdateTime")) & "')""><IMG class=icon src=""images/jscript.gif"">"& rs("ZoneName") &"</DIV>"
                rs.movenext
            loop
        end if 
        rs.close
        set rs=nothing
        conn.close
        set conn=nothing
    %>
    </DIV>
    <!--  ���λ��ǩ����  -->
     <!-- ����JS��ǩ  -->
     <DIV class=topFolder id=OtherJS><IMG id=OtherJSImg class=icon src="images/foldericon1.gif">����JS��ǩ</DIV>
     <DIV class=sub id=OtherJSSub>
     <DIV class=subItem onclick="InsertAdjs('SwitchFont','{$InstallDir}js/gb_big5.js')"><IMG class=icon src="images/jscript.gif">�л������w����</DIV>
     <DIV class=subItem onclick="FunctionLabel2('ResumeError')"><IMG class=icon src="images/jscript.gif">����ҳ��JS����</DIV>
</DIV>
  </td>
 </tr>
   </td>
  </tr>
</table>
<!-- ******** �˵�Ч������ ******** -->
    <!-- ��ʾ˵�� -->
    <table width='100%' height='60' border='0' align='center' cellpadding='0' cellspacing='0' bgcolor="#EEF4FF" style='border: 1px solid #0066FF;'>
      <tr align="center">
        <td height="22" colspan="2" bgcolor='#0066FF'><font color="#FFFFFF">==&gt;&nbsp;��ʾ˵��&nbsp;&lt;==</font></td>
      </tr>
      <tr>
        <td width="9%" rowspan="3">&nbsp;</td>
        <td width="91%"><IMG class=icon src="images/label.gif"> >>>  ��ͨ��ǩ </td>
      </tr>
      <tr>
        <td><IMG class=icon src="images/label2.gif"> >>> ������ǩ </td>
      </tr>
      <tr>
        <td><IMG class=icon src="images/label3.gif"> >>> ����������ǩ </td>
      </tr>
    </table>
    <!-- ��ʾ���� -->
</body>
</html>
