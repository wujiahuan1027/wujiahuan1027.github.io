<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>选取课时</title>
<style>
body,form,legend,td,fieldset {font-size:9pt;cursor:default}
input {border-width:1px;font-size:9pt;}
label {cursor:hand;}
.content{padding-left:20px;}
</style>
</head>

<body style="background-color:buttonface;padding:0;margin:3;border:0;overflow:auto">
<%
Response.Write "<form method='post' name='myform' action=''>" & vbCrLf
Response.Write "<fieldset style='width:95%;height:100%'>" & vbCrLf
Response.Write "<legend>固定排课</legend>" & vbCrLf
Response.Write "<table align='center' width='95%'  border='0' cellspacing='0' cellpadding='0'>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td colspan='4'> 课时列表： </td>" & vbCrLf
Response.Write "    <td>&nbsp;</td>" & vbCrLf
Response.Write "    <td> <strong>已经选定的课时： </strong> </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align='center'><a href='#' onclick=""add('1')"">1节</a></td>" & vbCrLf
Response.Write "    <td align='center'><a href='#' onclick=""add('2')"">2节</a></td>" & vbCrLf
Response.Write "    <td align='center'><a href='#' onclick=""add('3')"">3节</a></td>" & vbCrLf
Response.Write "    <td align='center'><a href='#' onclick=""add('4')"">4节</a></td>" & vbCrLf
Response.Write "    <td>&nbsp;&gt;&gt;</td>" & vbCrLf
Response.Write "    <td><input type='text' name='LessonList' size='20' maxlength='100' readonly='readonly'>" & vbCrLf
Response.Write "    </td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align='center'><a href='#' onclick=""add('5')"">5节</a></td>" & vbCrLf
Response.Write "    <td align='center'><a href='#' onclick=""add('6')"">6节</a></td>" & vbCrLf
Response.Write "    <td align='center'><a href='#' onclick=""add('7')"">7节</a></td>" & vbCrLf
Response.Write "    <td align='center'><a href='#' onclick=""add('8')"">8节</a></td>" & vbCrLf
Response.Write "    <td>&nbsp;</td>" & vbCrLf
Response.Write "    <td><input type='button' name='del1' onclick='del(1)' value='删除最后'> <input type='button' name='del2' onclick='del(0)' value='删除全部'></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "</table>" & vbCrLf
ShowStyle = PE_CLng(Trim(Request("Showstyle")))
If ShowStyle = 1 Then '星期
    DaySelectVB = CStr(Trim(Request("DaySelect")))
    DayTdVB = Right(DaySelectVB, Len(DaySelectVB) - 3)
    Dim A(20)
    For i = 0 To 22 - 1 - 1
        A(i) = DateAdd("d", i * 7, CDate(DayTdVB))
    Next
    astr = Join(A, """,""")

    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "var B = new Array("""&astr&""");" & vbCrLf
    Response.Write "var parentTd = new Array();" & vbCrLf
    Response.Write "var DaySelect = new Array();" & vbCrLf
    Response.Write "for (i=0; i <=20; i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "    parentTd[i] = opener.document.getElementById('sstr'+B[i]);" & vbCrLf
    Response.Write "    DaySelect[i] ='str'+ B[i];" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "var strLinka='<a href=""#"" onclick=\'window.open(""Admin_SelectTime.asp?ShowStyle=0&DaySelect=';" & vbCrLf
    Response.Write "var strLinkb='"", ""strLesson"", ""width=600,height=450,toolbar=yes,menubar=yes, scrollbars=yes, resizable=yes, location=yes, status=yes"");\'>';" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "var a=0;" & vbCrLf
    Response.Write "var oldLesson="""";" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==""""){return false;}" & vbCrLf
    Response.Write "    if(a==0){" & vbCrLf
    Response.Write "        opener.myform.all(DaySelect[0]).value="""";" & vbCrLf
    Response.Write "        a=1;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "    if(opener.myform.all(DaySelect[0]).value==""""){" & vbCrLf
    Response.Write "        for(i=0;i<parentTd.length;i++){" & vbCrLf
    Response.Write "            opener.myform.all(DaySelect[i]).value=obj;          " & vbCrLf
    Response.Write "            parentTd[i].innerHTML=strLinka+DaySelect[i]+strLinkb+'<span class=""style3"">'+DaySelect[i].substring(8,DaySelect[i].length)+'</span><br><span class=""style2"">'+opener.myform.all(DaySelect[i]).value+'</span></a>';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        myform.LessonList.value=opener.myform.all(DaySelect[0]).value;  " & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var singleLesson=obj.split("","");" & vbCrLf
    Response.Write "    var ignoreLesson="""";" & vbCrLf
    Response.Write "    for(i=0;i<singleLesson.length;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(checkLesson(opener.myform.all(DaySelect[0]).value,singleLesson[i]))" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            ignoreLesson=ignoreLesson+singleLesson[i]+"" """ & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            for(j=0;j<parentTd.length;j++){" & vbCrLf
    Response.Write "                opener.myform.all(DaySelect[j]).value=opener.myform.all(DaySelect[j]).value+"",""+singleLesson[i];                          " & vbCrLf
    Response.Write "                parentTd[j].innerHTML=strLinka+DaySelect[j]+strLinkb+'<span class=""style3"">'+DaySelect[j].substring(8,DaySelect[j].length)+'</span><br><span class=""style2"">'+opener.myform.all(DaySelect[j]).value+'</span></a>';" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            myform.LessonList.value=opener.myform.all(DaySelect[0]).value;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(ignoreLesson!="""")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        alert(ignoreLesson+"" 此课时已经存在，此操作已经忽略！"");" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function del(num)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if (num==0 || opener.myform.all(DaySelect[0]).value=="""" || opener.myform.all(DaySelect[0]).value=="","")" & vbCrLf
    Response.Write "    {   " & vbCrLf
    Response.Write "        for(i=0;i<parentTd.length;i++){" & vbCrLf
    Response.Write "            opener.myform.all(DaySelect[i]).value="""";         " & vbCrLf
    Response.Write "            parentTd[i].innerHTML=strLinka+DaySelect[i]+strLinkb+'<span class=""style3"">'+DaySelect[i].substring(8,DaySelect[i].length)+'</span></a>';;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        myform.LessonList.value="""";" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "    var strDel=opener.myform.all(DaySelect[0]).value;" & vbCrLf
    Response.Write "    var s=strDel.split("","");" & vbCrLf
    Response.Write "    for(i=0;i<parentTd.length;i++){" & vbCrLf
    Response.Write "        opener.myform.all(DaySelect[i]).value=strDel.substring(0,strDel.length-s[s.length-1].length-1);     " & vbCrLf
    Response.Write "        parentTd[i].innerHTML=strLinka+DaySelect[i]+strLinkb+'<span class=""style3"">'+DaySelect[i].substring(8,DaySelect[i].length)+'</span><br><span class=""style2"">'+opener.myform.all(DaySelect[0]).value+'</span></a>';;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    myform.LessonList.value=opener.myform.all(DaySelect[0]).value;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
ElseIf ShowStyle = 2 Then '周次
    DaySelectVB = CStr(Trim(Request("DaySelect")))
    DayTdVB = Right(DaySelectVB, Len(DaySelectVB) - 3)
    ReDim A(6)
    For i = 0 To 6
        A(i) = DateAdd("d", i, CDate(DayTdVB))
    Next
    astr = Join(A, """,""")
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "var B = new Array("""&astr&""");" & vbCrLf
    Response.Write "var parentTd = new Array();" & vbCrLf
    Response.Write "var DaySelect = new Array();" & vbCrLf
    Response.Write "for (i=0; i <=6; i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "    parentTd[i] = opener.document.getElementById('sstr'+B[i]);" & vbCrLf
    Response.Write "    DaySelect[i] ='str'+ B[i];" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "var strLinka='<a href=""#"" onclick=\'window.open(""Admin_SelectTime.asp?ShowStyle=0&DaySelect=';" & vbCrLf
    Response.Write "var strLinkb='"", ""strLesson"", ""width=600,height=450,toolbar=yes,menubar=yes, scrollbars=yes, resizable=yes, location=yes, status=yes"");\'>';" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "var a=0;" & vbCrLf
    Response.Write "var oldLesson="""";" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==""""){return false;}" & vbCrLf
    Response.Write "    if(a==0){" & vbCrLf
    Response.Write "        opener.myform.all(DaySelect[0]).value="""";" & vbCrLf
    Response.Write "        a=1;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "    if(opener.myform.all(DaySelect[0]).value==""""){" & vbCrLf
    Response.Write "        for(i=0;i<parentTd.length;i++){" & vbCrLf
    Response.Write "            opener.myform.all(DaySelect[i]).value=obj;          " & vbCrLf
    Response.Write "            parentTd[i].innerHTML=strLinka+DaySelect[i]+strLinkb+'<span class=""style3"">'+DaySelect[i].substring(8,DaySelect[i].length)+'</span><br><span class=""style2"">'+opener.myform.all(DaySelect[i]).value+'</span></a>';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        myform.LessonList.value=opener.myform.all(DaySelect[0]).value;  " & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var singleLesson=obj.split("","");" & vbCrLf
    Response.Write "    var ignoreLesson="""";" & vbCrLf
    Response.Write "    for(i=0;i<singleLesson.length;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(checkLesson(opener.myform.all(DaySelect[0]).value,singleLesson[i]))" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            ignoreLesson=ignoreLesson+singleLesson[i]+"" """ & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            for(j=0;j<parentTd.length;j++){" & vbCrLf
    Response.Write "                opener.myform.all(DaySelect[j]).value=opener.myform.all(DaySelect[j]).value+"",""+singleLesson[i];                          " & vbCrLf
    Response.Write "                parentTd[j].innerHTML=strLinka+DaySelect[j]+strLinkb+'<span class=""style3"">'+DaySelect[j].substring(8,DaySelect[j].length)+'</span><br><span class=""style2"">'+opener.myform.all(DaySelect[j]).value+'</span></a>';" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            myform.LessonList.value=opener.myform.all(DaySelect[0]).value;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(ignoreLesson!="""")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        alert(ignoreLesson+"" 此课时已经存在，此操作已经忽略！"");" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function del(num)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if (num==0 || opener.myform.all(DaySelect[0]).value=="""" || opener.myform.all(DaySelect[0]).value=="","")" & vbCrLf
    Response.Write "    {   " & vbCrLf
    Response.Write "        for(i=0;i<parentTd.length;i++){" & vbCrLf
    Response.Write "            opener.myform.all(DaySelect[i]).value="""";         " & vbCrLf
    Response.Write "            parentTd[i].innerHTML=strLinka+DaySelect[i]+strLinkb+'<span class=""style3"">'+DaySelect[i].substring(8,DaySelect[i].length)+'</span></a>';;" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        myform.LessonList.value="""";" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "    var strDel=opener.myform.all(DaySelect[0]).value;" & vbCrLf
    Response.Write "    var s=strDel.split("","");" & vbCrLf
    Response.Write "    for(i=0;i<parentTd.length;i++){" & vbCrLf
    Response.Write "        opener.myform.all(DaySelect[i]).value=strDel.substring(0,strDel.length-s[s.length-1].length-1);     " & vbCrLf
    Response.Write "        parentTd[i].innerHTML=strLinka+DaySelect[i]+strLinkb+'<span class=""style3"">'+DaySelect[i].substring(8,DaySelect[i].length)+'</span><br><span class=""style2"">'+opener.myform.all(DaySelect[0]).value+'</span></a>';;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    myform.LessonList.value=opener.myform.all(DaySelect[0]).value;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
Else
    Response.Write "<script language=""javascript"">" & vbCrLf
    Response.Write "var DaySelect      ='" & Request("DaySelect") & "'" & vbCrLf
    Response.Write "var parentTd=opener.document.getElementById('s'+DaySelect)" & vbCrLf
    Response.Write "var strLink='<a href=""#"" onclick=\'window.open(""Admin_SelectTime.asp?ShowStyle=0&DaySelect='+DaySelect+'"", ""strLesson"", ""width=600,height=450,toolbar=yes,menubar=yes, scrollbars=yes, resizable=yes, location=yes, status=yes"");\'>';" & vbCrLf
    Response.Write "myform.LessonList.value=opener.myform.all(DaySelect).value;" & vbCrLf
    Response.Write "parentTd.innerHTML=strLink+'<span class=""style3"">'+DaySelect.substring(8,DaySelect.length)+'</span><br><span class=""style2"">'+myform.LessonList.value+'</span></a>';" & vbCrLf
    Response.Write "var oldLesson="""";" & vbCrLf
    Response.Write "function add(obj)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if(obj==""""){return false;}" & vbCrLf
    Response.Write "    if(opener.myform.all(DaySelect).value=="""")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        opener.myform.all(DaySelect).value=obj;" & vbCrLf
    Response.Write "        myform.LessonList.value=opener.myform.all(DaySelect).value;" & vbCrLf
    Response.Write "        parentTd.innerHTML=strLink+'<span class=""style3"">'+DaySelect.substring(8,DaySelect.length)+'</span><br><span class=""style2"">'+myform.LessonList.value+'</span></a>';;" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    var singleLesson=obj.split("","");" & vbCrLf
    Response.Write "    var ignoreLesson="""";" & vbCrLf
    Response.Write "    for(i=0;i<singleLesson.length;i++)" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        if(checkLesson(opener.myform.all(DaySelect).value,singleLesson[i]))" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            ignoreLesson=ignoreLesson+singleLesson[i]+"" """ & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        else" & vbCrLf
    Response.Write "        {" & vbCrLf
    Response.Write "            opener.myform.all(DaySelect).value=opener.myform.all(DaySelect).value+"",""+singleLesson[i];" & vbCrLf
    Response.Write "            myform.LessonList.value=opener.myform.all(DaySelect).value;" & vbCrLf
    Response.Write "            parentTd.innerHTML=strLink+'<span class=""style3"">'+DaySelect.substring(8,DaySelect.length)+'</span><br><span class=""style2"">'+myform.LessonList.value+'</span></a>';" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "    if(ignoreLesson!="""")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        alert(ignoreLesson+"" 此课时已经存在，此操作已经忽略！"");" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "function del(num)" & vbCrLf
    Response.Write "{" & vbCrLf
    Response.Write "    if (num==0 || opener.myform.all(DaySelect).value=="""" || opener.myform.all(DaySelect).value=="","")" & vbCrLf
    Response.Write "    {" & vbCrLf
    Response.Write "        opener.myform.all(DaySelect).value="""";" & vbCrLf
    Response.Write "        myform.LessonList.value="""";" & vbCrLf
    Response.Write "        parentTd.innerHTML=strLink+'<span class=""style3"">'+DaySelect.substring(8,DaySelect.length)+'</span><br><span class=""style2"">'+myform.LessonList.value+'</span></a>';;" & vbCrLf
    Response.Write "        return false;" & vbCrLf
    Response.Write "    }" & vbCrLf
    Response.Write "" & vbCrLf
    Response.Write "    var strDel=opener.myform.all(DaySelect).value;" & vbCrLf
    Response.Write "    var s=strDel.split("","");" & vbCrLf
    Response.Write "    opener.myform.all(DaySelect).value=strDel.substring(0,strDel.length-s[s.length-1].length-1);" & vbCrLf
    Response.Write "    myform.LessonList.value=opener.myform.all(DaySelect).value;" & vbCrLf
    Response.Write "    parentTd.innerHTML=strLink+'<span class=""style3"">'+DaySelect.substring(8,DaySelect.length)+'</span><br><span class=""style2"">'+myform.LessonList.value+'</span></a>';;" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "" & vbCrLf
End If
Response.Write "function checkLesson(LessonList,thisLesson)" & vbCrLf
Response.Write "{" & vbCrLf
Response.Write "  if (LessonList==thisLesson){" & vbCrLf
Response.Write "        return true;" & vbCrLf
Response.Write "  }" & vbCrLf
Response.Write "  else{" & vbCrLf
Response.Write "    var s=LessonList.split("","");" & vbCrLf
Response.Write "    for (j=0;j<s.length;j++){" & vbCrLf
Response.Write "        if(s[j]==thisLesson)" & vbCrLf
Response.Write "            return true;" & vbCrLf
Response.Write "    }" & vbCrLf
Response.Write "    return false;" & vbCrLf
Response.Write "  }" & vbCrLf
Response.Write "}" & vbCrLf
Response.Write "</script>" & vbCrLf
Response.Write "" & vbCrLf

Function PE_CLng(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng = CLng(str1)
    Else
        PE_CLng = 0
    End If
End Function

Response.Write "</fieldset>" & vbCrLf
Response.Write "</form>" & vbCrLf
%>
</body>
</html>
