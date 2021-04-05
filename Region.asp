<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#Include file="conn.asp"-->
<%
Response.Buffer = True

Dim Country, Province, City
Dim ShowCountry, ShowProvince, ShowCity
Dim TempConn, TempRs

Country = ReplaceBadChar(Trim(Request("Country")))
Province = ReplaceBadChar(Trim(Request("Province")))
City = ReplaceBadChar(Trim(Request("City")))
If Country = "" Then Country = "中华人民共和国"
If Province = "" Then Province = "北京市"
If City = "" Then City = "海淀区"

'On Error Resume Next
Call OpenConn
Set TempRs = Conn.Execute("SELECT Country FROM PE_Country ORDER BY Country")
If Err Or TempRs.EOF Then
    FoundErr = True
Else
    ShowCountry = TempRs.GetRows(-1)
End If
Set TempRs = Conn.Execute("SELECT Province FROM PE_Province WHERE Country='" & Country & "' ORDER BY ProvinceID")
If Err Or TempRs.EOF Then
    ReDim ShowProvince(0, 0)
    Province = Trim(Request.QueryString("Province"))
Else
    ShowProvince = TempRs.GetRows(-1)
End If
Set TempRs = Conn.Execute("SELECT DISTINCT City FROM PE_City WHERE Province='" & Province & "'")
If Err Or TempRs.EOF Then
    ReDim ShowCity(0, 0)
Else
    ShowCity = TempRs.GetRows(-1)
End If
Set TempRs = Nothing
Call CloseConn
If FoundErr = True Then
    Response.Write "Conn连接错误，请检查Conn连接！" & vbCrLf
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="<%=AdminDir%>/Admin_Style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="100%"  border="0" cellpadding="2" cellspacing="1" bgcolor="#FFFFFF">
    <form name="regionform" id="regionform" action="Region.asp" method="post">
    <tr class="tdbg">
        <td width="100" align="right" class="tdbg5">
            国家/地区：
        </td>
        <td colspan="2">
            <select name="Country" id="Country" onChange="document.regionform.submit();">
                <%
                Dim i
                i = 0
                For i = 0 To UBound(ShowCountry, 2)
                    Response.Write "<option value='" & ShowCountry(0, i) & "'"
                    If Country = ShowCountry(0, i) Then Response.Write " selected"
                    Response.Write ">" & ShowCountry(0, i) & "</option>" & vbCrLf
                Next
                %>
            </select>
        </td>
    </tr>
    <tr class="tdbg">
        <td align="right" class="tdbg5">
            省/市/自治区：
        </td>
        <td>
            <select name="Province" id="Province" onChange="document.regionform.submit();">
                <%
                If ShowProvince(0, 0) = "" Then
                    Response.Write "<option>请输入</option>" & vbCrLf
                Else
                    i = 0
                    For i = 0 To UBound(ShowProvince, 2)
                        Response.Write "<option value='" & ShowProvince(0, i) & "'"
                        If Province = ShowProvince(0, i) Then Response.Write " selected"
                        Response.Write ">" & ShowProvince(0, i) & "</option>" & vbCrLf
                    Next
                End If
                %>
            </select>
        </td>
    </tr>
    <tr class="tdbg">
        <td align="right" class="tdbg5">
            市/县/区/旗：
        </td>
        <td>
            <select name="City" id="City">
                <%
                If ShowCity(0, 0) = "" Then
                    Response.Write "<option>请输入</option>" & vbCrLf
                Else
                    i = 0
                    For i = 0 To UBound(ShowCity, 2)
                        Response.Write "<option value='" & ShowCity(0, i) & "'"
                        If City = ShowCity(0, i) Then Response.Write " selected"
                        Response.Write ">" & ShowCity(0, i) & "</option>" & vbCrLf
                    Next
                End If
                %>
            </select>
        </td>
    </tr>
    </form>
</table>
</body>
</html>
<%

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
%>