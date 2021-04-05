<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="../Conn.asp"-->
<%
Dim ChannelID, ShowType, Site_Sn
Dim FilesPath, sql, rs
Dim AdminName,UserName,UserPassword,LastPassword,UserSetting
Dim sqlChannel,rsChannel
Dim ModuleType,DialogType,IsUpload

Site_Sn = Replace(Replace(LCase(request.ServerVariables("SERVER_NAME") & InstallDir), "/", ""), ".", "")

AdminName = ReplaceBadChar(Trim(request.Cookies(Site_Sn)("AdminName")))
UserName = ReplaceBadChar(Trim(request.Cookies(Site_Sn)("UserName")))
UserPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("UserPassword")))
LastPassword = ReplaceBadChar(Trim(Request.Cookies(Site_Sn)("LastPassword")))

ChannelID = Trim(request("ChannelID"))
ShowType = PE_Clng(Trim(request("ShowType")))

If ChannelID = "" Then
    response.write "频道参数丢失！"
    response.End
Else
    ChannelID = PE_CLng(ChannelID)
End If

If AdminName = "" And UserName = "" Then
    Response.Write "请先登录后再使用此功能！"
    Response.End
Else
    If AdminName <> "" And (ShowType=0 or ShowType=4 or ShowType=5) Then
        IsUpload = True
    ElseIf UserName <> "" Then
        If (UserName = "" Or UserPassword = "" Or LastPassword = "") Then
            IsUpload = False
        Else
            sql = "SELECT U.UserID,U.SpecialPermission,U.UserSetting,G.GroupSetting FROM PE_User U inner join PE_UserGroup G on U.GroupID=G.GroupID WHERE"
            sql = sql & " UserName='" & UserName & "' AND UserPassword='" & UserPassword & "' AND LastPassword='" & LastPassword & "' and IsLocked=" & PE_False & ""
            Set rs = Conn.Execute(sql)
            If rs.BOF And rs.EOF Then
                IsUpload = False
            Else
                If rs("SpecialPermission") = True Then
                    UserSetting = Split(Trim(rs("UserSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                Else
                    UserSetting = Split(Trim(rs("GroupSetting")) & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0", ",")
                End If
                If CBool(PE_CLng(UserSetting(9))) = True Then
                    IsUpload = True
                End If
            End If
            Set rs = Nothing
        End If
    End If
End If

sqlChannel = "select ChannelDir,UploadDir,Disabled,EnableUploadFile,ModuleType from PE_Channel where ChannelID=" & ChannelID
Set rsChannel = Server.CreateObject("adodb.recordset")
rsChannel.Open sqlChannel, Conn, 1, 1
If rsChannel.BOF And rsChannel.EOF Then
    IsUpload = False
Else
    If rsChannel("Disabled") = True Then
        IsUpload = False
    Else
        If rsChannel("EnableUploadFile") = False Then
            IsUpload = False
        End If
        FilesPath = InstallDir & rsChannel("ChannelDir") & "/" & rsChannel("UploadDir") & "/"
        ModuleType = rsChannel("ModuleType")
    End If
End If
rsChannel.Close
Set rsChannel = Nothing

If IsUpload = True Then
    Select Case ModuleType
    Case 1, 2, 3, 5, 6, 7, 8 
        IsUpload = True
    Case Else
        IsUpload = False
    End Select
End If


Call CloseConn


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
%>