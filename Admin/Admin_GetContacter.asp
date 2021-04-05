<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private"
Response.CacheControl = "no-cache"
Response.Charset = "GB2312"
Response.ContentType = "text/html"


Const PurviewLevel = 0
Const PurviewLevel_Channel = 0
Const PurviewLevel_Others = ""
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<!--#include file="Admin_ChkPurview.asp"-->

<%
Response.Write "<select name='ContacterID'>" & vbCrLf
Dim rsContacter
ClientID = PE_CLng(Trim(Request("ClientID")))
Set rsContacter = Conn.Execute("select ContacterID,TrueName from PE_Contacter where ClientID=" & ClientID & "")
Do While Not rsContacter.EOF
    Response.Write "<option value='" & rsContacter(0) & "'>" & rsContacter(1) & "</option>" & vbCrLf
    rsContacter.movenext
Loop
rsContacter.Close
Set rsContacter = Nothing
Call CloseConn
Response.Write "</select>" & vbCrLf

%>
