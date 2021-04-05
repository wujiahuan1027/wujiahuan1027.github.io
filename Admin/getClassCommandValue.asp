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
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
    Dim strSql, ClassID
    ClassID = PE_CLng(Trim(Request("ClassID")))
    strSql = "Select CommandClassPoint From PE_Class Where ClassID=" & ClassID & ""
    Call OpenConn
    Response.Write PE_CLng(Conn.Execute(strSql)(0))
    Call CloseConn
%>
