<%@language=vbscript codepage=936 %>
<%
Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim ArticleID, Action, sql, rs, Hits, ShowType
ArticleID = Trim(request("ArticleID"))
Action = Trim(request("Action"))
ShowType = Trim(request("ShowType"))
If IsNumeric(ShowType) Then
    ShowType = CLng(ShowType)
Else
    ShowType = 1
End If
%>
<!--#include file="../conn.asp"-->
<!--#include file="Channel_Config.asp"-->
<%
If Action = "Count" Then
    sql = "select sum(Hits) from PE_Article where ChannelID=" & ChannelID
    Set rs = conn.execute(sql)
    If IsNull(rs(0)) Then
        Hits = 0
    Else
        Hits = rs(0)
    End If
    rs.Close
    Set rs = Nothing
Else
    If ArticleID = "" Then
        Hits = 0
    Else
        ArticleID = CLng(ArticleID)
        conn.execute ("update PE_Channel set HitsCount=HitsCount+1 where ChannelID=" & ChannelID & "")
        sql = "select Hits from PE_Article where Deleted=" & PE_False & " and Status=3 and ArticleID=" & ArticleID & " and ChannelID=" & ChannelID & ""
        Set rs = server.CreateObject("ADODB.recordset")
        rs.open sql, conn, 1, 3
        If rs.bof And rs.EOF Then
            Hits = 0
        Else
            Hits = rs(0) + 1
            rs(0) = Hits
            rs.Update
        End If
        rs.Close
        Set rs = Nothing
    End If
End If
Select Case ShowType
Case 0

Case 1
    Response.Write "document.write('" & Hits & "');"
End Select
Call CloseConn
%>