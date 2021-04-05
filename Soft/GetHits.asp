<%@language=vbscript codepage=936 %>
<%
Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Dim SoftID, Action, HitsType, sql, rs, Hits
SoftID = Trim(request("SoftID"))
Action = Trim(request("Action"))
HitsType = Trim(request("HitsType"))
If HitsType = "" Then
    HitsType = 0
Else
    HitsType = CLng(HitsType)
End If

%>
<!--#include file="../conn.asp"-->
<!--#include file="Channel_Config.asp"-->
<%
If Action = "Count" Then
    sql = "select sum(Hits) from PE_Soft where ChannelID=" & ChannelID
    Set rs = conn.execute(sql)
    If IsNull(rs(0)) Then
        Hits = 0
    Else
        Hits = rs(0)
    End If
    rs.Close
    Set rs = Nothing
ElseIf Action = "SoftDown" Then
    Hits = ""
    If SoftID = "" Then
        SoftID = 0
    Else
        SoftID = CLng(SoftID)
    End If
    Dim rsSoft
    sql = "select * from PE_Soft where Deleted=" & PE_False & " and Status=3 and SoftID=" & SoftID & " and ChannelID=" & ChannelID & ""
    Set rsSoft = server.CreateObject("ADODB.Recordset")
    rsSoft.open sql, conn, 1, 3
    If Not (rsSoft.bof And rsSoft.bof) Then
        rsSoft("Hits") = rsSoft("Hits") + 1
        If DateDiff("D", rsSoft("LastHitTime"), Now()) <= 0 Then
            rsSoft("DayHits") = rsSoft("DayHits") + 1
        Else
            rsSoft("DayHits") = 1
        End If
        If DateDiff("ww", rsSoft("LastHitTime"), Now()) <= 0 Then
            rsSoft("WeekHits") = rsSoft("WeekHits") + 1
        Else
            rsSoft("WeekHits") = 1
        End If
        If DateDiff("m", rsSoft("LastHitTime"), Now()) <= 0 Then
            rsSoft("MonthHits") = rsSoft("MonthHits") + 1
        Else
            rsSoft("MonthHits") = 1
        End If
        rsSoft("LastHitTime") = Now()
        rsSoft.Update
    End If
    rsSoft.Close
    Set rsSoft = Nothing
Else
    If SoftID = "" Then
        Hits = 0
    Else
        SoftID = CLng(SoftID)
        Select Case HitsType
        Case 0
            sql = "select Hits from PE_Soft where SoftID=" & SoftID
        Case 1
            sql = "select DayHits from PE_Soft where SoftID=" & SoftID
        Case 2
            sql = "select WeekHits from PE_Soft where SoftID=" & SoftID
        Case 3
            sql = "select MonthHits from PE_Soft where SoftID=" & SoftID
        Case Else
            conn.execute ("update PE_Channel set HitsCount=HitsCount+1 where ChannelID=" & ChannelID & "")
            sql = "select Hits from PE_Soft where SoftID=" & SoftID
        End Select
        Set rs = server.CreateObject("ADODB.recordset")
        rs.open sql, conn, 1, 1
        If rs.bof And rs.EOF Then
            Hits = 0
        Else
            Hits = rs(0)
        End If
        rs.Close
        Set rs = Nothing
    End If
End If
Response.Write "document.write('" & Hits & "');"
Call CloseConn
%>