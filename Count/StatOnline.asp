<!--#include file="conn_counter.asp"-->

<%
Dim RegCount_Fill,OnlineTime
Response.Expires = 0
call OpenConn_Counter()

if IsEmpty(Application("OnlineTime")) then
	Dim rs
	set Rs=conn_counter.execute("select * from PE_StatInfoList")
	if not Rs.bof and not Rs.eof then
		OnlineTime=Rs("OnlineTime")
		Application("OnlineTime")=OnlineTime
	end if
	set Rs=nothing
else
	OnlineTime=Application("OnlineTime")
end if

Dim PE_IP,PE_Agent,PE_Page,OnNowTime
PE_IP			= Request.ServerVariables("Remote_Addr")
PE_Agent		= Request.ServerVariables("HTTP_USER_AGENT")
PE_Page		= Request.ServerVariables("HTTP_REFERER")

PE_Agent		= replace(PE_Agent,"'","")
PE_Page		= replace(PE_Page,"'","")

OnNowTime = dateadd("s",0-OnlineTime,now())
'response.write "now="&now()&"OnNowTime="&OnNowTime&"OnlineTime="&OnlineTime
dim rsOnline,rsOd
if CountDatabaseType="SQL" then
	set rsOnline = conn_counter.Execute("select * from PE_StatOnline where LastTime>'"&OnNowTime&"' and UserIP='"&PE_IP&"'")
else
	set rsOnline = conn_counter.Execute("select * from PE_StatOnline where LastTime>#"&OnNowTime&"# and UserIP='"&PE_IP&"'")
end if
if rsOnline.eof then
	if CountDatabaseType="SQL" then
		set rsOd = conn_counter.Execute("select top 1 id from PE_StatOnline where LastTime<'"&OnNowTime&"'  order by LastTime")
	else
		set rsOd = conn_counter.Execute("select top 1 id from PE_StatOnline where LastTime<#"&OnNowTime&"#  order by LastTime")
	end if
	if rsOd.eof then
		conn_counter.Execute "insert into PE_StatOnline (UserIP,UserAgent,UserPage,OnTime,LastTime) VALUES('"&PE_IP&"','"&PE_Agent&"','"&PE_Page&"'," & PECount_Now & "," & PECount_Now & ")"
	else
		conn_counter.Execute "update PE_StatOnline set UserIP='"&PE_IP&"',UserAgent='"&PE_Agent&"',UserPage='"&PE_Page&"',Ontime=" & PECount_Now & ",LastTime=" & PECount_Now & " where id=" & rsOd("id")
	end if
	Set rsOd=Nothing 
else
	if CountDatabaseType="SQL" then
		conn_counter.Execute("update PE_StatOnline set LastTime=" & PECount_Now & ",UserPage='"&PE_Page&"' where LastTime>'"&OnNowTime&"' and UserIP='"&PE_IP&"'" )
	else
		conn_counter.Execute("update PE_StatOnline set LastTime=" & PECount_Now & ",UserPage='"&PE_Page&"' where LastTime>#"&OnNowTime&"# and UserIP='"&PE_IP&"'" )
	end if
end if
Set rsOnline=Nothing 
Call CloseConn_counter()
Server.Transfer("Image/powereasyimg.gif")

%>
