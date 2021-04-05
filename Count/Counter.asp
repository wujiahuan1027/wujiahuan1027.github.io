<!--#include file="Conn_Counter.asp"-->
<%

Dim Ip,LastIPCache,Sip,Address,Scope,Referer,VisitorKeyword,WebUrl,Visit,StatIP,strIP
Dim Agent,System,Browser,BcType,Mozilla,Height,Width,Screen,Color,Timezone,Ver,VisitTimezone
Dim StrYear,StrMonth,StrDay,StrHour,Strweek,StrHourLong,StrDayLong,StrMonthLong,OldDay
Dim Num,I,nYesterDayNum,CacheData
Dim Province,OnlineNum,ShowInfo
Dim OnNowTime,style
Dim RegCount_Fill,OnlineTime,VisitRecord,KillRefresh
Dim DayNum,AllNum,TotalView,StartDate,StatDayNum,AveDayNum

call OpenConn_Counter()

Dim Sql,Rs
set Rs=conn_counter.execute("select * from PE_StatInfoList")
if not Rs.bof and not Rs.eof then
	RegCount_Fill = rs("RegFields_Fill")
	OnlineTime=Rs("OnlineTime")
	VisitRecord=Rs("VisitRecord")
	KillRefresh=Rs("KillRefresh")
    DayNum=Rs("DayNum")
    AllNum=Rs("TotalNum")+Rs("OldTotalNum")
    TotalView=Rs("TotalView")+Rs("OldTotalView")
    StartDate=Rs("StartDate")
    StatDayNum=DateDiff("D",StartDate,Date)+1
    if StatDayNum<=0 or isnumeric(StatDayNum)=0 then
    	AveDayNum=StatDayNum
    Else
        AveDayNum=Clng(AllNum/StatDayNum)
    end if
end if
set Rs=nothing

Response.Expires = 0
LastIPCache="Powereasy_LastIP"
if isempty(Application(LastIPCache)) then Application(LastIPCache)="#0.0.0.0#"

Ip=Request.ServerVariables("REMOTE_ADDR")

If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
	If OnlineTime="" Or isnumeric(OnlineTime)=0 Then OnlineTime=100	  
    OnNowTime = dateadd("s",-OnlineTime,now())
    dim rsOnline
	if CountDatabaseType="SQL" then
		set rsonline = conn_counter.execute("select count(UserIP) from PE_Statonline where LastTime>'"&OnNowTime&"'")
	else
		set rsonline = conn_counter.execute("select count(UserIP) from PE_Statonline where LastTime>#"&OnNowTime&"#")
	end if
    OnlineNum = rsonline(0)		' 当前在线人数
    Set rsonline=Nothing 
	if CountDatabaseType="SQL" then
		Set rsonline = conn_counter.execute("select LastTime,OnTime from PE_Statonline where LastTime>'"&OnNowTime&"' and UserIP='"&IP&"'")
    else
		Set rsonline = conn_counter.execute("select LastTime,OnTime from PE_Statonline where LastTime>#"&OnNowTime&"# and UserIP='"&IP&"'")
	end if
	If rsOnline.eof then
		Update()
	Else
		if rsonline(0)=rsonline(1) Then
			Update()
		else
			conn_counter.Execute("Update PE_StatInfoList set TotalView=TotalView+1")
		end if
	End If
	Set rsonline=Nothing
Else
    if instr(Application(LastIPCache),"#" & IP & "#") then	' 如果IP已经存在于保存的列表中，是刷新	
        conn_counter.Execute("Update PE_StatInfoList set TotalView=TotalView+1")
    Else
    	Application.Lock 
    	Application(LastIPCache)=SaveIP(Application(LastIPCache))		' 更新最近需要防刷的IP
    	Application.UnLock
    	Update()
    End If
End If


style=lcase(trim(Request("style")))
select case style
case "simple"
	ShowInfo="总访问量：" & AllNum & "人次<br>"
	If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
		ShowInfo=ShowInfo&"当前在线：" & OnlineNum & "人"
	End If
case "all"
	ShowInfo=ShowInfo&"总访问量：" & AllNum & "人次<br>"
	ShowInfo=ShowInfo&"总浏览量：" & TotalView & "人次<br>"
'	ShowInfo=ShowInfo&"统计天数：" & StatDayNum & "天<br>"
	If FoundInArr(RegCount_Fill, "FYesterDay", ",") = True Then
		Call GetYesterdayNum()
		ShowInfo=ShowInfo&"昨日访问：" & nYesterDayNum & "人<br>"
	end if
	ShowInfo=ShowInfo&"今日访问：" & DayNum & "人次<br>"
	ShowInfo=ShowInfo&"日均访问：" & AveDayNum & "人次<br>"
	If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
		ShowInfo=ShowInfo&"当前在线：" & OnlineNum & "人"
	End If
case "common"
	ShowInfo="总访问量：" & AllNum & "人次<br>"
	ShowInfo=ShowInfo&"总浏览量：" & TotalView & "人次<br>"
	If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
		ShowInfo=ShowInfo&"当前在线：" & OnlineNum & "人"
	End If
end select
if style<>"none" then
	Response.Write "document.write(" & chr(34) & ShowInfo & chr(34) & ")"
end if

Call CloseConn_counter()
sub Update()
	If FoundInArr(RegCount_Fill, "FIP", ",") = True Then
		strIP=split(IP,".")
		if isNumeric(strIP(0))=0 or isNumeric(strIP(1))=0 or isNumeric(strIP(2))=0 or isNumeric(strIP(3))=0 then
			Sip=0
		else
			Sip=cint(strIP(0))*256*256*256+cint(strIP(1))*256*256+cint(strIP(2))*256+cint(strIP(3))-1
		end if
		if (167772159 < Sip and Sip< 184549374) or (2886729727 < Sip and Sip < 2887778302) or (3232235519 < Sip and Sip < 3232301054) then
			StatIP=IP
		else
			StatIP=strIP(0) & "." & strIP(1) & ".*"
		end if
	else
		StatIP=""
	End If
	SIp=Ip
	set rs=server.createobject("adodb.recordset")
	if SIP="127.0.0.1" then
		Address="本机地址"
		Scope="ChinaNum"
	else
		strIP=split(Sip,".")
		if isNumeric(strIP(0))=0 or isNumeric(strIP(1))=0 or isNumeric(strIP(2))=0 or isNumeric(strIP(3))=0 then
			Sip=0
		else
			Sip=cint(strIP(0))*256*256*256+cint(strIP(1))*256*256+cint(strIP(2))*256+cint(strIP(3))-1
		end if

		dim RsAdress
		set RsAdress=conn_counter.execute("Select Top 1 Address From PE_StatIpInfo Where StartIp<="&Sip&" and EndIp>="&Sip&" Order By EndIp-StartIp Asc")
		If RsAdress.Eof Then
			Address="其它地区"
		Else
			Address=RsAdress(0)
		End If
		set RsAdress=nothing
		Province="北京天津上海重庆黑龙江吉林辽宁江苏浙江安徽河南河北湖南湖北山东山西内蒙古陕西甘肃宁夏青海新疆西藏云南贵州四川广东广西福建江西海南香港澳门台湾内部网未知"
		if instr(Province,left(Address,2))>0 then
			Scope="ChinaNum"
		Else
			Scope="OtherNum"
		End if
	end if

	Referer=Request.QueryString("Referer")	
	If Referer="" Then Referer="直接输入或书签导入"
	Referer=left(Referer,100)

		'response.write"11="&Referer
		'response.end

	If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then
		WebUrl=Left(Request.QueryString("Referer"),Instr(8,Referer,"/"))
		If WebUrl="" Then WebUrl="直接输入或书签导入"
		WebUrl=Left(WebUrl,50)
	else
		WebUrl=""
	End If

	Width=Request.QueryString("Width")
	Height=Request.QueryString("Height")
	If Height="" Or isnumeric(Height)=0 Or Width="" Or isnumeric(Width)=0 Then 
		Screen="其它"
	Else
		Screen=Cstr(Width)&"x"&Cstr(Height)
	End If
	Screen=left(Screen,10)



	Color=Request.QueryString("Color")
	If Color="" Or isnumeric(Color)=0 Then 
		Color="其它"
	Else
		Select Case Color
		Case 4:
			 Color="16 色"
		Case 8:
			 Color="256 色"
		Case 16:
			 Color="增强色（16位）"
		Case 24:
			 Color="真彩色（24位）"
		Case 32:
			 Color="真彩色（32位）"
		End Select
	End If


	Mozilla=Request.ServerVariables("HTTP_USER_AGENT")
	Mozilla=left(Mozilla,100)
	Agent=Request.ServerVariables("HTTP_USER_AGENT")
	Agent=Split(Agent,";")
	BcType=0
	If Instr(Agent(1),"U") Or Instr(Agent(1),"I") Then BcType=1
	If InStr(Agent(1),"MSIE") Then BcType=2
	Select Case BcType
	Case 0:
		 Browser="其它"
		 System="其它"
	Case 1:
		 Ver=Mid(Agent(0),InStr(Agent(0),"/")+1)
		 Ver=Mid(Ver,1,InStr(Ver," ")-1)
		 Browser="Netscape"&Ver
		 System=Mid(Agent(0),InStr(Agent(0),"(")+1)		 
	case 2:
		 Browser=Agent(1)
		 System=Agent(2)
		 System=Replace(System,")","")		 
	End Select
	System=Replace(Replace(Replace(Replace(Replace(Replace(System," ",""),"Win","Windows"),"NT5.0","2000"),"NT5.1","XP"),"NT5.2","2003"),"dowsdows","dows")
	Browser=Replace(Browser," ","")
	System=Left(System,20)

	Browser=Left(Browser,20)

	Timezone=Request.QueryString("Timezone")
	If Timezone="" Or isnumeric(Timezone)=0 Then 
	   Timezone="其它"
	   VisitTimezone=0
	Else
		VisitTimezone=Timezone\60
		If Timezone<0 Then
			Timezone="GMT+"&Abs(Timezone)\60&":"&(Abs(Timezone) Mod 60)
		Else
			Timezone="GMT-"&Abs(Timezone)\60&":"&(Abs(Timezone) Mod 60)
		End If
	End If


	If FoundInArr(RegCount_Fill, "FVisit", ",") = True Then
		Visit=Request.Cookies("VisitNum")
		If Visit<>"" Then
			Visit=Visit+1
		Else
			Visit=1
		End If
		Response.Cookies("VisitNum")=Visit
		Response.Cookies("VisitNum").Expires="January 01, 2010"
		Sql="Select * From PE_StatVisit"
		Rs.Open Sql,conn_counter,1,3
		If Rs.Eof Or Rs.Bof Then
			Rs.AddNew
		End If
		If Visit<=10 Then
			If Isnumeric(Rs(Visit-1))=0 Then
				Rs(Visit-1)=1
			Else
				Rs(Visit-1)=Rs(Visit-1)+1
				If Visit>1 Then
				   If Rs(Visit-2)>0 Then Rs(Visit-2)=Rs(Visit-2)-1
				End If
			End If
		End If
		Rs.Update
		Rs.Close
	end if

	Call UpdateVisit()

	StrHour=Cstr(hour(time))
	StrDay=Cstr(Day(Date))
	StrMonth=Cstr(Month(Date))
	StrYear=Cstr(Year(Date))
	StrWeek=Cstr(Weekday(Date))
	StrDayLong=Cstr(Year(Date)&"-"&Month(Date)&"-"&Day(date))
	StrMonthLong=Cstr(Year(Date)&"-"&Month(Date))
	StrHourLong=StrDayLong&" "&Cstr(Hour(Time))&":00:00"

	Sql="Select * From PE_StatInfoList"
	Rs.Open Sql,conn_counter,1,3
	Rs("TotalNum")=Rs("TotalNum")+1
	Rs("TotalView")=Rs("TotalView")+1
	Rs(Scope)=Rs(Scope)+1
	If IsNull(Rs("StartDate")) Then Rs("StartDate")=StrDayLong
	If IsNull(Rs("OldDay")) Then Rs("OldDay")=StrDayLong
	OldDay=Rs("OldDay")
	Rs.Update
	Rs.Close
	Call ModiMaxNum

	If VisitorKeyword<> "" And FoundInArr(RegCount_Fill, "FKeyword", ",") = True then  
		VisitorKeyword=FindKeystr(Request.QueryString("Referer"))
		VisitorKeyword=replace(trim(Lcase(VisitorKeyword)),"'","")
		AddNum VisitorKeyword,"PE_Statkeyword","Tkeyword","TkeywordNum"
	End If
	If FoundInArr(RegCount_Fill, "FSystem", ",") = True Then
		AddNum System,"PE_StatSystem","TSystem","TSysNum"
	End If
	If FoundInArr(RegCount_Fill, "FBrowser", ",") = True Then
		AddNum Browser,"PE_StatBrowser","TBrowser","TBrwNum"
	End If
	If FoundInArr(RegCount_Fill, "FMozilla", ",") = True Then
		AddNum Mozilla,"PE_StatMozilla","TMozilla","TMozNum"
	End If
	If FoundInArr(RegCount_Fill, "FScreen", ",") = True Then
		AddNum Screen,"PE_StatScreen","TScreen","TScrNum"
	End If
	If FoundInArr(RegCount_Fill, "FColor", ",") = True Then
		AddNum Color,"PE_StatColor","TColor","TColNum"
	End If
	If FoundInArr(RegCount_Fill, "FTimezone", ",") = True Then
		AddNum Timezone,"PE_StatTimezone","TTimezone","TTimNum"
	End If
	If FoundInArr(RegCount_Fill, "FRefer", ",") = True Then
		AddNum Referer,"PE_StatRefer","TRefer","TRefNum"
	End If
	If FoundInArr(RegCount_Fill, "FWeburl", ",") = True Then
		AddNum Weburl,"PE_StatWeburl","TWeburl","TWebNum"
	End If
	If FoundInArr(RegCount_Fill, "FAddress", ",") = True Then
		AddNum Address,"PE_StatAddress","TAddress","TAddNum"
	End If
	If FoundInArr(RegCount_Fill, "FIP", ",") = True Then
		AddNum Ip,"PE_StatIp","TIp","TIpNum"
	End If

	AddNum StrDayLong,"PE_StatDay","TDay",StrHour
	AddNum "Total","PE_StatDay","TDay",StrHour
	AddNum StrYear,"PE_StatYear","TYear",StrMonth
	AddNum "Total","PE_StatYear","TYear",StrMonth
	AddNum StrMonthLong,"PE_StatMonth","TMonth",StrDay
	AddNum "Total","PE_StatMonth","TMonth",StrDay
	AddNum "Total","PE_StatWeek","TWeek",StrWeek
	If DateDiff("Ww",Cdate(OldDay),Date)>0 Then
	    Sql="Delete From PE_StatWeek Where TWeek='Current'"
	    conn_counter.Execute(Sql)
	End If
	AddNum "Current","PE_StatWeek","TWeek",StrWeek
end sub

Sub AddNum(Data,TableName,CompareField,AddField)
	Dim RowCount
	conn_counter.execute "update "&TableName&" set ["&AddField&"]=["&AddField&"]+1 where  "&CompareField&"='"&Data&"'", RowCount
	If RowCount = 0 Then conn_counter.execute "insert into "&TableName&" ("&CompareField&",["&AddField&"]) values ('"&Data&"',1)"
End Sub

Sub ModiMaxNum()
    Sql="Select * From PE_StatInfoList"
    Rs.Open Sql,conn_counter,1,3
    If Rs("OldMonth")=StrMonthLong Then
        Rs("MonthNum")=Rs("MonthNum")+1
    Else
        Rs("OldMonth")=StrMonthLong
        Rs("MonthNum")=1
    End If
    If Rs("MonthNum")>Rs("MonthMaxNum") Then 
        Rs("MonthMaxNum")=Rs("MonthNum")
        Rs("MonthMaxDate")=StrMonthLong
    End If
	If Rs("OldDay")=StrDayLong Then
        Rs("DayNum")=Rs("DayNum")+1
    Else
        Rs("OldDay")=StrDayLong
        Rs("DayNum")=1
    End If
    If Rs("DayNum")>Rs("DayMaxNum") Then 
        Rs("DayMaxNum")=Rs("DayNum")
        Rs("DayMaxDate")=StrDayLong
    End If
	If Rs("OldHour")=StrHourLong Then
        Rs("HourNum")=Rs("HourNum")+1
    Else
        Rs("OldHour")=StrHourLong
        Rs("HourNum")=1
    End If
    If Rs("HourNum")>Rs("HourMaxNum") Then 
        Rs("HourMaxNum")=Rs("HourNum")
        Rs("HourMaxTime")=StrHourLong
    End If
    Rs.Update
    Rs.Close
End Sub

Sub UpdateVisit()
    Dim rsOut,VisitCount,OutNum
    VisitCount = 0
    Set rsOut = conn_counter.Execute("select count(ID) From PE_StatVisitor")

    VisitCount = rsOut(0)
    If VisitCount >= VisitRecord Then
		dim rsOd
		set rsOd = conn_counter.Execute("select top 1 VTime from PE_StatVisitor order by VTime asc")
		if CountDatabaseType="SQL" then
			conn_counter.Execute("update PE_StatVisitor set VTime='"&Now()&"',IP='"&IP&"',Address='"&Address&"',Browser='"&Browser&"',System='"&System&"',Screen='"&Screen&"',Color='"&Color&"',Timezone="&VisitTimezone&",Referer='"&Referer&"' where VTime='" & rsOd("VTime") & "'")
		else
			conn_counter.Execute("update PE_StatVisitor set VTime='"&Now()&"',IP='"&IP&"',Address='"&Address&"',Browser='"&Browser&"',System='"&System&"',Screen='"&Screen&"',Color='"&Color&"',Timezone="&VisitTimezone&",Referer='"&Referer&"' where VTime=#" & rsOd("VTime") & "#")
		end if
		Set rsOd = Nothing
	else
		conn_counter.Execute  "insert into PE_StatVisitor (VTime,IP,Address,Browser,System,Screen,Color,Timezone,Referer) Values('"&Now()&"','"&IP&"','"&Address&"','"&Browser&"','"&System&"','"&Screen&"','"&Color&"',"&VisitTimezone&",'"&Referer&"')"
	End If
    Set rsOut = Nothing
End Sub

Function SaveIP(InIP)
	SaveIP=left(InIP,len(InIP)-1)
	SaveIP=right(SaveIP,len(SaveIP)-1)
	Dim FriendIP
	FriendIP=split(SaveIP,"#")
	if ubound(FriendIP) < KillRefresh then
		SaveIP="#" & SaveIP & "#" & IP & "#"
	else
		SaveIP=replace("#" & SaveIP,"#" & FriendIP(0) & "#","#") & "#" & IP & "#"
	end if
End Function

' 从URL中获取关键词
Function FindKeystr(urlstr)
    dim regEx,vKey,findKeystr1
    FindKeystr=""
    set regEx=new regexp
    regEx.Global = true
    regEx.IgnoreCase = true
    regEx.Pattern = "(?:yahoo.+?[\?|&]p=|openfind.+?q=|google.+?q=|lycos.+?query=|aol.+?query=|onseek.+?keyword=|search\.tom.+?word=|search\.qq\.com.+?word=|zhongsou\.com.+?word=|search\.msn\.com.+?q=|yisou\.com.+?p=|sina.+?word=|sina.+?query=|sina.+?_searchkey=|sohu.+?word=|sohu.+?key_word=|sohu.+?query=|163.+?q=|baidu.+?word=|3721\.com.+?name=|Alltheweb.+?q=|3721\.com.+?p=|baidu.+?wd=)([^&]*)"
  
    set Matches = regEx.Execute(urlstr)
    for each Match in Matches
  	    findKeystr1 = regEx.replace(Match.value,"$1")
    next
  
    if findKeystr1<> "" then
        FindKeystr=lcase(decodeURI(findkeystr1))
        if FindKeystr = "undefined" then
  	        FindKeystr = URLDecode(findKeystr1)
        end if
    end if
End Function

' 解开URL编码的函数(这是别人写的,地方标注为: 来源： CSDN  作者： dyydyy )
Function URLDecode(enStr)
    dim deStr
    dim c,i,v
    deStr=""
    for i=1 to len(enStr)
        c=Mid(enStr,i,1)
        if c="%" then
            v=eval("&h"+Mid(enStr,i+1,2))
            if v<128 then
                deStr=deStr&chr(v)
                i=i+2
            else
                if isvalidhex(mid(enstr,i,3)) then
					if isvalidhex(mid(enstr,i+3,3)) then
  					    v=eval("&h"+Mid(enStr,i+1,2)+Mid(enStr,i+4,2))
  					    deStr=deStr&chr(v)
  					    i=i+5
					else
					    v=eval("&h"+Mid(enStr,i+1,2)+cstr(hex(asc(Mid(enStr,i+3,1)))))
					    deStr=deStr&chr(v)
					    i=i+3 
					end if 
                else 
					destr=destr&c
                end if
            end if
        else
			if c="+" then
				deStr=deStr&" "
            else
				deStr=deStr&c
            end if
        end if
    next
    URLDecode=deStr
End Function
Function GetYesterdayNum()
	If CacheIsEmpty("nYesterDayVisitorNum") Then
		dim YesterdayStrLong
		YesterdayStrLong=year(dateadd("d","-1",date()))&"-"&month(dateadd("d","-1",date()))&"-"&day(dateadd("d","-1",date()))
		set rs=server.createobject("adodb.recordset")
		If CountDatabaseType="SQL" Then
			sql="SELECT * FROM PE_StatDay WHERE TDay='"&YesterdayStrLong&"'"
		Else
			sql="SELECT * FROM PE_StatDay WHERE TDay=#"&YesterdayStrLong&"#"
		End If
		rs.Open sql,conn_counter,1,1
		If Not rs.BOF or Not rs.EOF then
			for i=0 to 23
				nYesterDayNum=nYesterDayNum+rs(CStr(i))
			next
		else
			nYesterDayNum=0
		end if
		CacheData = Application("nYesterDayVisitorNum")
		If IsArray(CacheData) Then
			CacheData(0) = nYesterDayNum
			CacheData(1) = Now()
		Else
			ReDim CacheData(2)
			CacheData(0) = nYesterDayNum
			CacheData(1) = Now()
		End If
		Application.Lock
		Application("nYesterDayVisitorNum") = CacheData
		Application.UnLock	
	Else
		CacheData = Application("nYesterDayVisitorNum")
		If IsArray(CacheData) Then
			nYesterDayNum = CacheData(0)
		Else
			nYesterDayNum = 0
		End If
	End If
end Function

Function CacheIsEmpty(MyCacheName)
    CacheIsEmpty = True
    CacheData = Application(MyCacheName)
    If Not IsArray(CacheData) Then Exit Function
    If Not IsDate(CacheData(1)) Then Exit Function
    If DateDiff("s", CDate(CacheData(1)), Now()) < 60 * 1440 Then
        CacheIsEmpty = False
    End If
End Function
%>
<script language="javascript" runat="server" type="text/javascript">	
//解码URI
function decodeURI(furl){
	var a=furl;
	try{return decodeURIComponent(a)}catch(e){return 'undefined'};
	return '';
}
</script>