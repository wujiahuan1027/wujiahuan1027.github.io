<!--#include file="Conn_Counter.asp"-->
<%
Dim RegCount_Fill,IntervalNum,OnlineTime
dim style,theurl
if IsEmpty(Application("RegFields_Fill")) Or IsEmpty(Application("IntervalNum")) then
	call OpenConn_Counter()
	Dim rs
	set Rs=conn_counter.execute("select * from PE_StatInfoList")
	if not Rs.bof and not Rs.eof then
		IntervalNum=Rs("IntervalNum")
		RegCount_Fill = rs("RegFields_Fill")
		OnlineTime=Rs("OnlineTime")
		Application("OnlineTime")=OnlineTime
		Application("IntervalNum")=IntervalNum
		Application("RegFields_Fill") = RegCount_Fill
	end if
	set Rs=nothing
	Call CloseConn_counter()
else
	IntervalNum=Application("IntervalNum")
	RegCount_Fill=Application("RegFields_Fill")
end if

style=Request("style")
theurl="http://" & Request.ServerVariables("http_host") & finddir(Request.ServerVariables("url"))
if right(theurl,1)<>"/" then
	theurl=theurl & "/"
end if
%>

var style      ='<%=style%>';
var url        ='<%=theurl%>';
var IntervalNum=<%=IntervalNum%>;
var i=0;
<%
If FoundInArr(RegCount_Fill, "IsCountOnline", ",") = True Then
%>
PowerEasyRef(0);
<%
End if
%>
document.write("<scr"+"ipt language=javascript src="+url+"counter.asp?style="+style+"&Referer="+escape(document.referrer)+"&Timezone="+escape((new Date()).getTimezoneOffset())+"&Width="+escape(screen.width)+"&Height="+escape(screen.height)+"&Color="+escape(screen.colorDepth)+"></sc"+"ript>");
function PowerEasyRef(){
	if(i <= IntervalNum){
		var PowerEasyImg=new Image();
		PowerEasyImg.src=url+'statonline.asp';
		setTimeout('PowerEasyRef()',60000);
	}
	i+=1;
}
