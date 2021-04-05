<%
dim action,PageTitle,strHtml
PageTitle="请选择头像"

strHtml="<html><head><title>" & PageTitle & "</title></head>" & vbcrlf
strHTML=strHTML & "<script>" & vbcrlf
strHTML=strHTML & "window.focus();" & vbcrlf
strHTML=strHTML & "function changeimage(imagename){" & vbcrlf
strHTML=strHTML & "  window.opener.document.myform.GuestImages.value=imagename;" & vbcrlf
strHTML=strHTML & "  window.opener.document.myform.showimages.src='../guestbook/images/face/'+imagename+'.gif';" & vbcrlf
strHTML=strHTML & "}" & vbcrlf
strHTML=strHTML & "</script>" & vbcrlf
strHTML=strHTML & "<body>" & vbcrlf

strHTML=strHTML & "<table align=center width=95% cellpadding=5><tr><td>"
for i=1 to 22
	strHTML=strHTML & "<img src='images/face/"
	if i<10 then
		i= "0" & i
	end if
	strHTML=strHTML & i & ".gif' border=0 onclick=""changeimage('" & i & "') "" style=cursor:hand>"
	if i mod 5 =0 then
		strHTML=strHTML & "<br>"
	end if
next
strHTML=strHTML & "</td></tr></table>"

strHTML=strHTML & "<div align='center'><font size='2'>【<a href='javascript:window.close();'>关闭窗口</a>】</font></div>"
strHTML=strHTML & "</body>"
strHTML=strHTML & "</html>"
response.write strHTML
%>


