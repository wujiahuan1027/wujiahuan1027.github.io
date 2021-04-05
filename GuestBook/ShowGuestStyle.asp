<%
dim ShowGStyle
ShowGStyle=request("ShowGStyle")
if ShowGStyle="" or not isNumeric(ShowGStyle) then
	ShowGStyle=1
else
	ShowGStyle=int(ShowGStyle)
end if
response.cookies("ShowGStyle")=ShowGStyle
response.redirect request.servervariables("http_referer")
%>