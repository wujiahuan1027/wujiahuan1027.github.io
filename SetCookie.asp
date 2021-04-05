<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="inc/function.asp"-->

<%
Dim SkinID
Action = Trim(Request("Action"))
ComeUrl = Request.ServerVariables("HTTP_REFERER")
SkinID = Trim(Request("SkinID"))

If Action = "SetSkin" Then
    If SkinID = "" Then
        SkinID = 0
    Else
        SkinID = CLng(SkinID)
    End If
    Response.Cookies("asp163")("SkinID") = SkinID
End If
Response.Redirect ComeUrl
%>