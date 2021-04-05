<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit
Response.Buffer = True
Server.ScriptTimeOut = 9999999
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
If CheckUserLogined() = False Then
    Call CloseConn
    Response.Redirect "User_Login.asp"
End If
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SiteName & " >> 会员中心"%></title>
<link href="../Skin/DefaultSkin.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include file="top.asp"-->
<table width="756" border="0" align="center" cellpadding="0" cellspacing="0" class="user_border">
  <tr>
    <td valign="top">
      <table width="100%" border="0" cellpadding="5" cellspacing="0" class="user_box">
        <tr>
          <td class="user_righttitle"><img src="Images/point2.gif" align="absmiddle"><%
Response.Write "您现在的位置：<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>会员中心</a> >> 下载软件"
          %></td>
        </tr>
        <tr>
          <td height="200" valign='top'>
<%Call ShowSoftList%>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<!--#include file="bottom.asp"-->
</body>
</html>
<%
Call CloseConn

Sub ShowSoftList()
    Dim rsOrderItem, i
    Dim AllowDownload

    Set rsOrderItem = Conn.Execute("select P.ProductID,P.ProductName,P.ProductModel,I.Price,I.TruePrice,I.Amount,P.Unit,I.BeginDate,I.ServiceTerm,P.Remark,P.DownloadUrl from PE_OrderFormItem I inner join PE_Product P on I.ProductID=P.ProductID where I.OrderFormID in (select OrderFormID from PE_OrderForm where UserName='" & UserName & "' and EnableDownload=" & PE_True & ") and P.ProductKind=2 order by I.ItemID")
    If rsOrderItem.bof And rsOrderItem.EOF Then
        Response.Write "<br><p align='center'>您还没有购买任何软件产品或购买的产品还没有开通下载。</p>"
    Else
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border3'>" & vbCrLf
        Response.Write "  <tr align='center' class='tdbg3'>" & vbCrLf
        Response.Write "    <td width='200' height='25'><b>商 品 名 称</b></td>" & vbCrLf
        Response.Write "    <td width='40'><b>单 位</b></td>" & vbCrLf
        Response.Write "    <td width='40'><b>数 量</b></td>" & vbCrLf
        Response.Write "    <td width='60'><b>购买日期</b></td>" & vbCrLf
        Response.Write "    <td width='60'><b>服务期限</b></td>" & vbCrLf
        Response.Write "    <td width='44'><b>已到期</b></td>" & vbCrLf
        Response.Write "    <td><b>备注/说明</b></td>" & vbCrLf
        Response.Write "    <td width='60'><b>下载地址</b></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Do While Not rsOrderItem.EOF
            AllowDownload = True

            Response.Write "  <tr valign='middle' class='tdbg3'>"
            Response.Write "    <td width='200' height='20'>" & rsOrderItem("ProductName") & "</td>"
            Response.Write "    <td width='40' align='center'>" & rsOrderItem("Unit") & "</td>"
            Response.Write "    <td width='40' align='center'>" & rsOrderItem("Amount") & "</td>"
            Response.Write "    <td width='60' align='center'>" & rsOrderItem("BeginDate") & "</td>"
            Response.Write "    <td width='60' align='center'>"
            Select Case rsOrderItem("ServiceTerm")
            Case -1
                Response.Write "无限期"
            Case 1
                Response.Write "一年"
            Case 2
                Response.Write "两年"
            Case 3
                Response.Write "三年"
            Case Else
                Response.Write "未知"
            End Select
            Response.Write "    </td>"
            Response.Write "    <td width='44' align='center'>"
            If rsOrderItem("ServiceTerm") > 0 Then
                If DateAdd("yyyy", rsOrderItem("ServiceTerm"), rsOrderItem("BeginDate")) <= Now() Then
                    AllowDownload = False
                    Response.Write "<font color=red>到期</font>"
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td align='center'>" & rsOrderItem("Remark") & "</td>"
            Response.Write "    <td width='60' align=center>"
            If AllowDownload = True Then
                Response.Write "<a href='" & rsOrderItem("DownloadUrl") & "' target='_blank'>点此下载</a>"
            End If
            Response.Write "</td></tr>"
            i = i + 1
            rsOrderItem.MoveNext
        Loop
    End If
    rsOrderItem.Close
    Set rsOrderItem = Nothing
End Sub
%>
