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
<title><%=SiteName & " >> ��Ա����"%></title>
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
Response.Write "�����ڵ�λ�ã�<a href='../'>" & SiteName & "</a> >> <a href='Index.asp'>��Ա����</a> >> �������"
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
        Response.Write "<br><p align='center'>����û�й����κ������Ʒ����Ĳ�Ʒ��û�п�ͨ���ء�</p>"
    Else
        Response.Write "<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border3'>" & vbCrLf
        Response.Write "  <tr align='center' class='tdbg3'>" & vbCrLf
        Response.Write "    <td width='200' height='25'><b>�� Ʒ �� ��</b></td>" & vbCrLf
        Response.Write "    <td width='40'><b>�� λ</b></td>" & vbCrLf
        Response.Write "    <td width='40'><b>�� ��</b></td>" & vbCrLf
        Response.Write "    <td width='60'><b>��������</b></td>" & vbCrLf
        Response.Write "    <td width='60'><b>��������</b></td>" & vbCrLf
        Response.Write "    <td width='44'><b>�ѵ���</b></td>" & vbCrLf
        Response.Write "    <td><b>��ע/˵��</b></td>" & vbCrLf
        Response.Write "    <td width='60'><b>���ص�ַ</b></td>" & vbCrLf
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
                Response.Write "������"
            Case 1
                Response.Write "һ��"
            Case 2
                Response.Write "����"
            Case 3
                Response.Write "����"
            Case Else
                Response.Write "δ֪"
            End Select
            Response.Write "    </td>"
            Response.Write "    <td width='44' align='center'>"
            If rsOrderItem("ServiceTerm") > 0 Then
                If DateAdd("yyyy", rsOrderItem("ServiceTerm"), rsOrderItem("BeginDate")) <= Now() Then
                    AllowDownload = False
                    Response.Write "<font color=red>����</font>"
                End If
            End If
            Response.Write "</td>"
            Response.Write "    <td align='center'>" & rsOrderItem("Remark") & "</td>"
            Response.Write "    <td width='60' align=center>"
            If AllowDownload = True Then
                Response.Write "<a href='" & rsOrderItem("DownloadUrl") & "' target='_blank'>�������</a>"
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
