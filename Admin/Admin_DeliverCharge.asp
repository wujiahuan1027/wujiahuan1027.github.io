<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.buffer = True
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
Dim DeliverTypeID
Call Execute
Public Sub Execute()
    DeliverTypeID = PE_CLng(Trim(Request("DeliverTypeID")))
    If DeliverTypeID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��DeliverTypeID��</li>"
        Exit Sub
    End If
    Response.Write "<html>" & vbCrLf
    Response.Write "<head>" & vbCrLf
    Response.Write "<title>��ʡ�˷ѱ�׼</title>" & vbCrLf
    Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link rel='stylesheet' href='Admin_Style.css'>" & vbCrLf
    Response.Write "</head>" & vbCrLf
    Response.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
    Select Case Action
    Case "Add"
        Call Add
    Case "Modify"
        Call Modify
    Case "SaveAdd", "SaveModify"
        Call SaveCharge
    Case "Del"
        Call Del
    Case Else
        Call Main
    End Select
    If FoundErr = True Then
        Response.Write ErrMsg
    End If
    Response.Write "</body></html>" & vbCrLf
    Call CloseConn
End Sub

Private Sub Main()
    Response.Write "<table width='100%' border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg2'>" & vbCrLf
    Response.Write "    <td>ʡ���б�</td>" & vbCrLf
    Response.Write "    <td width='60'>�����˷�</td>" & vbCrLf
    Response.Write "    <td width='60'>��������</td>" & vbCrLf
    Response.Write "    <td width='60'>��λ�˷�</td>" & vbCrLf
    Response.Write "    <td width='60'>��λ����</td>" & vbCrLf
    Response.Write "    <td width='60'>����˷�</td>" & vbCrLf
    Response.Write "    <td width='60'>����</td>" & vbCrLf
    Response.Write "  </tr>" & vbCrLf
    Dim rsChargeList, arrProvince
    arrProvince = ""
    Set rsChargeList = Conn.Execute("select * from PE_DeliverCharge where AreaType=4 and DeliverTypeID=" & DeliverTypeID & "")
    If rsChargeList.bof And rsChargeList.EOF Then
        Response.Write "<tr class='tdbg'><td colspan='10' height='40' align='center'>Ŀǰû���κ���ʡ�˷ѱ�׼</td></tr>"
    Else
        Do While Not rsChargeList.EOF
            If arrProvince = "" Then
                arrProvince = rsChargeList("arrArea")
            Else
                arrProvince = arrProvince & "," & rsChargeList("arrArea")
            End If
            Response.Write "  <tr class='tdbg'>" & vbCrLf
            Response.Write "    <td>" & rsChargeList("arrArea") & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='right'>��" & FormatNumber(rsChargeList("Charge_Min"), 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='right'>" & FormatNumber(rsChargeList("Weight_Min"), 2, vbTrue, vbFalse) & "Kg</td>" & vbCrLf
            Response.Write "    <td width='60' align='right'>��" & FormatNumber(rsChargeList("ChargePerUnit"), 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='right'>" & FormatNumber(rsChargeList("WeightPerUnit"), 2, vbTrue, vbFalse, vbTrue) & "Kg</td>" & vbCrLf
            Response.Write "    <td width='60' align='right'>��" & FormatNumber(rsChargeList("Charge_Max"), 2, vbTrue, vbFalse, vbTrue) & "</td>" & vbCrLf
            Response.Write "    <td width='60' align='center'><a href='#' onClick=""window.open('Admin_DeliverCharge.asp?DeliverTypeID=" & DeliverTypeID & "&Action=Modify&ID=" & rsChargeList("ID") & "','Charge','height=360, width=640');"">�޸�</a> <a href='Admin_DeliverCharge.asp?Action=Del&DeliverTypeID=" & DeliverTypeID & "&ID=" & rsChargeList("ID") & "' onclick=""return confirm('ȷ��Ҫɾ�����˷ѱ�׼��');"">ɾ��</a></td>" & vbCrLf
            Response.Write "  </tr>" & vbCrLf
            rsChargeList.MoveNext
        Loop
    End If
    Set rsChargeList = Nothing
    Response.Write "</table>" & vbCrLf
    
    Dim rsProvince, strProvince
    strProvince = ""
    Set rsProvince = Conn.Execute("select DISTINCT Province from PE_City")
    Do While Not rsProvince.EOF
        If FoundInArr(arrProvince, rsProvince(0), ",") = False Then
            If strProvince = "" Then
                strProvince = rsProvince(0)
            Else
                strProvince = strProvince & "," & rsProvince(0)
            End If
        End If
        rsProvince.MoveNext
    Loop
    Set rsProvince = Nothing
    If strProvince <> "" Then
        Response.Write "<b>��������ʡ��û���趨�˷ѱ�׼��</b><br>" & strProvince & "<br>"
        Response.Write "<div align='center'><input type='button' name='Submit' value=' �����ʡ�˷ѱ�׼ ' onClick=""window.open('Admin_DeliverCharge.asp?DeliverTypeID=" & DeliverTypeID & "&Action=Add','Charge','height=360, width=640');""></div>" & vbCrLf
    End If
End Sub

Sub Add()
    Response.Write "<form name='myform' method='post' action='Admin_DeliverCharge.asp'>" & vbCrLf
    Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
    Response.Write "    <tr align='center' class='topbg'>" & vbCrLf
    Response.Write "      <td colspan='2'><b>�� �� �� ʡ �� �� �� ׼</b></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "    <tr class='tdbg'>" & vbCrLf
    Response.Write "      <td width='180' align='center'><b>����ѡ���ʡ�ݣ�" & vbCrLf
    Response.Write "          <select name='arrArea' size='2' multiple style='height:200;width:150 '>" & GetProvince("")
    Response.Write "        </select>" & vbCrLf
    Response.Write "        <br>" & vbCrLf
    Response.Write "        ��Ctrl��Shift�����Զ�ѡ" & vbCrLf
    Response.Write "</b></td>" & vbCrLf
    Response.Write "      <td valign='top'><table width='100%'  border='0' cellpadding='2' cellspacing='1'>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width='140' align='right'><b>�����˷ѣ�</b></td>" & vbCrLf
    Response.Write "          <td><input name='Charge_Min' type='text' id='Charge_Min' value='10' size='10' maxlength='10' style='text-align:center '>" & vbCrLf
    Response.Write "      Ԫ</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width='140' align='right' valign='top'><b>�����˷ѵ�����������</b></td>" & vbCrLf
    Response.Write "          <td><input name='Weight_Min' type='text' id='Weight_Min' value='1' size='10' maxlength='10' style='text-align:center '>" & vbCrLf
    Response.Write "      ǧ�ˣ�Kg��<br>" & vbCrLf
    Response.Write "      ����Ʒ��������������ָ����������ʱ��ʵ���˷Ѱ������˷Ѽ��㡣</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width='140' align='right' valign='top'><b>�����˷ѣ�</b></td>" & vbCrLf
    Response.Write "          <td>����Ʒ���������������˷ѵ����������󣬳�����ȡ�����˷��⣬<br>" & vbCrLf
    Response.Write "            ÿ" & vbCrLf
    Response.Write "              <input name='WeightPerUnit' type='text' id='WeightPerUnit' value='1' size='6' maxlength='6' style='text-align:center '>" & vbCrLf
    Response.Write "      ǧ�˵���Ʒ�����˷�" & vbCrLf
    Response.Write "      <input name='ChargePerUnit' type='text' id='ChargePerUnit' value='5' size='10' maxlength='10' style='text-align:center '>" & vbCrLf
    Response.Write "      Ԫ</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "        <tr>" & vbCrLf
    Response.Write "          <td width='140' align='right' valign='top'><b>����˷ѣ�</b></td>" & vbCrLf
    Response.Write "          <td><input name='Charge_Max' type='text' id='Charge_Max' value='100' size='10' maxlength='10' style='text-align:center '>" & vbCrLf
    Response.Write "      Ԫ���������˷ѣ������˷ѳ�������˷�ʱ��ʵ���˷Ѱ�����˷Ѽ��㣩</td>" & vbCrLf
    Response.Write "        </tr>" & vbCrLf
    Response.Write "      </table>" & vbCrLf
    Response.Write "      <p align='center'>" & vbCrLf
    Response.Write "        <input name='DeliverTypeID' type='hidden' id='DeliverTypeID' value='" & DeliverTypeID & "'>" & vbCrLf
    Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveAdd'>" & vbCrLf
    Response.Write "        <input type='submit' name='Submit' value=' �� �� '>" & vbCrLf
    Response.Write "&nbsp;&nbsp;&nbsp;        " & vbCrLf
    Response.Write "<input type='button' name='Submit' value=' ȡ �� ' onclick='window.close()'>" & vbCrLf
    Response.Write "      </p>" & vbCrLf
    Response.Write "      <p align='left'>���������ѡ���ʡ�ݡ��б���û��������Ҫ��ʡ�ݣ����������Ϊ���ʡ���������˷ѱ�׼���Ѿ����ڣ�����Ҫ���޸������˷ѱ�׼��ȥ����Ӧ��ʡ�ݣ�Ȼ���������˷ѱ�׼��</p></td>" & vbCrLf
    Response.Write "    </tr>" & vbCrLf
    Response.Write "  </table>" & vbCrLf
    Response.Write "</form>" & vbCrLf
End Sub
Sub Modify()
    Dim ID
    Dim rsCharge
    ID = PE_CLng(Trim(Request("ID")))
    If ID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��ID��</li>"
        Exit Sub
    End If
    Set rsCharge = Conn.Execute("select * from PE_DeliverCharge where ID=" & ID & "")
    If rsCharge.bof And rsCharge.EOF Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>�Ҳ���ָ���ļ�¼��</li>"
    Else
        Response.Write "<form name='myform' method='post' action='Admin_DeliverCharge.asp'>" & vbCrLf
        Response.Write "  <table width='100%'  border='0' cellpadding='2' cellspacing='1' class='border'>" & vbCrLf
        Response.Write "    <tr align='center' class='topbg'>" & vbCrLf
        Response.Write "      <td colspan='2'><b>�� �� �� ʡ �� �� �� ׼</b></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "    <tr class='tdbg'>" & vbCrLf
        Response.Write "      <td width='180' align='center'><b>����ѡ���ʡ�ݣ�" & vbCrLf
        Response.Write "          <select name='arrArea' size='2' multiple style='height:200;width:150 '>" & GetProvince(rsCharge("arrArea"))
        Response.Write "        </select>" & vbCrLf
        Response.Write "        <br>" & vbCrLf
        Response.Write "        ��Ctrl��Shift�����Զ�ѡ" & vbCrLf
        Response.Write "</b></td>" & vbCrLf
        Response.Write "      <td valign='top'><table width='100%'  border='0' cellpadding='2' cellspacing='1'>" & vbCrLf
        Response.Write "        <tr>" & vbCrLf
        Response.Write "          <td width='140' align='right'><b>�����˷ѣ�</b></td>" & vbCrLf
        Response.Write "          <td><input name='Charge_Min' type='text' id='Charge_Min' value='" & FormatNumber(rsCharge("Charge_Min"), 2, vbTrue, vbFalse, vbTrue) & "' size='10' maxlength='10' style='text-align:center '>" & vbCrLf
        Response.Write "      Ԫ</td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "        <tr>" & vbCrLf
        Response.Write "          <td width='140' align='right' valign='top'><b>�����˷ѵ�����������</b></td>" & vbCrLf
        Response.Write "          <td><input name='Weight_Min' type='text' id='Weight_Min' value='" & FormatNumber(rsCharge("Weight_Min"), 2, vbTrue, vbFalse, vbTrue) & "' size='10' maxlength='10' style='text-align:center '>" & vbCrLf
        Response.Write "      ǧ�ˣ�Kg��<br>" & vbCrLf
        Response.Write "      ����Ʒ��������������ָ����������ʱ��ʵ���˷Ѱ������˷Ѽ��㡣</td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "        <tr>" & vbCrLf
        Response.Write "          <td width='140' align='right' valign='top'><b>�����˷ѣ�</b></td>" & vbCrLf
        Response.Write "          <td>����Ʒ���������������˷ѵ����������󣬳�����ȡ�����˷��⣬<br>" & vbCrLf
        Response.Write "            ÿ" & vbCrLf
        Response.Write "              <input name='WeightPerUnit' type='text' id='WeightPerUnit' value='" & FormatNumber(rsCharge("WeightPerUnit"), 2, vbTrue, vbFalse, vbTrue) & "' size='6' maxlength='6' style='text-align:center '>" & vbCrLf
        Response.Write "      ǧ�˵���Ʒ�����˷�" & vbCrLf
        Response.Write "      <input name='ChargePerUnit' type='text' id='ChargePerUnit' value='" & FormatNumber(rsCharge("ChargePerUnit"), 2, vbTrue, vbFalse, vbTrue) & "' size='10' maxlength='10' style='text-align:center '>" & vbCrLf
        Response.Write "      Ԫ</td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "        <tr>" & vbCrLf
        Response.Write "          <td width='140' align='right' valign='top'><b>����˷ѣ�</b></td>" & vbCrLf
        Response.Write "          <td><input name='Charge_Max' type='text' id='Charge_Max' value='" & FormatNumber(rsCharge("Charge_Max"), 2, vbTrue, vbFalse, vbTrue) & "' size='10' maxlength='10' style='text-align:center '>" & vbCrLf
        Response.Write "      Ԫ���������˷ѣ������˷ѳ�������˷�ʱ��ʵ���˷Ѱ�����˷Ѽ��㣩</td>" & vbCrLf
        Response.Write "        </tr>" & vbCrLf
        Response.Write "      </table>" & vbCrLf
        Response.Write "      <p align='center'>" & vbCrLf
        Response.Write "        <input name='ID' type='hidden' id='ID' value='" & ID & "'>" & vbCrLf
        Response.Write "        <input name='DeliverTypeID' type='hidden' id='DeliverTypeID' value='" & DeliverTypeID & "'>" & vbCrLf
        Response.Write "        <input name='Action' type='hidden' id='Action' value='SaveModify'>" & vbCrLf
        Response.Write "        <input type='submit' name='Submit' value=' �� �� '>" & vbCrLf
        Response.Write "&nbsp;&nbsp;&nbsp;        " & vbCrLf
        Response.Write "<input type='button' name='Submit' value=' ȡ �� ' onclick='window.close()'>" & vbCrLf
        Response.Write "      </p>" & vbCrLf
        Response.Write "      <p align='left'>���������ѡ���ʡ�ݡ��б���û��������Ҫ��ʡ�ݣ����������Ϊ���ʡ���������˷ѱ�׼���Ѿ����ڣ�����Ҫ���޸������˷ѱ�׼��ȥ����Ӧ��ʡ�ݣ�Ȼ���������˷ѱ�׼��</p></td>" & vbCrLf
        Response.Write "    </tr>" & vbCrLf
        Response.Write "  </table>" & vbCrLf
        Response.Write "</form>" & vbCrLf
    End If
    Set rsCharge = Nothing
End Sub

Sub SaveCharge()
    Dim ID, arrArea, Charge_Min, Weight_Min, ChargePerUnit, WeightPerUnit, Charge_Max
    Dim rsCharge, sqlCharge
    ID = PE_CLng(Trim(Request("ID")))
    arrArea = Replace(Trim(Request("arrArea")), " ", "")
    Charge_Min = PE_CDbl(Trim(Request("Charge_Min")))
    Weight_Min = PE_CDbl(Trim(Request("Weight_Min")))
    ChargePerUnit = PE_CDbl(Trim(Request("ChargePerUnit")))
    WeightPerUnit = PE_CDbl(Trim(Request("WeightPerUnit")))
    Charge_Max = PE_CDbl(Trim(Request("Charge_Max")))
    FoundErr = False
    If arrArea = "" Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��ʡ�ݣ�</li>"
    End If
    If Charge_Min > Charge_Max Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>����˷Ѳ���С�ڻ����˷ѣ�</li>"
    End If
    If WeightPerUnit <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��λ����Ӧ�ô��ڣ���</li>"
    End If
    If Action = "SaveAdd" Then
        sqlCharge = "select top 1 * from PE_DeliverCharge"
    Else
        If ID <= 0 Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>��ָ��ID��</li>"
        Else
            sqlCharge = "select * from PE_DeliverCharge where ID=" & ID
        End If
    End If
    If FoundErr = True Then
        Exit Sub
    End If
    
    Set rsCharge = Server.CreateObject("adodb.recordset")
    rsCharge.open sqlCharge, Conn, 1, 3
    If Action = "SaveAdd" Then
        rsCharge.addnew
        rsCharge("DeliverTypeID") = DeliverTypeID
        rsCharge("AreaType") = 4
    Else
        If rsCharge.bof And rsCharge.EOF Then
            FoundErr = True
            ErrMsg = ErrMsg & "<li>�Ҳ���ָ�����˷ѱ�׼��Ŀ��</li>"
        End If
    End If
    If FoundErr = False Then
        rsCharge("arrArea") = arrArea
        rsCharge("Charge_Min") = Charge_Min
        rsCharge("Weight_Min") = Weight_Min
        rsCharge("ChargePerUnit") = ChargePerUnit
        rsCharge("WeightPerUnit") = WeightPerUnit
        rsCharge("Charge_Max") = Charge_Max
        rsCharge.Update
    End If
    rsCharge.Close
    Set rsCharge = Nothing
    If FoundErr = False Then
        Response.Write "<script>window.opener.location.reload();window.close();</script>"
    End If
End Sub
Sub Del()
    Dim ID
    Dim rsCharge
    ID = PE_CLng(Trim(Request("ID")))
    If ID <= 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>��ָ��ID��</li>"
        Exit Sub
    End If
    Conn.Execute ("delete from PE_DeliverCharge where ID=" & ID & "")
    Response.redirect "Admin_DeliverCharge.asp?DeliverTypeID=" & DeliverTypeID
End Sub

Private Function GetProvince(arrArea)
    Dim rsProvince, strProvince, rsProvinceExists, arrProvince
    Dim IsExists, IsInArr
    arrProvince = ""
    Set rsProvinceExists = Conn.Execute("select arrArea from PE_DeliverCharge where DeliverTypeID=" & DeliverTypeID & " and AreaType=4")
    Do While Not rsProvinceExists.EOF
        If arrProvince = "" Then
            arrProvince = rsProvinceExists(0)
        Else
            arrProvince = arrProvince & "," & rsProvinceExists(0)
        End If
        rsProvinceExists.MoveNext
    Loop
    Set rsProvinceExists = Nothing
    
    Set rsProvince = Conn.Execute("select DISTINCT Province from PE_City")
    Do While Not rsProvince.EOF
        IsExists = FoundInArr(arrProvince, rsProvince(0), ",")
        IsInArr = FoundInArr(arrArea, rsProvince(0), ",")
        If IsExists = False Or IsInArr = True Then
            strProvince = strProvince & "<option value='" & rsProvince(0) & "'"
            If IsInArr = True Then
                strProvince = strProvince & " selected"
            End If
            strProvince = strProvince & ">" & rsProvince(0) & "</option>"
        End If
        rsProvince.MoveNext
    Loop
    Set rsProvince = Nothing
    GetProvince = strProvince
End Function
%>
