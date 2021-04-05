<%@ Language = "VBScript" Codepage = "936" %>
<!--#include file="../conn.asp"-->
<!--#include file="../Inc/function.asp"-->
<!--#include file="../Inc/Md5.asp"-->
<!--#include file="../API/API_Config.asp"-->
<!--#include file="../API/API_Function.asp"-->
<%

Call User_CheckReg
If FoundErr = True Then
    Call WriteErrMsg(ErrMsg, ComeUrl)
End If
Call CloseConn

Sub User_CheckReg()
    Dim RegUserName
    RegUserName = Trim(request("UserName"))
    If InStr(RegUserName, "=") > 0 Or InStr(RegUserName, "%") > 0 Or InStr(RegUserName, Chr(32)) > 0 Or InStr(RegUserName, "?") > 0 Or InStr(RegUserName, "&") > 0 Or InStr(RegUserName, ";") > 0 Or InStr(RegUserName, ",") > 0 Or InStr(RegUserName, "'") > 0 Or InStr(RegUserName, ",") > 0 Or InStr(RegUserName, Chr(34)) > 0 Or InStr(RegUserName, Chr(9)) > 0 Or InStr(RegUserName, "��") > 0 Or InStr(RegUserName, "$") > 0 Or InStr(RegUserName, "*") Or InStr(RegUserName, "|") Or InStr(RegUserName, """") > 0 Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�û����к��зǷ��ַ�</li>"
    End If
    If FoundErr = True Then Exit Sub

    Dim rsCheckReg, UserName_RegDisabled
    Set rsCheckReg = Conn.Execute("select UserNameLimit,UserNameMax,UserName_RegDisabled from PE_Config")
    If Not (rsCheckReg.bof And rsCheckReg.EOF) Then
        UserNameLimit = rsCheckReg(0)
        UserNameMax = rsCheckReg(1)
        UserName_RegDisabled = rsCheckReg(2)
    Else
        UserNameLimit = 4
        UserNameMax = 20
    End If
    Set rsCheckReg = Nothing

    If RegUserName = "" Or strLength(RegUserName) > UserNameMax Or strLength(RegUserName) < UserNameLimit Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>�������û���(���ܴ���" & UserNameMax & "С��" & UserNameLimit & ")</li>"
    End If

    If FoundInArr(UserName_RegDisabled, RegUserName, "|") = True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��������û���Ϊϵͳ��ֹע����û�����</li>"
    End If

    Set rsCheckReg = Conn.Execute("select UserName from PE_User where UserName='" & RegUserName & "'")
    If Not (rsCheckReg.bof And rsCheckReg.EOF) Then
        FoundErr = True
        ErrMsg = ErrMsg & "<br><li>��" & RegUserName & "���Ѿ����ڣ��뻻һ���û��������ԣ�</li>"
    End If

    rsCheckReg.Close
    Set rsCheckReg = Nothing
    If FoundErr = True Then Exit Sub    
    
    '��Ӷ����Ͻӿڵ�֧��
    If API_Enable Then
        sPE_Items(conAction,1) = "checkname"
        sPE_Items(conUsername,1) = RegUserName
        If createXmlDom Then
            prepareXml True
            SendPost
            If FoundErr Then
                ErrMsg = "<li>" & ErrMsg & "</li>" & vbNewLine
            End If
        Else
            FoundErr = True
            ErrMsg = "<li>��������֧��MSXML����ע����񲻿���! [APIError-XmlDom-Runtime]</li>" & vbNewLine
        End If
    End If
    '���
    If FoundErr = True Then Exit Sub

    Call WriteSuccessMsg("��" & RegUserName & "�� ��δ����ʹ�ã��Ͻ�ע��ɣ�", ComeUrl)
End Sub
%>