<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
Server.ScriptTimeOut = 9999999
%>
<!--#include file="../Conn.asp"-->
<%

Sub PE_Execute(strDllName, strClassName)
    On Error Resume Next
    If strDllName = "" Or IsNull(strDllName) Then
        Response.Write "��ָ�������������"
        Exit Sub
    End If
    If strClassName = "" Or IsNull(strClassName) Then
        Response.Write "��ָ����������ṩ��������"
        Exit Sub
    End If
    Dim PE_User, objName
    objName = strDllName & "." & strClassName
    Set PE_User = Server.CreateObject(objName)
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������" & strDllName & ".dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_User.iConnStr = ConnStr
    PE_User.iSystemDatabaseType = SystemDatabaseType
    If strClassName = "User_Enrol" Then
        PE_User.iStartDay = "2006-1-1"
    End If
    Call PE_User.Execute
    Set PE_User = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub

%>
