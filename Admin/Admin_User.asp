<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
Server.ScriptTimeOut = 9999999
%>
<!--#include file="../Conn.asp"-->
<!--#include file="Admin_ChkCode.asp"-->
<!--#include file="../API/API_Config.asp"-->
<%
Call CheckSiteManageCode
Call CloseConn
Call PE_Execute("PE_CMS6", "Admin_User")

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
    Dim PE_Admin, objName
    objName = strDllName & "." & strClassName
    Set PE_Admin = Server.CreateObject(objName)
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������" & strDllName & ".dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_Admin.iConnStr = ConnStr
    PE_Admin.iCMS_Edition = CMS_Edition
    PE_Admin.iSystemDatabaseType = SystemDatabaseType
    PE_Admin.iAPI_Key = API_Key
    PE_Admin.iAPI_Urls = API_Urls
    PE_Admin.iAPI_Enable = API_Enable
    PE_Admin.iAPI_Timeout = API_Timeout
    Call PE_Admin.Execute
    Set PE_Admin = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>