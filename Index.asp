<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="conn.asp"-->
<%
If FileExt_SiteIndex < 4 Then
    Call CloseConn
    Response.redirect "Index" & GetFileExt(FileExt_SiteIndex)
Else
    Call CloseConn
    Call Main
End If

Function GetFileExt(FileExtType)
    Select Case FileExtType
    Case 0
        GetFileExt = ".html"
    Case 1
        GetFileExt = ".htm"
    Case 2
        GetFileExt = ".shtml"
    Case 3
        GetFileExt = ".shtm"
    Case 4
        GetFileExt = ".asp"
    End Select
End Function

Sub Main()
    On Error Resume Next
    Dim PE_Index
    Set PE_Index = Server.CreateObject("PE_CMS6.CreateIndex")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_Index.iConnStr = ConnStr
    PE_Index.iSystemDatabaseType = SystemDatabaseType
    Call PE_Index.ShowHTML
    Set PE_Index = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>