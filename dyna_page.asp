<!--#include file="conn.asp"-->
<%
Call CloseConn
Response.ContentType = "text/xml; charset=gb2312"
Call Main

Sub Main()
    On Error Resume Next
    Dim PE_Index
    Set PE_Index = Server.CreateObject("PE_CMS6.DynaPage")
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