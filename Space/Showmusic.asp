<!--#include file="../Conn.asp"-->
<%
Response.Expires = -1
Response.ContentType = "text/xml; charset=gb2312"
Call CloseConn
Call Main

Sub Main()
    On Error Resume Next
    Dim PE_Blog
    Set PE_Blog = Server.CreateObject("PE_CMS6.ShowBlog")
    If Err Then
        Err.Clear
        Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Exit Sub
    End If
    PE_Blog.iConnStr = ConnStr
    PE_Blog.iSystemDatabaseType = SystemDatabaseType
    PE_Blog.iShowType = "ShowMusic"
    Call PE_Blog.ShowHTML
    Set PE_Blog = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>