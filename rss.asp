<!--#include file="conn.asp"-->
<%
Call CloseConn

If RssCodeType = True Then
    Response.ContentType = "text/xml; charset=gb2312"
Else
    Response.ContentType = "text/xml; charset=utf-8"
End If
Call Main

Sub Main()
    On Error Resume Next
    Dim PE_Index
    Set PE_Index = Server.CreateObject("PE_CMS6.ShowRss")
    If Err Then
        Err.Clear
        If RssCodeType = True Then
            Response.Write "�Բ�����ķ�����û�а�װ���������PE_CMS6.dll�������Բ���ʹ�ö���ϵͳ�������Ŀռ�����ϵ�԰�װ���������"
        Else
            Response.Write "&#x5BF9;&#x4E0D;&#x8D77;&#xFF0C;&#x4F60;&#x7684;&#x670D;&#x52A1;&#x5668;&#x6CA1;&#x6709;&#x5B89;&#x88C5;&#x52A8;&#x6613;&#x7EC4;&#x4EF6;&#xFF08;PE_CMS6.dll&#xFF09;&#xFF0C;&#x6240;&#x4EE5;&#x4E0D;&#x80FD;&#x4F7F;&#x7528;&#x52A8;&#x6613;&#x7CFB;&#x7EDF;&#x3002;&#x8BF7;&#x548C;&#x4F60;&#x7684;&#x7A7A;&#x95F4;&#x5546;&#x8054;&#x7CFB;&#x4EE5;&#x5B89;&#x88C5;&#x52A8;&#x6613;&#x7EC4;&#x4EF6;&#x3002;"
        End If
        Exit Sub
    End If
    PE_Index.iConnStr = ConnStr
    PE_Index.iSystemDatabaseType = SystemDatabaseType
    Call PE_Index.Main
    Set PE_Index = Nothing
    If Err Then
        Response.Write "�� �� �ţ�" & Err.Number & "<BR>"
        Response.Write "����������" & Err.Description & "<BR>"
        Response.Write "������Դ��" & Err.Source & "<BR>"
        Err.Clear
    End If
End Sub
%>