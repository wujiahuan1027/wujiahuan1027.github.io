<%@language=vbscript codepage=936 %>
<%
Option Explicit
Response.Buffer = True
%>
<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
Dim strHTML, PE_Site

Call Main
Call CloseConn

Sub Main()
    Dim strPath
    If EnableUserReg <> True Then
        FoundErr = True
        ErrMsg = ErrMsg & "<li>对不起，本站暂停新用户注册服务！</li>"
        Call WriteErrMsg(ErrMsg, ComeUrl)
        Exit Sub
    End If

    Dim rsConfig, EnableCheckCodeOfReg, EnableQAofReg, QAofReg
    Set rsConfig = Conn.Execute("select top 1 * from PE_Config")
    If rsConfig.BOF And rsConfig.EOF Then
        rsConfig.Close
        Set rsConfig = Nothing
        Response.Write "网站配置数据丢失！系统无法正常运行！"
        Response.end
    Else
        EnableCheckCodeOfReg = rsConfig("EnableCheckCodeOfReg")
        EnableQAofReg = rsConfig("EnableQAofReg")
        QAofReg = rsConfig("QAofReg")
    End If
    rsConfig.Close
    Set rsConfig = Nothing

    Set PE_Site = Server.CreateObject("PE_CMS6.Site")
    PE_Site.iConnStr = ConnStr
    PE_Site.iSystemDatabaseType = SystemDatabaseType
    PE_Site.CurrentChannelID = 0
    PE_Site.Init
    
    strHTML = PE_Site.GetTemplate(0, 19, 0)
    strHTML = PE_Site.ReplaceCommon(strHTML)
    
    strPath = "您现在的位置：&nbsp;<a href='" & SiteUrl & "'>" & SiteName & "</a>&nbsp;&gt;&gt;&nbsp;新会员注册"

    strHTML = Replace(strHTML, "{$PageTitle}", SiteTitle & " >> 新会员注册")
    strHTML = Replace(strHTML, "{$ShowPath}", strPath)

    strHTML = Replace(strHTML, "{$MenuJS}", PE_Site.GetMenuJS("", False))
    strHTML = Replace(strHTML, "{$Skin_CSS}", PE_Site.GetSkin_CSS(0))
    Set PE_Site = Nothing
    
    strHTML = Replace(strHTML, "{$Display_Homepage}", IsDisplay(FoundInArr(RegFields_MustFill, "Homepage", ",")))
    strHTML = Replace(strHTML, "{$Display_QQ}", IsDisplay(FoundInArr(RegFields_MustFill, "QQ", ",")))
    strHTML = Replace(strHTML, "{$Display_ICQ}", IsDisplay(FoundInArr(RegFields_MustFill, "ICQ", ",")))
    strHTML = Replace(strHTML, "{$Display_MSN}", IsDisplay(FoundInArr(RegFields_MustFill, "MSN", ",")))
    strHTML = Replace(strHTML, "{$Display_Yahoo}", IsDisplay(FoundInArr(RegFields_MustFill, "Yahoo", ",")))
    strHTML = Replace(strHTML, "{$Display_UC}", IsDisplay(FoundInArr(RegFields_MustFill, "UC", ",")))
    strHTML = Replace(strHTML, "{$Display_Aim}", IsDisplay(FoundInArr(RegFields_MustFill, "Aim", ",")))
    strHTML = Replace(strHTML, "{$Display_OfficePhone}", IsDisplay(FoundInArr(RegFields_MustFill, "OfficePhone", ",")))
    strHTML = Replace(strHTML, "{$Display_HomePhone}", IsDisplay(FoundInArr(RegFields_MustFill, "HomePhone", ",")))
    strHTML = Replace(strHTML, "{$Display_Fax}", IsDisplay(FoundInArr(RegFields_MustFill, "Fax", ",")))
    strHTML = Replace(strHTML, "{$Display_Mobile}", IsDisplay(FoundInArr(RegFields_MustFill, "Mobile", ",")))
    strHTML = Replace(strHTML, "{$Display_Region}", IsDisplay(FoundInArr(RegFields_MustFill, "Region", ",")))
    strHTML = Replace(strHTML, "{$Display_Address}", IsDisplay(FoundInArr(RegFields_MustFill, "Address", ",")))
    strHTML = Replace(strHTML, "{$Display_ZipCode}", IsDisplay(FoundInArr(RegFields_MustFill, "ZipCode", ",")))
    strHTML = Replace(strHTML, "{$Display_TrueName}", IsDisplay(FoundInArr(RegFields_MustFill, "TrueName", ",")))
    strHTML = Replace(strHTML, "{$Display_Birthday}", IsDisplay(FoundInArr(RegFields_MustFill, "Birthday", ",")))
    strHTML = Replace(strHTML, "{$Display_IDCard}", IsDisplay(FoundInArr(RegFields_MustFill, "IDCard", ",")))
    strHTML = Replace(strHTML, "{$Display_Vocation}", IsDisplay(FoundInArr(RegFields_MustFill, "Vocation", ",")))
    strHTML = Replace(strHTML, "{$Display_Company}", IsDisplay(FoundInArr(RegFields_MustFill, "Company", ",")))
    strHTML = Replace(strHTML, "{$Display_Department}", IsDisplay(FoundInArr(RegFields_MustFill, "Department", ",")))
    strHTML = Replace(strHTML, "{$Display_PosTitle}", IsDisplay(FoundInArr(RegFields_MustFill, "PosTitle", ",")))
    strHTML = Replace(strHTML, "{$Display_Marriage}", IsDisplay(FoundInArr(RegFields_MustFill, "Marriage", ",")))
    strHTML = Replace(strHTML, "{$Display_Income}", IsDisplay(FoundInArr(RegFields_MustFill, "Income", ",")))
    strHTML = Replace(strHTML, "{$Display_UserFace}", IsDisplay(FoundInArr(RegFields_MustFill, "UserFace", ",")))
    strHTML = Replace(strHTML, "{$Display_FaceWidth}", IsDisplay(FoundInArr(RegFields_MustFill, "FaceWidth", ",")))
    strHTML = Replace(strHTML, "{$Display_FaceHeight}", IsDisplay(FoundInArr(RegFields_MustFill, "FaceHeight", ",")))
    strHTML = Replace(strHTML, "{$Display_Sign}", IsDisplay(FoundInArr(RegFields_MustFill, "Sign", ",")))
    strHTML = Replace(strHTML, "{$Display_SpareEmail}", IsDisplay(False))
    strHTML = Replace(strHTML, "{$Display_Privacy}", IsDisplay(FoundInArr(RegFields_MustFill, "Privacy", ",")))
    strHTML = Replace(strHTML, "{$Display_CheckCode}", IsDisplay(EnableCheckCodeOfReg))
    strHTML = Replace(strHTML, "{$Display_QAofReg}", IsDisplay(EnableQAofReg))
    strHTML = Replace(strHTML, "{$QAofReg}", GetQAofReg(QAofReg))
    
    Response.Write strHTML
End Sub

Function IsDisplay(Display)
    If Display = True Then
        IsDisplay = ""
    Else
        IsDisplay = " Style='display:none'"
    End If
End Function
Function GetQAofReg(QAofReg)
    Dim arrQAofReg, i, strTemp
    arrQAofReg = Split(QAofReg & "", "$$$")
    For i = 0 To 2
        If Trim(arrQAofReg(i * 2)) <> "" Then
            strTemp = strTemp & Cstr(i + 1) & "、" & arrQAofReg(i * 2) & "<br><input type='text' name='RegAnswer" & i & "' size='30'><br><br>"
        End If
    Next
    GetQAofReg = strTemp
End Function
%>
