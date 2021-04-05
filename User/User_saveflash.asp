<!--#include file="../conn.asp"-->
<!--#include file="../inc/function.asp"-->
<%
'目前是保存为172 * 130格式的BMP图象

Dim act, ObjInstalled_FSO, color_name, Create_1, imgurl, SaveFileName, dirMonth
objName_FSO = Application("objName_FSO")
ObjInstalled_FSO = IsObjInstalled(objName_FSO)
If ObjInstalled_FSO = True Then
    Set fso = Server.CreateObject(objName_FSO)
Else
    Response.Write "&&SendFlag=保存 >>> NO"
    Response.end
End If

act = trim(request("act"))
If act="" Then
    Call Main()
Else
    Call CoverColorFile()
End If
set fso = Nothing


Sub Main()
    If CheckUserLogined() = False Then
        Call CloseConn
        Set fso = Nothing
        Response.Write "&&SendFlag=保存 >>> NO"
        exit sub
    End If
    if len(trim(request("rgb_color")))<1000 then
        Response.Write "&&SendFlag=保存 >>> NO"
    else
        color_name = "flashimg/" & Year(now()) & Month(now()) & Day(now()) & Hour(now()) & Minute(now()) & Session.SessionID & ".c"
        Set Create_1 = fso.CreateTextFile(server.MapPath(color_name))
        Create_1.write(trim(request("rgb_color")))
        Create_1.close

        dim urlwords, wordnum, i, weburl
        urlwords = split(trim(request.ServerVariables("SCRIPT_NAME")),"/")
        wordnum = UBound(urlwords)
        for i=1 to wordnum-1
        	weburl = weburl & "/" & urlwords(i)
        next

        imgurl = "http://" & request.ServerVariables("SERVER_NAME") & weburl & "/"

        imgurl = imgurl & "User_saveflash.asp?act=2&color_url=" & color_name '图片远程地址。

        SaveFileName = InstallDir & "Space/" & UserName & "/" &  Year(Now()) & Right("0" & Month(Now()), 2) & "/"
        If fso.FolderExists(Server.MapPath(SaveFileName)) = False Then fso.CreateFolder Server.MapPath(SaveFileName)

        SaveFileName = SaveFileName & Year(now()) & Month(now()) & Day(now()) & Hour(now()) & Minute(now()) & Session.SessionID & ".bmp"
        call SaveImg(SaveFileName,imgurl)
    end if
End Sub

sub SaveImg(FileName,strUrl)
    dim curlpath, Retrieval
    Set Retrieval = Server.CreateObject("MSXML2.ServerXMLHTTP")
    Retrieval.Open "Get", strUrl, False, "", ""
    Retrieval.Send
    If Retrieval.ReadyState = 4 Then
        set ads=server.CreateObject("Adodb.Stream")
        ads.Type=1
        ads.Mode = 3
        ads.Open
        ads.Write Retrieval.ResponseBody
        ads.SaveToFile server.MapPath(FileName),2
        ads.Close()
        set ads=nothing
    End If
    Set Retrieval = Nothing

    Response.Write "&&SendFlag=" & FileName
end sub

Sub CoverColorFile()
    Dim whichfile, head, Colortxt, i, rline, badwords
    Response.Expires = -9999 
    Response.AddHeader "Pragma","no-cache"
    Response.AddHeader "cache-ctrol","no-cache"
    Response.ContentType = "Image/bmp"

    '输出图像文件头   
    head = ChrB(66) & ChrB(77) & ChrB(118) & ChrB(250) & ChrB(1) & ChrB(0) & ChrB(0) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(54) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(40) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(172) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(130) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(1) & ChrB(0) & ChrB(24) & ChrB(0) & ChrB(0) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(64) & ChrB(250) & ChrB(1) & ChrB(0) & ChrB(0) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) &_
    ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0)
  
    Response.BinaryWrite head

    whichfile=trim(request("color_url"))

    Set Colortxt = fso.OpenTextFile(server.mappath(whichfile),1)
    rline = Colortxt.ReadLine
    badwords = split(rline,"|")
    Colortxt.Close

    fso.deleteFile(server.mappath(whichfile))
 
    for i=0 to UBound(badwords)
        Response.BinaryWrite to3(badwords(i))
    next
End Sub

Function chn10(nums)
    Dim tmp,tmpstr,i
    nums_len=Len(nums)
    For i=1 To nums_len
        tmp=Mid(nums,i,1)
        If IsNumeric(tmp) Then
            tmp=tmp * 16 * (16^(nums_len-i-1))
        Else
            tmp=(ASC(UCase(tmp))-55) * (16^(nums_len-i))
        End If
        tmpstr=tmpstr+tmp
    Next
    chn10 = tmpstr
End Function

Function to3(nums)
    Dim tmp,i
    Dim myArray()
    For i=1 To 3
        tmp=Mid(nums,i*2-1,2)
        Redim Preserve myArray(i)
        myArray(i) = chn10(tmp)
    Next
    to3 = ChrB(myArray(3))&ChrB(myArray(2))&ChrB(myArray(1))
End Function
%>