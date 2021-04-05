<%@language=vbscript codepage=936 %>
<%
Option Explicit
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
%>
<!-- #include File="../conn.asp" -->
<%

Dim Action, Title, ModuleType, ChannelShortName, ChannelShowType, imageproperty, rs
Dim editLabel, Labletemp
Dim ClassID, NClassID, IncludeChild, SpecialID, Num, ProductType, IsHot, IsElite, AuthorName, DateNum
Dim OrderType, ShowType, TitleLen, ContentLen, ShowClassName, ShowPropertyType, ShowIncludePic, ShowAuthor
Dim ShowDateType, ShowHits, ShowHotSign, ShowNewSign, ShowTips, ShowCommentLink, UsePage, OpenType, Cols
Dim ImgWidth, ImgHeight, iTimeOut, urltype, CssNameA, CssName1, CssName2, effectID, IntervalLines
'商城
Dim ShowTableTitle, TableTitleStr, ShowProductModel, ShowProductStandard, ShowUnit, ShowStocksType, ShowPriceType
Dim ShowWeight, ShowPrice_Market, ShowPrice_Original, ShowPrice, ShowPrice_Member, ShowDiscount, ShowButtonType, ButtonStyle
Dim CssNameTable, CssNameTitle
'人才招聘
Dim PositionNum, IsUrgent, WorkPlaceNameLen, SubCompanyNameLen, PShowPoints, WShowPoints, SShowPoints, ShowPositionID, ShowPositionName, ShowWorkPlaceName, ShowSubCompanyName, ShowPositionNum, ShowPositionStatus, ShowValidDate, ShowUrgentSign, ShowNum

'是模板还是右键
Dim InsertTemplate
Dim ChannelID, iChannelID, dChannelID

ChannelID = Trim(request("ChannelID"))
dChannelID = Trim(request("dChannelID"))

NClassID = False

If dChannelID = "" Then
   dChannelID = ChannelID
End If
If ChannelID = "" And iChannelID = "" Then
    Response.Write "频道参数丢失！"
    Response.End
End If

If ChannelID = "ChannelID" Then
    iChannelID = Trim(dChannelID)
Else
    ChannelID = PE_CLng(ChannelID)
    iChannelID = ChannelID
End If

Action = Trim(request.querystring("Action"))
Title = Trim(request.querystring("Title"))
ModuleType = PE_CLng(Trim(request.querystring("ModuleType")))
ChannelShowType = Trim(request.querystring("ChannelShowType"))
InsertTemplate = PE_CLng(Trim(request.querystring("InsertTemplate")))

If SpecialID = "" Then SpecialID = 0
If Trim(request.querystring("editLabel")) <> "" Then
    editLabel = True
End If

Call Execute
Call strJS

   
Sub Execute()

    If Action = "" Then
        Action = Trim(request.Form("Action"))
        Title = Trim(request.Form("Title"))
        ModuleType = PE_CLng(Trim(request.Form("ModuleType")))
        If Trim(request.Form("ChannelID")) = "ChannelID" Then
            ChannelID = Trim(request.Form("ChannelID"))
        Else
            ChannelID = PE_CLng(Trim(request.Form("ChannelID")))
        End If
        ChannelShowType = Trim(request.Form("ChannelShowType"))
        InsertTemplate = PE_CLng(Trim(request.Form("InsertTemplate")))
        If Trim(request.Form("SpecialID")) = "SpecialID" Then
            SpecialID = Trim(request.Form("SpecialID"))
        Else
            SpecialID = PE_CLng(Trim(request.Form("SpecialID")))
        End If
        If editLabel = "" Then
            editLabel = Trim(request.Form("editLabel"))
        End If
    End If
    If Trim(request.querystring("editLabel")) = "" Then
        
        If ModuleType = 1 Then
            ChannelShortName = "文章"
            imageproperty = "article"
        ElseIf ModuleType = 2 Then
            ChannelShortName = "软件"
            imageproperty = "Soft"
        ElseIf ModuleType = 3 Then
            ChannelShortName = "图片"
            imageproperty = "Photo"
        ElseIf ModuleType = 5 Then
            iChannelID = 1000
            ChannelShortName = "商品"
            imageproperty = "Product"
        ElseIf ModuleType = 8 Then
            ChannelShortName = "职位"
            imageproperty = "Job"
        End If
    Else
        Call Modifylabel

        If ChannelID = "ChannelID" Then
            iChannelID = PE_CLng(Trim(dChannelID))
        Else
            ChannelID = PE_CLng(ChannelID)
            iChannelID = ChannelID
        End If
    End If

    Response.Write "<html><head><title>" & Title & "</title>" & vbCrLf
    Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
    Response.Write "<link href='" & InstallDir & AdminDir & "/Admin_Style.css' rel='stylesheet' type='text/css'>" & vbCrLf
    Response.Write "<base target='_self'>"
    Response.Write "</head>" & vbCrLf
    Response.Write "<body leftmargin=0 topmargin=0>" & vbCrLf
    Response.Write "<form action='editor_label.asp' method='post' name='myform' id='myform'>"
    Response.Write "  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
    Response.Write "    <tr class='title'>"
    Response.Write "      <td height='22' colspan='2' align='center'><strong>" & Title & "</strong></td>"
    Response.Write "    </tr>"
    If ModuleType <> 8 Then
        If ModuleType <> 5 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td height='25' align='right' class='tdbg5'><strong>所属频道：</strong></td>" & vbCrLf
            Response.Write "      <td height='25'><input type='hidden' name='iChannelID' value='" & ChannelID & "'><select name='ChannelID' onChange='document.myform.submit();'>" & GetChannel_Option(ModuleType, ChannelID) & "</select></td>"
            Response.Write "    </tr>"
        End If
        If PE_CLng(iChannelID) > 0 Or ModuleType = 5 Then
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td height='25' align='right' class='tdbg5'><strong>所属栏目：</strong></td>" & vbCrLf
            Response.Write "      <td height='25'><select name='ClassID' "
            If NClassID = True Then
                Response.Write "size='2' multiple style='height:250px;width:400px;'"
            Else
                Response.Write "size='1'"
            End If
            Response.Write ">" & GetClass_Channel(iChannelID, Trim(ClassID), NClassID) & "</select>"
            Response.Write " <input type='checkbox' name='IncludeChild' value='1' "
            If LCase(Trim(IncludeChild)) = "true" Then
                Response.Write " checked "
            End If
            Response.Write " >包含子栏目&nbsp;&nbsp;<font color='red'><b>注意：</b></font>不能指定为外部栏目 </font>"
            Response.Write "  <br><input type='checkbox' name='NClassChild' value='1' onClick=""javascript:NClassIDChild()"" "
            If NClassID = True Then
                Response.Write " checked "
            End If
            Response.Write " >是否选择多个栏目&nbsp;&nbsp;<font color='red'><b>注意：</b></font>多选红色的栏目不能选 </font>"
            Response.Write "      </td>"
            Response.Write "    </tr>"
            Response.Write "    <tr class='tdbg'>"
            Response.Write "      <td height='25' align='right' class='tdbg5'><strong>所属专题：</strong></td>"
            Response.Write "      <td height='25' ><select name='SpecialID' id='SpecialID'>" & GetSpecial_Option(iChannelID, SpecialID) & "</select></td>"
            Response.Write "    </tr>"
        Else
            Response.Write "<INPUT TYPE='hidden' name='ClassID' value='0' >"
            Response.Write "<INPUT TYPE='hidden' name='NClassChild' value='0' >"
            Response.Write "<INPUT TYPE='hidden' name='IncludeChild' value='true' >"
            Response.Write "<INPUT TYPE='hidden' name='SpecialID' value='0' >"
        End If

        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>标签说明：</strong></td>" & vbCrLf
        Response.Write "      <td height='25'><INPUT TYPE='text' NAME='lableExplain' value='' id='id' size='15' maxlength='20'>&nbsp;&nbsp;<FONT style='font-size:12px' color='blue'>请在这里填写标签的使用说明方便以后的查找</FONT> </td>"
        Response.Write "    </tr>"
    End If
    Select Case ChannelShowType
     
    Case "GetList"
        Call GetList
    Case "GetPic"
        Call GetPic
    Case "GetSlide"
        Call GetSlide
    Case "GetPositionList"
        Call GetPositionList
    Case "GetSearchResult"
        Call GetSearchResult
    Case Else
        Response.Write "错误的参数命令！"
        Response.End
    End Select

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='40' colspan='2' align='center'>"
    Response.Write "        <input name='Title' type='hidden' id='Title' value='" & Title & "'>"
    Response.Write "        <input name='Action' type='hidden' id='Action' value='" & Action & "'>"
    Response.Write "        <input name='editLabel' type='hidden' id='editLabel' value='" & editLabel & "'>"
    Response.Write "        <input name='dChannelID' type='hidden' id='dChannelID' value='" & dChannelID & "'>"
    Response.Write "        <input name='ModuleType' type='hidden' id='ModuleType' value='" & ModuleType & "'>"
    Response.Write "        <input name='InsertTemplate' type='hidden' id='InsertTemplate' value='" & InsertTemplate & "'>"
    Response.Write "        <input name='ChannelShowType' type='hidden' id='ChannelShowType' value='" & ChannelShowType & "'>"
    Response.Write "        <input name='MakeJS' type='button' id='MakeJS' onclick=""makejs('" & Action & "','" & ChannelShowType & "');"" value=' 确 定 '>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "  </table>"
    Response.Write "</form>"
    Response.Write "</body>"
    Response.Write "</html>"

End Sub



Sub GetList()

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>显示样式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='ShowType' id='ShowType'>"
    Response.Write "           <option value='1' "
    If Trim(ShowType) = "1" Then Response.Write "selected"
    Response.Write ">普通列表</option>"
    Response.Write "           <option value='2' "
    If Trim(ShowType) = "2" Then Response.Write "selected"
    Response.Write ">表格式</option>"
    Response.Write "           <option value='3' "
    If Trim(ShowType) = "3" Then Response.Write "selected"
    Response.Write ">各项独立式</option>"
    If ModuleType = 1 Then
        Response.Write "           <option value='4' "
        If Trim(ShowType) = "4" Then Response.Write "selected"
        Response.Write ">智能多列式</option>"
        Response.Write "           <option value='5' "
        If Trim(ShowType) = "5" Then Response.Write "selected"
        Response.Write ">输出DIV格式</option>"
        Response.Write "           <option value='6' "
        If Trim(ShowType) = "6" Then Response.Write "selected"
        Response.Write ">输出RSS格式</option>"
    Else
        Response.Write "           <option value='4' "
        If Trim(ShowType) = "4" Then Response.Write "selected"
        Response.Write ">输出DIV格式</option>"
        Response.Write "           <option value='5' "
        If Trim(ShowType) = "5" Then Response.Write "selected"
        Response.Write ">输出RSS格式</option>"
    End If
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "数目：</strong></td>"
    Response.Write "      <td height='25'><input name='Num' type='text' value='"
    If Trim(Num) = "" Then
        Response.Write "10"
    Else
        Response.Write Num
    End If
    Response.Write "' size='5' maxlength='3'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果为0，将显示所有" & ChannelShortName & "。</font></td>"
    Response.Write "    </tr>"
    If ModuleType = 5 Then
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong> 产品类型：</strong></td>"
    Response.Write "      <td height='25'><select name='ProductType' id='ProductType'>"
    Response.Write "        <option value='1'"
    If Trim(ProductType) = "1" Then Response.Write "selected"
    Response.Write ">正常销售商品</option>"
    Response.Write "        <option value='2'"
    If Trim(ProductType) = "2" Then Response.Write "selected"
    Response.Write ">涨价商品</option>"
    Response.Write "        <option value='3'"
    If Trim(ProductType) = "3" Then Response.Write "selected"
    Response.Write ">降价商品</option>"
    Response.Write "        <option value='4'"
    If Trim(ProductType) = "4" Then Response.Write "selected"
    Response.Write ">促销礼品</option>"
    Response.Write "        <option value='5'"
    If Trim(ProductType) = "5" Then Response.Write "selected"
    Response.Write ">正常销售和涨价商品</option>"
    Response.Write "        <option value='0'"
    If Trim(ProductType) = "0" Then Response.Write "selected"
    Response.Write ">所有商品</option>"
    Response.Write "        </select> </td>"
    Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "属性：</strong></td>"
    Response.Write "      <td height='25'><input name='IsHot' type='checkbox' id='IsHot' value='1'"
    If LCase(Trim(IsHot)) = "true" Then Response.Write "checked"
    Response.Write ">"
    Response.Write "        热门" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<input name='IsElite' type='checkbox' id='IsElite' value='1'"
    If LCase(Trim(IsElite)) = "true" Then Response.Write "checked"
    Response.Write ">"
    Response.Write "        推荐" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果都不选，将显示所有" & ChannelShortName & "。</font></td>"
    Response.Write "    </tr>"
    If ModuleType <> 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>作者姓名：</strong></td>"
        Response.Write "      <td height='25'><input name='AuthorName' type='text' value='"
        If Trim(AuthorName) = """" Then
            Response.Write ""
        Else
            Response.Write AuthorName
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果不为空，则只显示指定录入者的" & ChannelShortName & "，用于个人文集。</font></td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "属性图片：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "       <table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' valign='top'>"
    Response.Write "        <tr>"
    Response.Write "          <td width='100'>"
    Response.Write "            <select name='ShowPropertyType' id='ShowPropertyType' onChange=""javascript:change_item(this)"">"
    Response.Write "           <option value='0' "
    If Trim(ShowPropertyType) = "0" Then Response.Write "selected"
    Response.Write ">不显示</option>"
    Response.Write "           <option value='2' "
    If Trim(ShowPropertyType) = "2" Then Response.Write "selected"
    Response.Write ">符号</option>"
    Response.Write "           <option value='1' "
    If Trim(ShowPropertyType) = "1" Then Response.Write "selected"
    Response.Write ">小图片（样式 1）</option>"
    Response.Write "           <option value='3' "
    If Trim(ShowPropertyType) = "3" Then Response.Write "selected"
    Response.Write ">小图片（样式 2）</option>"
    Response.Write "           <option value='4' "
    If Trim(ShowPropertyType) = "4" Then Response.Write "selected"
    Response.Write ">小图片（样式 3）</option>"
    Response.Write "           <option value='5' "
    If Trim(ShowPropertyType) = "5" Then Response.Write "selected"
    Response.Write ">小图片（样式 4）</option>"
    Response.Write "           <option value='6' "
    If Trim(ShowPropertyType) = "6" Then Response.Write "selected"
    Response.Write ">小图片（样式 5）</option>"
    If ModuleType = 1 Then
        Response.Write "           <option value='7' "
        If Trim(ShowPropertyType) = "7" Then Response.Write "selected"
        Response.Write ">小图片（样式 6）</option>"
        Response.Write "           <option value='8' "
        If Trim(ShowPropertyType) = "8" Then Response.Write "selected"
        Response.Write ">小图片（样式 7）</option>"
        Response.Write "           <option value='9' "
        If Trim(ShowPropertyType) = "9" Then Response.Write "selected"
        Response.Write ">小图片（样式 8）</option>"
        Response.Write "           <option value='10' "
        If Trim(ShowPropertyType) = "10" Then Response.Write "selected"
        Response.Write ">小图片（样式 9）</option>"
    End If
    Response.Write "        </select>"
    Response.Write "         </td>"
    Response.Write "          <td id=objFiles style='display:none'>"
    Response.Write "&nbsp;&nbsp;普通图片&nbsp;&nbsp;<IMG id=common SRC='" & InstallDir & "images/" & imageproperty & "_common.gif' BORDER='0' ALT='普通图片'>&nbsp;&nbsp;推荐图片&nbsp;&nbsp;<IMG SRC='" & InstallDir & "images/" & imageproperty & "_elite.gif' id=elite BORDER='0' ALT='推荐图片'>&nbsp;&nbsp;固定图片&nbsp;&nbsp;<IMG SRC='" & InstallDir & "images/" & imageproperty & "_ontop.gif' id=ontop BORDER='0' ALT='固定图片'>"
    Response.Write "          </td>"
    Response.Write "        </tr>"
    Response.Write "       </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>日期范围：</strong></td>"
    Response.Write "      <td height='25'>只显示最近"
    Response.Write "        <input name='DateNum' type='text' id='DateNum' value="
    If Trim(DateNum) = "" Then
        Response.Write "30"
    Else
        Response.Write DateNum
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        天内更新的" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;如果为空或0，则显示所有天数的" & ChannelShortName & "。</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>排序方法：</strong></td>"
    Response.Write "      <td height='25'><select name='OrderType' id='OrderType'>"
    Response.Write "       <option value='1' "
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID（降序）</option>"
    Response.Write "       <option value='2' "
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID（升序）</option>"
    Response.Write "       <option value='3' "
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">更新时间（降序）</option>"
    Response.Write "       <option value='4' "
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">更新时间（升序）</option>"
    Response.Write "       <option value='5' "
    If Trim(OrderType) = "5" Then Response.Write "selected"
    Response.Write ">点击次数（降序）</option>"
    Response.Write "       <option value='6' "
    If Trim(OrderType) = "6" Then Response.Write "selected"
    Response.Write ">点击次数（升序）</option>"
    Response.Write "       <option value='7' "
    If Trim(OrderType) = "7" Then Response.Write "selected"
    Response.Write ">按评论数（降序）</option>"
    Response.Write "       <option value='8' "
    If Trim(OrderType) = "8" Then Response.Write "selected"
    Response.Write ">按评论数（升序）</option>"
    Response.Write "      </select></td>"
    Response.Write "    </tr>"
    Response.Write " <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>标题最多字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value="
    If Trim(TitleLen) = "" Then
        Response.Write "30"
    Else
        Response.Write TitleLen
    End If
    Response.Write "  size='5' maxlength='3'>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果为0，则显示完整标题。字母算一个字符，汉字算两个字符。</font></td>"
    Response.Write "    </tr>"

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "内容字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value="
    If Trim(ContentLen) = "" Then
        Response.Write "0"
    Else
        Response.Write ContentLen
    End If
    Response.Write "  size='5' maxlength='3'>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果大于0，则在标题下方面显示指定字数的" & ChannelShortName & "内容</font></td>"
    Response.Write "    </tr>"
    'If ModuleType = 1 Or ModuleType = 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>每行的列数：</strong></td>"
        Response.Write "      <td height='25'><INPUT TYPE='text' NAME='Cols' value="
        If Trim(Cols) = "" Then
            Response.Write "1"
        Else
            Response.Write Cols
        End If
        Response.Write "  id='id' size='5' maxlength='3'> &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>超过此列数就换行</font>"
        Response.Write "      <input type='hidden' name='urltype' value='0'></td>"
        Response.Write "    </tr>"
    'End If
    Response.Write " <tr class='tdbg'>"
    Response.Write "      <td height='50' align='right' class='tdbg5'><strong>显示内容：</strong></td>"
    Response.Write "      <td height='50'><table width='100%' border='0' cellpadding='1' cellspacing='2'>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowClassName' type='checkbox' id='ShowClassName' value='1' "
    If LCase(Trim(ShowClassName)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">所属栏目</td>"
    If ModuleType = 1 Then
        Response.Write "          <td><input name='ShowIncludePic' type='checkbox' id='ShowIncludePic' value='1' "
        If LCase(Trim(ShowIncludePic)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">“图文”标志</td>"
    End If
    If ModuleType <> 5 Then
        Response.Write "          <td><input name='ShowAuthor' type='checkbox' id='ShowAuthor' value='1' "
        If LCase(Trim(ShowAuthor)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">作者</td>"
    End If
    Response.Write "          <td>更新时间"
    Response.Write "              <select name='ShowDateType' id='ShowDateType'>"
    Response.Write "                <option value='0' "
    If Trim(ShowDateType) = "0" Then Response.Write "selected"
    Response.Write ">不显示</option>"
    Response.Write "                <option value='1' "
    If Trim(ShowDateType) = "1" Then Response.Write "selected"
    Response.Write ">年月日</option>"
    Response.Write "                <option value='2'"
    If Trim(ShowDateType) = "2" Then Response.Write "selected"
    Response.Write ">月日</option>"
    Response.Write "                <option value='3' "
    If Trim(ShowDateType) = "3" Then Response.Write "selected"
    Response.Write ">月-日</option>"
    Response.Write "              </select>"
    Response.Write "          </td>"
    If ModuleType <> 5 Then
        Response.Write "          <td><input name='ShowHits' type='checkbox' id='ShowHits' value='1' "
        If LCase(Trim(ShowHits)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write " >点击次数</td>"
    End If
    Response.Write "        </tr>"
    Response.Write "        <tr>"
    Response.Write "          <td><input name='ShowHotSign' type='checkbox' id='ShowHotSign' value='1' "
    If LCase(Trim(ShowHotSign)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">热点" & ChannelShortName & "标志</td>"
    Response.Write "          <td><input name='ShowNewSign' type='checkbox' id='ShowNewSign' value='1' "
    If LCase(Trim(ShowNewSign)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">最新" & ChannelShortName & "标志</td>"
    If ModuleType <> 5 Then
        Response.Write "          <td><input name='ShowTips' type='checkbox' id='ShowTips' value='1' "
        If LCase(Trim(ShowTips)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">显示提示信息</td>"
    End If
    If ModuleType = 1 Then
        Response.Write "          <td><input name='ShowCommentLink' type='checkbox' id='ShowCommentLink' value='1' "
        If LCase(Trim(ShowCommentLink)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">显示评论链接</td>"
    End If
    Response.Write "          <td><input name='UsePage' type='checkbox' id='UsePage' value='1'"
    If LCase(Trim(UsePage)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">是否分页显示</td>"
    Response.Write "        </tr>"
    Response.Write "      </table>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "打开方式：</strong></td>"
    Response.Write "      <td height='25'>"
    Response.Write "        <select name='OpenType' id='OpenType'>"
    Response.Write "          <option value='0' "
    If Trim(OpenType) = "0" Then Response.Write "selected"
    Response.Write ">在原窗口打开</option>"
    Response.Write "          <option value='1' "
    If Trim(OpenType) = "1" Then Response.Write "selected"
    Response.Write ">在新窗口打开</option>"
    Response.Write "        </select>"
    Response.Write "      </td>"
    Response.Write "    </tr>"
    If ModuleType = 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>每隔多少行空白一行：</strong></td>"
        Response.Write "      <td height='25'><input name='IntervalLines' type='text' value='"
        If Trim(IntervalLines) = """" Then
            Response.Write ""
        Else
            Response.Write IntervalLines
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;<font color=blue>为0时不空行</font></td>"
        Response.Write "    </tr>"
        Response.Write "     <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>表格头部文字：</strong></td>"
        Response.Write "      <td height='25' >"
        If Trim(TableTitleStr) = "" Or InStr(TableTitleStr, "|") <= 0 Or UBound(Split(TableTitleStr, "|")) > 12 Or UBound(Split(TableTitleStr, "|")) < 12 Then
            TableTitleStr = "商品名称|型号|规格|上市时间|单位|库存量|重量|市场价|商城价|优惠价|会员价|折扣率|操作"
        End If
        TableTitleStr = Split(TableTitleStr, "|")
        Response.Write "<table border='0' cellpadding='0' cellspacing='0' width='100%' height='100%' align='center'>"
        Response.Write " <tr class='tdbg' align='center'>"
        Response.Write "    <td>商品名称</td><td>型号</td><td>规格</td><td>上市时间</td><td>单位</td><td>库存量</td><td>重量</td>"
        Response.Write " </tr>"
        Response.Write " <tr class='tdbg' align='center'>"
        Response.Write "    <td><input name='TableTitleStr1' type='text' value='" & TableTitleStr(0) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr2' type='text' value='" & TableTitleStr(1) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr3' type='text' value='" & TableTitleStr(2) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr4' type='text' value='" & TableTitleStr(3) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr5' type='text' value='" & TableTitleStr(4) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr6' type='text' value='" & TableTitleStr(5) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr7' type='text' value='" & TableTitleStr(6) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write " </tr>"
        Response.Write "  <tr class='tdbg' align='center'>"
        Response.Write "    <td>市场价</td><td>商城价</td><td>优惠价</td><td>会员价</td><td>折扣率</td><td>操作</td>"
        Response.Write " </tr>"
        Response.Write "  <tr class='tdbg' align='center'>"
        Response.Write "    <td><input name='TableTitleStr8' type='text' value='" & TableTitleStr(7) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr9' type='text' value='" & TableTitleStr(8) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr10' type='text' value='" & TableTitleStr(9) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr11' type='text' value='" & TableTitleStr(10) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr12' type='text' value='" & TableTitleStr(11) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "    <td><input name='TableTitleStr13' type='text' value='" & TableTitleStr(12) & "'  size='10' maxlength='20' style='text-align: center;'></td>"
        Response.Write "  </tr>"
        Response.Write " </table>"
        Response.Write "     </td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>显示商品库存方式：</strong></td>"
        Response.Write "      <td height='25'><select name='ShowStocksType' id='ShowStocksType'>"
        Response.Write "       <option value='0' "
        If Trim(ShowStocksType) = "0" Then Response.Write "selected"
        Response.Write ">不显示</option>"
        Response.Write "       <option value='1' "
        If Trim(ShowStocksType) = "1" Then Response.Write "selected"
        Response.Write ">显示虚拟库存</option>"
        Response.Write "       <option value='2' "
        If Trim(ShowStocksType) = "2" Then Response.Write "selected"
        Response.Write ">显示实际库存</option>"
        Response.Write "      </select></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>按钮显示方式：</strong></td>"
        Response.Write "      <td height='25'><select name='ShowButtonType' id='ShowButtonType'>"
        Response.Write "       <option value='0' "
        If Trim(ShowButtonType) = "0" Then Response.Write "selected"
        Response.Write ">不显示</option>"
        Response.Write "       <option value='1' "
        If Trim(ShowButtonType) = "1" Then Response.Write "selected"
        Response.Write ">显示购买按钮</option>"
        Response.Write "       <option value='2' "
        If Trim(ShowButtonType) = "2" Then Response.Write "selected"
        Response.Write ">显示详细按钮</option>"
        Response.Write "       <option value='3' "
        If Trim(ShowButtonType) = "3" Then Response.Write "selected"
        Response.Write ">显示收藏按钮</option>"
        Response.Write "       <option value='4' "
        If Trim(ShowButtonType) = "4" Then Response.Write "selected"
        Response.Write ">显示购买＋详细按钮</option>"
        Response.Write "       <option value='5' "
        If Trim(ShowButtonType) = "5" Then Response.Write "selected"
        Response.Write ">显示购买＋收藏按钮</option>"
        Response.Write "       <option value='6' "
        If Trim(ShowButtonType) = "6" Then Response.Write "selected"
        Response.Write ">详细＋收藏按钮</option>"
        Response.Write "       <option value='7' "
        If Trim(ShowButtonType) = "7" Then Response.Write "selected"
        Response.Write ">三个都显示</option>"
        Response.Write "      </select></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='50' align='right' class='tdbg5'><strong>显示商品详细信息：</strong></td>"
        Response.Write "      <td height='50'><table width='100%' border='0' cellpadding='1' cellspacing='2'>"
        Response.Write "        <tr>"
        Response.Write "          <td><input name='ShowTableTitle' type='checkbox' id='ShowTableTitle' value='1' "
        If LCase(Trim(ShowTableTitle)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">表格头部文字</td>"
        Response.Write "          <td><input name='ShowProductModel' type='checkbox' id='ShowProductModel' value='1' "
        If LCase(Trim(ShowProductModel)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示商品型号</td>"
        Response.Write "          <td><input name='ShowProductStandard' type='checkbox' id='ShowProductStandard' value='1' "
        If LCase(Trim(ShowProductStandard)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示商品规格</td>"
        Response.Write "        </tr>"
        Response.Write "        <tr>"
        Response.Write "          <td><input name='ShowUnit' type='checkbox' id='ShowUnit' value='1' "
        If LCase(Trim(ShowUnit)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示商品单位</td>"
        Response.Write "          <td><input name='ShowWeight' type='checkbox' id='ShowWeight' value='1' "
        If LCase(Trim(ShowWeight)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示商品重量</td>"
        Response.Write "          <td><input name='ShowPrice_Market' type='checkbox' id='ShowPrice_Market' value='1' "
        If LCase(Trim(ShowPrice_Market)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示市场价</td>"
        Response.Write "        </tr>"
        Response.Write "      <tr>"
        Response.Write "          <td><input name='ShowPrice_Original' type='checkbox' id='ShowPrice_Original' value='1' "
        If LCase(Trim(ShowPrice_Original)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示原价</td>"

        Response.Write "          <td><input name='ShowPrice' type='checkbox' id='ShowPrice' value='1' "
        If LCase(Trim(ShowPrice)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示当前零售价</td>"

        Response.Write "          <td><input name='ShowPrice_Member' type='checkbox' id='ShowPrice_Member' value='1' "
        If LCase(Trim(ShowPrice_Member)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示会员价</td>"

        Response.Write "          <td><input name='ShowDiscount' type='checkbox' id='ShowDiscount' value='1' "
        If LCase(Trim(ShowDiscount)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">是否显示折扣率</td>"
        Response.Write "        </tr>"
        Response.Write "      </table>"
        Response.Write "     </td>"
        Response.Write "   </tr>"
        Response.Write "   <tr class='tdbg'>"
        Response.Write "     <td height='25' align='right' class='tdbg5'><strong>按钮样式：</strong></td>"
        Response.Write "     <td height='25' ><input name='ButtonStyle' type='text' value='"
        If Trim(ButtonStyle) = """" Then
            Response.Write ""
        Else
            Response.Write ButtonStyle
        End If
        Response.Write "'  size='10' maxlength='20'>&nbsp;&nbsp;<font color='blue'>请填写定义图片数字</font><br>"
        Response.Write "举例：<br>"
        Response.Write "　" & InstallDir & "Shop/images/ProductBuy<FONT color='blue'>“数字”</FONT>.gif<br>"
        Response.Write "　" & InstallDir & "Shop/images/ProductContent<FONT color='blue'>“数字”</FONT>.gif<br>"
        Response.Write "　" & InstallDir & "Shop/images/ProductFav<FONT color='blue'>“数字”</FONT>.gif<br>"
        Response.Write "&nbsp;&nbsp;<font color='blue'>请按以上方式制作上传自定义按钮图片</font></td>"
        Response.Write "   </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>表格CSS：</strong></td>"
        Response.Write "      <td height='25'><input name='CssNameTable' type='text' value='"
        If Trim(CssNameTable) = """" Then
            Response.Write ""
        Else
            Response.Write CssNameTable
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>表格的CSS类名，可选参数(仅在表格式有效)</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>表格头部CSS：</strong></td>"
        Response.Write "      <td height='25'><input name='CssNameTitle' type='text' value='"
        If Trim(CssNameTitle) = """" Then
            Response.Write ""
        Else
            Response.Write CssNameTitle
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>表格头部行的CSS类名，可选参数。(仅在表格式有效)</font></td>"
        Response.Write "    </tr>"
    End If
    'If ModuleType = 1 Or ModuleType = 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>CSS类名：</strong></td>"
        Response.Write "      <td height='25'><input name='CssNameA' type='text' value='"
        If Trim(CssNameA) = """" Then
            Response.Write ""
        Else
            Response.Write CssNameA
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>列表中文字链接调用的CSS类名，可选参数(仅在表格式有效)</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>风格样式1：</strong></td>"
        Response.Write "      <td height='25'><input name='CssName1' type='text' value='"
        If Trim(CssName1) = """" Then
            Response.Write ""
        Else
            Response.Write CssName1
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>列表中奇数行的CSS效果的类名，可选参数(仅在表格式有效)</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>风格样式2：</strong></td>"
        Response.Write "      <td height='25'><input name='CssName2' type='text' value='"
        If Trim(CssName2) = """" Then
            Response.Write ""
        Else
            Response.Write CssName2
        End If
        Response.Write "'  size='10' maxlength='10'>&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>列表中偶数行的CSS效果的类名，可选参数(仅在表格式有效)</font></td>"
        Response.Write "    </tr>"
   ' End If
End Sub

Sub GetPic()

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "数目：</strong></td>"
    Response.Write "      <td height='25'><input name='Num' type='text' value="
    If Trim(Num) = "" Then
        Response.Write "4"
    Else
        Response.Write Num
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "      <font color='#FF0000'>如果为0，将显示所有" & ChannelShortName & "。</font></td>"
    Response.Write "    </tr>"
    If ModuleType = 5 Then
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong> 产品类型：</strong></td>"
    Response.Write "      <td height='25'><select name='ProductType' id='ProductType'>"
    Response.Write "        <option value='1'"
    If Trim(ProductType) = "1" Then Response.Write "selected"
    Response.Write ">正常销售商品</option>"
    Response.Write "        <option value='2'"
    If Trim(ProductType) = "2" Then Response.Write "selected"
    Response.Write ">涨价商品</option>"
    Response.Write "        <option value='3'"
    If Trim(ProductType) = "3" Then Response.Write "selected"
    Response.Write ">降价商品</option>"
    Response.Write "        <option value='4'"
    If Trim(ProductType) = "4" Then Response.Write "selected"
    Response.Write ">促销礼品</option>"
    Response.Write "        <option value='5'"
    If Trim(ProductType) = "5" Then Response.Write "selected"
    Response.Write ">正常销售和涨价商品</option>"
    Response.Write "        <option value='0'"
    If Trim(ProductType) = "0" Then Response.Write "selected"
    Response.Write ">所有商品</option>"
    Response.Write "        </select> </td>"
    Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "属性：</strong></td>"
    Response.Write "      <td height='25'> <input name='IsHot' type='checkbox' id='IsHot' value='1' "
    If LCase(Trim(IsHot)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">"
    Response.Write "        热门" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='checkbox' id='IsElite' value='1' "
    If LCase(Trim(IsElite)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">"
    Response.Write "        推荐" & ChannelShortName & " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果都不选，将显示所有" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>日期范围：</strong></td>"
    Response.Write "      <td height='25'>只显示最近"
    Response.Write "        <input name='DateNum' type='text' id='DateNum' value="
    If Trim(DateNum) = "" Then
        Response.Write "30"
    Else
        Response.Write DateNum
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        天内更新的" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;如果为空或0，则显示所有天数的" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>排序方法：</strong></td>"
    Response.Write "      <td height='25'><select name='OrderType' id='OrderType'>"
    Response.Write "       <option value='1' "
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID（降序）</option>"
    Response.Write "       <option value='2' "
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID（升序）</option>"
    Response.Write "       <option value='3' "
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">更新时间（降序）</option>"
    Response.Write "       <option value='4' "
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">更新时间（升序）</option>"
    Response.Write "       <option value='5' "
    If Trim(OrderType) = "5" Then Response.Write "selected"
    Response.Write ">点击次数（降序）</option>"
    Response.Write "       <option value='6' "
    If Trim(OrderType) = "6" Then Response.Write "selected"
    Response.Write ">点击次数（升序）</option>"
    Response.Write "      </select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>显示样式：</strong></td>"
    Response.Write "      <td height='25'><select name='ShowType' id='ShowType'>"
    If ModuleType = 5 Then
        Response.Write "        <option value='1' "
        If Trim(ShowType) = "1" Then Response.Write "selected"
        Response.Write " >图片+名称+价格+按钮：上下排列</option>"
        Response.Write "        <option value='2' "
        If Trim(ShowType) = "2" Then Response.Write "selected"
        Response.Write " >（图片+名称：上下排列）+（名称+价格+按钮）</option>"
        Response.Write "        <option value='3' "
        If Trim(ShowType) = "3" Then Response.Write "selected"
        Response.Write " >图片+（名称+价格+按钮：上下排列）：左右排列</option>"
        Response.Write "        <option value='4' "
        If Trim(ShowType) = "4" Then Response.Write "selected"
        Response.Write " >图片+名称+价格：上下排列</option>"
        Response.Write "        <option value='5' "
        If Trim(ShowType) = "5" Then Response.Write "selected"
        Response.Write " >（图片+名称：上下排列）+价格：左右排列</option>"
        Response.Write "        <option value='6' "
        If Trim(ShowType) = "6" Then Response.Write "selected"
        Response.Write " >图片+（名称+价格：上下排列）：左右排列</option>"
        Response.Write "        <option value='7' "
        If Trim(ShowType) = "7" Then Response.Write "selected"
        Response.Write " >图片+名称+按钮：上下排列</option>"
        Response.Write "        <option value='8' "
        If Trim(ShowType) = "8" Then Response.Write "selected"
        Response.Write " >图片+名称：上下排列</option>"
        Response.Write "        <option value='9' "
        If Trim(ShowType) = "9" Then Response.Write "selected"
        Response.Write " >图片+按钮：上下排列</option>"
        Response.Write "        <option value='10' "
        If Trim(ShowType) = "10" Then Response.Write "selected"
        Response.Write " >只显示图片</option>"
        Response.Write "        <option value='11' "
        If Trim(ShowType) = "11" Then Response.Write "selected"
        Response.Write " >输出DIV格式</option>"
    Else
        Response.Write "        <option value='1' "
        If Trim(ShowType) = "1" Then Response.Write "selected"
        Response.Write " >图片+标题+内容简介：上下排列</option>"
        Response.Write "        <option value='2' "
        If Trim(ShowType) = "2" Then Response.Write "selected"
        Response.Write " >（图片+标题：上下排列）+内容简介：左右排列</option>"
        Response.Write "        <option value='3' "
        If Trim(ShowType) = "3" Then Response.Write "selected"
        Response.Write " >图片+（标题+内容简介：上下排列）：左右排列</option>"
        Response.Write "        <option value='4' "
        If Trim(ShowType) = "4" Then Response.Write "selected"
        Response.Write " >输出DIV格式</option>"
        Response.Write "        <option value='5' "
        If Trim(ShowType) = "5" Then Response.Write "selected"
        Response.Write " >输出RSS格式</option>"
    End If
    Response.Write "        </select>"
    Response.Write "     </td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>首页图片设置：</b></td>"
    Response.Write "      <td height='25'>&nbsp;宽度："
    Response.Write "        <input name='ImgWidth' type='text' id='ImgWidth' value="
    If Trim(ImgWidth) = "" Then
        Response.Write "130"
    Else
        Response.Write ImgWidth
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        像素&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "  高度： <input name='ImgHeight' type='text' id='ImgHeight' value="
    If Trim(ImgHeight) = "" Then
        Response.Write "90"
    Else
        Response.Write ImgHeight
    End If
    Response.Write "  size='5' maxlength='3'>"
    Response.Write "        像素</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>标题最多字符数：</strong></td>"
    Response.Write "      <td height='25'><input name='TitleLen' type='text' id='TitleLen' value="
    If Trim(TitleLen) = "" Then
        Response.Write "30"
    Else
        Response.Write TitleLen
    End If
    Response.Write "   size='5' maxlength='3'>"
    Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>若为0，则不显示标题；若为-1，则显示完整标题。字母算一个字符，汉字算两个字符。</font></td>"
    Response.Write "    </tr>"
    If ModuleType <> 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "内容字符数：</strong></td>"
        Response.Write "      <td height='25'><input name='ContentLen' type='text' id='ContentLen' value="
        If Trim(ContentLen) = "" Then
            Response.Write "0"
        Else
            Response.Write ContentLen
        End If
        Response.Write "  size='5' maxlength='3'>"
        Response.Write "        &nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果大于0，则显示指定字数的" & ChannelShortName & "内容</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>显示内容：</strong></td>"
        Response.Write "      <td height='25'><input name='ShowTips' type='checkbox' id='ShowTips' value='1' "
        If LCase(Trim(ShowTips)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">"
        Response.Write "      显示作者、更新时间、点击数等提示信息</td>"
        Response.Write "    </tr>"
    End If
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>每行显示" & ChannelShortName & "数：</strong></td>"
    Response.Write "      <td height='25'><select name='Cols' id='Cols'>"
    Response.Write "      <option value='1' "
    If Trim(Cols) = "1" Then Response.Write "selected"
    Response.Write ">1</option>"
    Response.Write "      <option value='2' "
    If Trim(Cols) = "2" Then Response.Write "selected"
    Response.Write ">2</option>"
    Response.Write "      <option value='3' "
    If Trim(Cols) = "3" Then Response.Write "selected"
    Response.Write ">3</option>"
    Response.Write "      <option value='4' "
    If Trim(Cols) = "4" Then Response.Write "selected"
    Response.Write ">4</option>"
    Response.Write "      <option value='5' "
    If Trim(Cols) = "5" Then Response.Write "selected"
    Response.Write ">5</option>"
    Response.Write "      <option value='6' "
    If Trim(Cols) = "6" Then Response.Write "selected"
    Response.Write ">6</option>"
    Response.Write "      <option value='7' "
    If Trim(Cols) = "7" Then Response.Write "selected"
    Response.Write ">7</option>"
    Response.Write "      <option value='8' "
    If Trim(Cols) = "8" Then Response.Write "selected"
    Response.Write ">8</option>"
    Response.Write "      <option value='9' "
    If Trim(Cols) = "9" Then Response.Write "selected"
    Response.Write ">9</option>"
    Response.Write "      <option value='10' "
    If Trim(Cols) = "10" Then Response.Write "selected"
    Response.Write ">10</option>"
    Response.Write "      <option value='11' "
    If Trim(Cols) = "11" Then Response.Write "selected"
    Response.Write ">11</option>"
    Response.Write "      <option value='12' "
    If Trim(Cols) = "12" Then Response.Write "selected"
    Response.Write ">12</option>"
    Response.Write "      </select>"
    Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;超过指行列数就会换行</td>"
    Response.Write "    </tr>"
    If ModuleType = 5 Then
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>显示价格方式：</strong></td>"
        Response.Write "      <td height='25'><select name='ShowPriceType' id='ShowPriceType'>"
        Response.Write "      <option value='0' "
        If Trim(ShowPriceType) = "0" Then Response.Write "selected"
        Response.Write ">自动显示</option>"
        Response.Write "      <option value='1' "
        If Trim(ShowPriceType) = "1" Then Response.Write "selected"
        Response.Write ">只显示原价</option>"
        Response.Write "      <option value='2' "
        If Trim(ShowPriceType) = "2" Then Response.Write "selected"
        Response.Write ">只显示当前价</option>"
        Response.Write "      <option value='3' "
        If Trim(ShowPriceType) = "3" Then Response.Write "selected"
        Response.Write ">只显示市场价与原价</option>"
        Response.Write "      <option value='4' "
        If Trim(ShowPriceType) = "4" Then Response.Write "selected"
        Response.Write ">只显示市场价与当前价</option>"
        Response.Write "      <option value='5' "
        If Trim(ShowPriceType) = "5" Then Response.Write "selected"
        Response.Write ">只显示原价与当前价</option>"
        Response.Write "      <option value='6' "
        If Trim(ShowPriceType) = "6" Then Response.Write "selected"
        Response.Write ">只显示原价与会员价</option>"
        Response.Write "      <option value='7' "
        If Trim(ShowPriceType) = "7" Then Response.Write "selected"
        Response.Write ">显示市场价、原价和当前价</option>"
        Response.Write "      <option value='8' "
        If Trim(ShowPriceType) = "8" Then Response.Write "selected"
        Response.Write ">显示市场价、原价和会员价</option>"
        Response.Write "      <option value='9' "
        If Trim(ShowPriceType) = "9" Then Response.Write "selected"
        Response.Write ">显示市场价、当前价和会员价</option>"
        Response.Write "      <option value='10' "
        If Trim(ShowPriceType) = "10" Then Response.Write "selected"
        Response.Write ">显示市场价、原价、当前价和会员价</option>"
        Response.Write "      </select>"
        Response.Write "      &nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>只有当ShowType参数设为含价格方式时才有效</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>是否显示折扣率：</strong></td>"
        Response.Write "          <td><input name='ShowDiscount' type='checkbox' id='ShowDiscount' value='1' "
        If LCase(Trim(ShowDiscount)) = "true" Then
            Response.Write "checked"
        End If
        Response.Write ">&nbsp;&nbsp;&nbsp;&nbsp;<font color='blue'>只有当ShowType参数设为含价格方式时才有效</font></td>"
        Response.Write "    </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>按钮显示方式：</strong></td>"
        Response.Write "      <td height='25'><select name='ShowButtonType' id='ShowButtonType'>"
        Response.Write "       <option value='0' "
        If Trim(ShowButtonType) = "0" Then Response.Write "selected"
        Response.Write ">不显示</option>"
        Response.Write "       <option value='1' "
        If Trim(ShowButtonType) = "1" Then Response.Write "selected"
        Response.Write ">显示购买按钮</option>"
        Response.Write "       <option value='2' "
        If Trim(ShowButtonType) = "2" Then Response.Write "selected"
        Response.Write ">显示详细按钮</option>"
        Response.Write "       <option value='3' "
        If Trim(ShowButtonType) = "3" Then Response.Write "selected"
        Response.Write ">显示收藏按钮</option>"
        Response.Write "       <option value='4' "
        If Trim(ShowButtonType) = "4" Then Response.Write "selected"
        Response.Write ">显示购买＋详细按钮</option>"
        Response.Write "       <option value='5' "
        If Trim(ShowButtonType) = "5" Then Response.Write "selected"
        Response.Write ">显示购买＋收藏按钮</option>"
        Response.Write "       <option value='6' "
        If Trim(ShowButtonType) = "6" Then Response.Write "selected"
        Response.Write ">详细＋收藏按钮</option>"
        Response.Write "       <option value='7' "
        If Trim(ShowButtonType) = "7" Then Response.Write "selected"
        Response.Write ">三个都显示</option>"
        Response.Write "      </select></td>"
        Response.Write "    </tr>"
        Response.Write "   <tr class='tdbg'>"
        Response.Write "     <td height='25' align='right' class='tdbg5'><strong>按钮样式：</strong></td>"
        Response.Write "     <td height='25' ><input name='ButtonStyle' type='text' value='"
        If Trim(ButtonStyle) = """" Then
            Response.Write ""
        Else
            Response.Write ButtonStyle
        End If
        Response.Write "'  size='10' maxlength='20'></td>"
        Response.Write "   </tr>"
        Response.Write "    <tr class='tdbg'>"
        Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "打开方式：</strong></td>"
        Response.Write "      <td height='25'>"
        Response.Write "        <select name='OpenType' id='OpenType'>"
        Response.Write "          <option value='0' "
        If Trim(OpenType) = "0" Then Response.Write "selected"
        Response.Write ">在原窗口打开</option>"
        Response.Write "          <option value='1' "
        If Trim(OpenType) = "1" Then Response.Write "selected"
        Response.Write ">在新窗口打开</option>"
        Response.Write "        </select>"
        Response.Write "      </td>"
        Response.Write "    </tr>"
    End If
End Sub

Sub GetSlide()

    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "数目：</strong></td>"
    Response.Write "      <td height='25'><input name='Num' type='text' value="
    If Trim(Num) = "" Then
        Response.Write "4"
    Else
        Response.Write Num
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "      <font color='#FF0000'>如果为0，将显示所有" & ChannelShortName & "。</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "属性：</strong></td>"
    Response.Write "      <td height='25'> <input name='IsHot' type='checkbox' id='IsHot' value='1' "
    If LCase(Trim(IsHot)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">"
    Response.Write "        热门" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp; <input name='IsElite' type='checkbox' id='IsElite' value='1' "
    If LCase(Trim(IsElite)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">"
    Response.Write "        推荐" & ChannelShortName & " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>如果都不选，将显示所有" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>日期范围：</strong></td>"
    Response.Write "      <td height='25'>只显示最近"
    Response.Write "        <input name='DateNum' type='text' id='DateNum' value="
    If Trim(DateNum) = "" Then
        Response.Write "30"
    Else
        Response.Write DateNum
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        天内更新的" & ChannelShortName & "&nbsp;&nbsp;&nbsp;&nbsp;<font color='#FF0000'>&nbsp;&nbsp;如果为空或0，则显示所有天数的" & ChannelShortName & "</font></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><strong>排序方法：</strong></td>"
    Response.Write "      <td height='25'><select name='OrderType' id='OrderType'>"
    Response.Write "       <option value='1' "
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID（降序）</option>"
    Response.Write "       <option value='2' "
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">" & ChannelShortName & "ID（升序）</option>"
    Response.Write "       <option value='3' "
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">更新时间（降序）</option>"
    Response.Write "       <option value='4' "
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">更新时间（升序）</option>"
    Response.Write "       <option value='5' "
    If Trim(OrderType) = "5" Then Response.Write "selected"
    Response.Write ">点击次数（降序）</option>"
    Response.Write "       <option value='6' "
    If Trim(OrderType) = "6" Then Response.Write "selected"
    Response.Write ">点击次数（升序）</option>"
    Response.Write "      </select></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>首页图片设置：</b></td>"
    Response.Write "      <td height='25'>&nbsp;宽度："
    Response.Write "        <input name='ImgWidth' type='text' id='ImgWidth' value="
    If Trim(ImgWidth) = "" Then
        Response.Write "130"
    Else
        Response.Write ImgWidth
    End If
    Response.Write " size='5' maxlength='3'>"
    Response.Write "        像素&nbsp;&nbsp;&nbsp;&nbsp;"
    Response.Write "  高度： <input name='ImgHeight' type='text' id='ImgHeight' value="
    If Trim(ImgHeight) = "" Then
        Response.Write "90"
    Else
        Response.Write ImgHeight
    End If
    Response.Write "  size='5' maxlength='3'>"
    Response.Write "        像素</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>标题/名称长度</b></td>"
    Response.Write "      <td height='25'> <input name='TitleLen' type='text' id='TitleLen' value="
    If Trim(TitleLen) = "" Then
        Response.Write "30"
    Else
        Response.Write TitleLen
    End If
    Response.Write "  size='5' maxlength='3'> 个字符</td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>效果变换间隔时间</b></td>"
    
    Response.Write "      <td height='25'> <input name='iTimeOut' type='text' id='iTimeOut' value="
    If Trim(iTimeOut) = "" Then
        Response.Write "5000"
    Else
        Response.Write iTimeOut
    End If
    Response.Write "  size='5' maxlength='5'>&nbsp;&nbsp;<font color=blue><b>毫秒为单位</b></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class='tdbg'>"
    Response.Write "      <td height='25' align='right' class='tdbg5'><b>图片转换效果</b></td>"
    Response.Write "      <td height='25'> <input name='effectID' type='text' id='effectID' value="
    If Trim(effectID) = "" Then
        Response.Write "-1"
    Else
        Response.Write effectID
    End If
    Response.Write "  size='5' maxlength='3'>&nbsp;&nbsp;<font color=blue><b>-1表示随机效果，0至23指定某一种特效</b></td>"
    Response.Write "    </tr>"
    'Response.Write "    <tr class='tdbg'>"
    'Response.Write "      <td height='25' align='right' class='tdbg5'><strong>" & ChannelShortName & "打开方式：</strong></td>"
    'Response.Write "      <td height='25'>"
    'Response.Write "        <select name='OpenType' id='OpenType'>"
    'Response.Write "          <option value='0' "
    'If Trim(OpenType) = "0" Then Response.Write "selected"
    'Response.Write ">在原窗口打开</option>"
    'Response.Write "          <option value='1' "
    'If Trim(OpenType) = "1" Then Response.Write "selected"
    'Response.Write ">在新窗口打开</option>"
    'Response.Write "        </select>"
    'Response.Write "      </td>"
    'Response.Write "    </tr>"
End Sub

Sub GetPositionList()
    Response.Write "    <tr class=tdbg>"
    Response.Write "      <td align=left height=25>显示职位数：</td>"
    Response.Write "      <td colspan='1'><input name='PositionNum'  type='text' size='12' value='"
    If Trim(PositionNum) = "" Then
        Response.Write "0"
    Else
        Response.Write PositionNum
    End If
    Response.Write "'></td>"
    Response.Write "    </tr>"
    Response.Write "    <tr class=tdbg>"
    Response.Write "       <td align=left height=25>是否紧急招聘：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=IsUrgent name=IsUrgent>"
    Response.Write "             <Option value='True'"
    If Trim(IsUrgent) = "True" Then Response.Write "selected"
    Response.Write ">紧急招聘</Option>"
    Response.Write "             <Option value='False'"
    If Trim(IsUrgent) = "False" Then Response.Write "selected"
    Response.Write ">所有招聘</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>日期范围：</td>"
    Response.Write "       <td><input name='DateNum'  type='text' size='12' value='"
    If Trim(DateNum) = "" Then
        Response.Write "0"
    Else
        Response.Write DateNum
    End If
    Response.Write "'>"
    Response.Write "       &nbsp;&nbsp;&nbsp;<font color='red'>如果大于0，则只显示最近几天内更新的职位</font></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>排序方式</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=OrderType name=OrderType>"
    Response.Write "             <Option value='1'"
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">按职位ID降序</Option>"
    Response.Write "             <Option value='2'"
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">按职位ID升序</Option>"
    Response.Write "             <Option value='3'"
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">按更新时间降序</Option>"
    Response.Write "             <Option value='4'"
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">按更新时间升序</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>职位显示方式:</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "            <Select id=ShowType name=ShowType>"
    Response.Write "               <Option value='1'"
    If Trim(ShowType) = "1" Then Response.Write "selected"
    Response.Write ">紧急招聘样式</Option>"
    Response.Write "               <Option value='2'"
    If Trim(ShowType) = "2" Then Response.Write "selected"
    Response.Write ">最新招聘样式</Option>"
    Response.Write "               <Option value='3'"
    If Trim(ShowType) = "3" Then Response.Write "selected"
    Response.Write ">职位信息列表</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>职位名称长度：</td>"
    Response.Write "       <td><input name='TitleLen' type='text' size='12' value='"
    If Trim(TitleLen) = "" Then
        Response.Write "0"
    Else
        Response.Write TitleLen
    End If
    Response.Write "'>"
    Response.Write "       &nbsp;&nbsp;&nbsp;<font color='red'>一个汉字=两个英文字符,若为0，则显示完整职位名</font></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "      <td align=left height=25>工作地点名称长度：</td>"
    Response.Write "      <td colspan='1'><input name='WorkPlaceNameLen' type='text' size='12' value='"
    If Trim(WorkPlaceNameLen) = "" Then
        Response.Write "0"
    Else
        Response.Write WorkPlaceNameLen
    End If
    Response.Write "'></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "      <td align=left height=25>用人单位名称长度：</td>"
    Response.Write "      <td colspan='1'><input name='SubCompanyNameLen' type='text' size='12' value='"
    If Trim(SubCompanyNameLen) = "" Then
        Response.Write "0"
    Else
        Response.Write SubCompanyNameLen
    End If
    Response.Write "'></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>名称过长时否显示省略号设置：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "        <Input id='PShowPoints' type='checkbox' value='True' name='PShowPoints' "
    If LCase(Trim(PShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">职位名称"
    Response.Write "         <Input id='WShowPoints' type='checkbox' value='True' name='WShowPoints' "
    If LCase(Trim(WShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">工作地点"
    Response.Write "          <Input id='SShowPoints' type='checkbox' value='True' name='SShowPoints'"
    If LCase(Trim(SShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">用人单位"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>更新日期显示样式：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=ShowDateType name=ShowDateType>"
    Response.Write "             <Option value='0'"
    If Trim(ShowDateType) = "0" Then Response.Write "selected"
    Response.Write ">不显示</Option>"
    Response.Write "             <Option value='1'"
    If Trim(ShowDateType) = "1" Then Response.Write "selected"
    Response.Write ">显示年月日</Option>"
    Response.Write "             <Option value='2'"
    If Trim(ShowDateType) = "2" Then Response.Write "selected"
    Response.Write ">显示月日</Option>"
    Response.Write "             <Option value='3'"
    If Trim(ShowDateType) = "3" Then Response.Write "selected"
    Response.Write ">显示月日（月-日）</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>是否显示各项<br>职位信息选项：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Input id='ShowPositionID' type='checkbox' value='1' name='ShowPositionID'"
    If Trim(ShowPositionID) = "1" Then Response.Write "checked"
    Response.Write ">显示职位ID"
    Response.Write "          <Input id='ShowPositionName' type='checkbox' value='1' name='ShowPositionName'"
    If Trim(ShowPositionName) = "1" Then Response.Write "checked"
    Response.Write ">显示职位名称"
    Response.Write "          <Input id='ShowWorkPlaceName' type='checkbox' value='1' name='ShowWorkPlaceName'"
    If Trim(ShowWorkPlaceName) = "1" Then Response.Write "checked"
    Response.Write ">显示工作地点<br>"
    Response.Write "          <Input id='ShowSubCompanyName' type='checkbox' value='1' name='ShowSubCompanyName'"
    If Trim(ShowSubCompanyName) = "1" Then Response.Write "checked"
    Response.Write ">显示用人单位"
    Response.Write "          <Input id='ShowPositionNum' type='checkbox' value='1' name='ShowPositionNum'"
    If Trim(ShowPositionNum) = "1" Then Response.Write "checked"
    Response.Write ">显示招聘人数"
    Response.Write "          <Input id='ShowPositionStatus' type='checkbox' value='1' name='ShowPositionStatus' "
    If Trim(ShowPositionStatus) = "1" Then Response.Write "checked"
    Response.Write ">显示职位状态<br>"
    Response.Write "          <Input id='ShowValidDate' type='checkbox' value='1' name='ShowValidDate' "
    If Trim(ShowValidDate) = "1" Then Response.Write "checked"
    Response.Write ">显示有效期"
    Response.Write "          <Input id='ShowUrgentSign' type='checkbox' value='True' name='ShowUrgentSign'"
    If Trim(ShowUrgentSign) = "True" Then Response.Write "checked"
    Response.Write ">显示紧急招聘标志"
    Response.Write "          <Input id='ShowNewSign' type='checkbox' value='True' name='ShowNewSign'"
    If Trim(ShowNewSign) = "True" Then Response.Write "checked"
    Response.Write ">显示新招聘标志"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>是否分页显示：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "           <Select id=UsePage name=UsePage>"
    Response.Write "              <Option value='True'"
    If Trim(UsePage) = "True" Then Response.Write "selected"
    Response.Write ">分页显示</Option>"
    Response.Write "              <Option value='False'"
    If Trim(UsePage) = "False" Then Response.Write "selected"
    Response.Write ">不分页显示</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>申请职位页打开方式：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=OpenType name=OpenType>"
    Response.Write "             <Option value='0'"
    If Trim(OpenType) = "0" Then Response.Write "selected"
    Response.Write ">原窗口打开</Option>"
    Response.Write "             <Option value='1'"
    If Trim(OpenType) = "1" Then Response.Write "selected"
    Response.Write ">新窗口打开</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
End Sub

Sub GetSearchResult()
    Response.Write "    <tr class=tdbg>"
    Response.Write "      <td align=left height=25>显示记录数：</td>"
    Response.Write "      <td colspan='1'><input name='ShowNum'  type='text' size='12' value='"
    If Trim(ShowNum) = "" Then
        Response.Write "0"
    Else
        Response.Write ShowNum
    End If
    Response.Write "'></td>"
    Response.Write "    </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>排序方式</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=OrderType name=OrderType>"
    Response.Write "             <Option value='1'"
    If Trim(OrderType) = "1" Then Response.Write "selected"
    Response.Write ">按职位ID降序</Option>"
    Response.Write "             <Option value='2'"
    If Trim(OrderType) = "2" Then Response.Write "selected"
    Response.Write ">按职位ID升序</Option>"
    Response.Write "             <Option value='3'"
    If Trim(OrderType) = "3" Then Response.Write "selected"
    Response.Write ">按更新时间降序</Option>"
    Response.Write "             <Option value='4'"
    If Trim(OrderType) = "4" Then Response.Write "selected"
    Response.Write ">按更新时间升序</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>职位名称长度：</td>"
    Response.Write "       <td><input name='TitleLen' type='text' size='12' value='"
    If Trim(TitleLen) = "" Then
        Response.Write "0"
    Else
        Response.Write TitleLen
    End If
    Response.Write "'>"
    Response.Write "       &nbsp;&nbsp;&nbsp;<font color='red'>一个汉字=两个英文字符,若为0，则显示完整职位名</font></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "      <td align=left height=25>工作地点名称长度：</td>"
    Response.Write "      <td colspan='1'><input name='WorkPlaceNameLen' type='text' size='12' value='"
    If Trim(WorkPlaceNameLen) = "" Then
        Response.Write "0"
    Else
        Response.Write WorkPlaceNameLen
    End If
    Response.Write "'></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "      <td align=left height=25>用人单位名称长度：</td>"
    Response.Write "      <td colspan='1'><input name='SubCompanyNameLen' type='text' size='12' value='"
    If Trim(SubCompanyNameLen) = "" Then
        Response.Write "0"
    Else
        Response.Write SubCompanyNameLen
    End If
    Response.Write "'></td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>名称过长时否显示省略号设置：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "        <Input id='PShowPoints' type='checkbox' value='True' name='PShowPoints' "
    If LCase(Trim(PShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">职位名称"
    Response.Write "         <Input id='WShowPoints' type='checkbox' value='True' name='WShowPoints' "
    If LCase(Trim(WShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">工作地点"
    Response.Write "          <Input id='SShowPoints' type='checkbox' value='True' name='SShowPoints'"
    If LCase(Trim(SShowPoints)) = "true" Then
        Response.Write "checked"
    End If
    Response.Write ">用人单位"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>更新日期显示样式：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=ShowDateType name=ShowDateType>"
    Response.Write "             <Option value='0'"
    If Trim(ShowDateType) = "0" Then Response.Write "selected"
    Response.Write ">不显示</Option>"
    Response.Write "             <Option value='1'"
    If Trim(ShowDateType) = "1" Then Response.Write "selected"
    Response.Write ">显示年月日</Option>"
    Response.Write "             <Option value='2'"
    If Trim(ShowDateType) = "2" Then Response.Write "selected"
    Response.Write ">显示月日</Option>"
    Response.Write "             <Option value='3'"
    If Trim(ShowDateType) = "3" Then Response.Write "selected"
    Response.Write ">显示月日（月-日）</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>是否显示各项<br>职位信息选项：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Input id='ShowPositionID' type='checkbox' value='1' name='ShowPositionID'"
    If Trim(ShowPositionID) = "1" Then Response.Write "checked"
    Response.Write ">显示职位ID"
    Response.Write "          <Input id='ShowPositionName' type='checkbox' value='1' name='ShowPositionName'"
    If Trim(ShowPositionName) = "1" Then Response.Write "checked"
    Response.Write ">显示职位名称"
    Response.Write "          <Input id='ShowWorkPlaceName' type='checkbox' value='1' name='ShowWorkPlaceName'"
    If Trim(ShowWorkPlaceName) = "1" Then Response.Write "checked"
    Response.Write ">显示工作地点<br>"
    Response.Write "          <Input id='ShowSubCompanyName' type='checkbox' value='1' name='ShowSubCompanyName'"
    If Trim(ShowSubCompanyName) = "1" Then Response.Write "checked"
    Response.Write ">显示用人单位"
    Response.Write "          <Input id='ShowPositionNum' type='checkbox' value='1' name='ShowPositionNum'"
    If Trim(ShowPositionNum) = "1" Then Response.Write "checked"
    Response.Write ">显示招聘人数"
    Response.Write "          <Input id='ShowPositionStatus' type='checkbox' value='1' name='ShowPositionStatus' "
    If Trim(ShowPositionStatus) = "1" Then Response.Write "checked"
    Response.Write ">显示职位状态<br>"
    Response.Write "          <Input id='ShowValidDate' type='checkbox' value='1' name='ShowValidDate' "
    If Trim(ShowValidDate) = "1" Then Response.Write "checked"
    Response.Write ">显示有效期"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>是否分页显示：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "           <Select id=UsePage name=UsePage>"
    Response.Write "              <Option value='True'"
    If Trim(UsePage) = "True" Then Response.Write "selected"
    Response.Write ">分页显示</Option>"
    Response.Write "              <Option value='False'"
    If Trim(UsePage) = "False" Then Response.Write "selected"
    Response.Write ">不分页显示</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
    Response.Write "     <tr class=tdbg>"
    Response.Write "       <td align=left height=25>申请职位页打开方式：</td>"
    Response.Write "       <td height=25 colspan='1'>"
    Response.Write "          <Select id=OpenType name=OpenType>"
    Response.Write "             <Option value='0'"
    If Trim(OpenType) = "0" Then Response.Write "selected"
    Response.Write ">原窗口打开</Option>"
    Response.Write "             <Option value='1'"
    If Trim(OpenType) = "1" Then Response.Write "selected"
    Response.Write ">新窗口打开</Option>"
    Response.Write "          </Select>"
    Response.Write "       </td>"
    Response.Write "     </tr>"
End Sub

Private Function GetSpecial_Option(iChannelID, SpecialID)
    Dim sqlSpecial, rsSpecial, strOption, strOptionTemp
    sqlSpecial = "select ChannelID,SpecialID,SpecialName,OrderID from PE_Special where ChannelID=0 or ChannelID=" & iChannelID & "   order by ChannelID,OrderID"
    Set rsSpecial = Conn.Execute(sqlSpecial)
    If LCase(SpecialID) <> "specialid" Then
        If PE_CLng(SpecialID) = 0 Then
            strOption = "<option value='0'>不属于任何专题</option>"
        Else
            strOption = "<option value='0' selected>不属于任何专题</option>"
        End If
    End If
    If rsSpecial.bof And rsSpecial.bof Then
    Else
        Do While Not rsSpecial.EOF
            If rsSpecial("ChannelID") > 0 Then
                strOptionTemp = rsSpecial("SpecialName") & "（本频道）"
            Else
                strOptionTemp = rsSpecial("SpecialName") & "（全站）"
            End If
            If rsSpecial("SpecialID") = PE_CLng(SpecialID) Then
                strOption = strOption & "<option value='" & rsSpecial("SpecialID") & "' selected>" & strOptionTemp & "</option>"
            Else
                strOption = strOption & "<option value='" & rsSpecial("SpecialID") & "'>" & strOptionTemp & "</option>"
            End If
            rsSpecial.movenext
        Loop
    End If
    rsSpecial.Close
    Set rsSpecial = Nothing
    strOption = strOption & "<option value='SpecialID'"
    If SpecialID = "SpecialID" Then strOption = strOption & " selected"
    strOption = strOption & ">当前频道</option>"

    GetSpecial_Option = strOption
End Function

Private Function GetChannel_Option(iModuleType, ChannelID)
    Dim strChannel, sqlChannel, rsChannel
    sqlChannel = "select ChannelID,ChannelName from PE_Channel  where ModuleType=" & iModuleType & " and Disabled=" & PE_False & " and ChannelType<=1 order by OrderID"
    Set rsChannel = Conn.Execute(sqlChannel)
    Do While Not rsChannel.EOF
        If rsChannel(0) = PE_CLng(ChannelID) Then
            strChannel = strChannel & "<option value='" & rsChannel(0) & "' selected>" & rsChannel(1) & "</option>"
        Else
            strChannel = strChannel & "<option value='" & rsChannel(0) & "'>" & rsChannel(1) & "</option>"
        End If
        rsChannel.movenext
    Loop
    rsChannel.Close
    Set rsChannel = Nothing
    strChannel = strChannel & "<option value='0'"
    If ChannelID = "0" Then strChannel = strChannel & " selected"
    strChannel = strChannel & ">所有同类频道</option>"

    strChannel = strChannel & "<option value='ChannelID'"
    If ChannelID = "ChannelID" Then strChannel = strChannel & " selected"
    strChannel = strChannel & ">当前频道</option>"
    GetChannel_Option = strChannel
End Function

Private Function GetClass_Channel(ChannelID, ClassID, NClassID)
    Dim rsClass, sqlClass, strClass_Option, tmpDepth, i, classcss
    Dim arrShowLine(20)
    For i = 0 To UBound(arrShowLine)
    arrShowLine(i) = False
    Next
    sqlClass = "Select * from PE_Class where ChannelID=" & ChannelID & " order by RootID,OrderID"
    Set rsClass = Conn.Execute(sqlClass)
    If rsClass.bof And rsClass.bof Then
    strClass_Option = strClass_Option & "<option value='0'>请先添加栏目</option>"
    Else
        Do While Not rsClass.EOF
            tmpDepth = rsClass("Depth")
            If rsClass("NextID") > 0 Then
                arrShowLine(tmpDepth) = True
            Else
                arrShowLine(tmpDepth) = False
            End If

            If rsClass("ClassType") = 2 Then
                strClass_Option = strClass_Option & "<option value=''"
            Else
                strClass_Option = strClass_Option & "<option value='" & rsClass("ClassID") & "'"
                If NClassID = False Then
                    If ClassID <> "rsClass_arrChildID" Or ClassID <> "ClassID" Or ClassID <> "arrChildID" Then
                        If rsClass("ClassID") = PE_CLng(ClassID) Then
                            strClass_Option = strClass_Option & " selected"
                        End If
                    End If
                Else
                    If FoundInArr(ClassID, rsClass("ClassID"), "|") = True Then
                        strClass_Option = strClass_Option & " selected"
                    End If
                End If
            End If
            strClass_Option = strClass_Option & ">"
            
            If tmpDepth > 0 Then
            For i = 1 To tmpDepth
                strClass_Option = strClass_Option & "&nbsp;&nbsp;"
                If i = tmpDepth Then
                If rsClass("NextID") > 0 Then
                    strClass_Option = strClass_Option & "├&nbsp;"
                Else
                    strClass_Option = strClass_Option & "└&nbsp;"
                End If
                Else
                If arrShowLine(i) = True Then
                    strClass_Option = strClass_Option & "│"
                Else
                    strClass_Option = strClass_Option & "&nbsp;"
                End If
                End If
            Next
            End If
            strClass_Option = strClass_Option & rsClass("ClassName")
            If rsClass("ClassType") = 2 Then
                strClass_Option = strClass_Option & "（外）"
            End If
            strClass_Option = strClass_Option & "</option>"
            rsClass.movenext
        Loop
    End If
    rsClass.Close
    Set rsClass = Nothing
    If NClassID = False Then
        classcss = "style=''"
    Else
        classcss = "style='background:red'"
    End If
    
    If Trim(ClassID) = "rsClass_arrChildID" Then
        strClass_Option = strClass_Option & "<option value='rsClass_arrChildID' " & classcss & " selected >栏目循环中的栏目</option>"
    Else
        strClass_Option = strClass_Option & "<option value='rsClass_arrChildID' " & classcss & " >栏目循环中的栏目</option>"
    End If
    If Trim(ClassID) = "ClassID" Then
        strClass_Option = strClass_Option & "<option value='ClassID' " & classcss & " selected>当前栏目（不包含子栏目）</option>"
    Else
        strClass_Option = strClass_Option & "<option value='ClassID' " & classcss & ">当前栏目（不包含子栏目）</option>"
    End If
    If Trim(ClassID) = "arrChildID" Then
        strClass_Option = strClass_Option & "<option value='arrChildID' " & classcss & " selected>当前栏目及子栏目</option>"
    Else
        strClass_Option = strClass_Option & "<option value='arrChildID' " & classcss & ">当前栏目及子栏目</option>"
    End If
    If Trim(ClassID) = "0" Then
        strClass_Option = strClass_Option & "<option value='0' " & classcss & " selected>显示所有栏目</option>"
    Else
        strClass_Option = strClass_Option & "<option value='0' " & classcss & ">显示所有栏目</option>"
    End If

    If Trim(ClassID) = "-1" Then
        strClass_Option = strClass_Option & "<option value='-1' " & classcss & " selected>未指定任何栏目</option>"
    Else
        strClass_Option = strClass_Option & "<option value='-1' " & classcss & ">未指定任何栏目</option>"
    End If

    GetClass_Channel = strClass_Option
End Function

Public Function FoundInArr(strArr, strItem, strSplit)
    Dim arrTemp, i
    FoundInArr = False
    If InStr(strArr, strSplit) > 0 Then
        arrTemp = Split(strArr, strSplit)
        For i = 0 To UBound(arrTemp)
            If Trim(arrTemp(i)) = Trim(strItem) Then
                FoundInArr = True
                Exit For
            End If
        Next
    Else
        If Trim(strArr) = Trim(strItem) Then
            FoundInArr = True
        End If
    End If
End Function

Public Function PE_CLng(ByVal str1)
    If IsNumeric(str1) Then
        PE_CLng = CLng(str1)
    Else
        PE_CLng = 0
    End If
End Function

Private Sub Modifylabel()
    If InStr(Trim(request.querystring("editLabel")), "{$") = 0 Then
        Response.Write "<center><br><font color=red>您选择的不是标签</font></center>"
        Response.End
    End If

    Dim editLabeltemp
    editLabeltemp = Trim(Replace(Replace(Trim(request.querystring("editLabel")), "{$", ""), "}", ""))
    editLabeltemp = Replace(editLabeltemp, """", "")
    Action = Left(editLabeltemp, InStr(Trim(Replace(Replace(editLabeltemp, "{$", ""), "}", "")), "(") - 1)
    editLabeltemp = Trim(Replace(Replace(Replace(editLabeltemp, "(", ""), ")", ""), Action, ""))
    Labletemp = Split(editLabeltemp, ",")
     
    Select Case Action
    
    Case "GetArticleList"
        ChannelShortName = "文章"
        ChannelShowType = "GetList"
        imageproperty = "article"
        ModuleType = 1
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        urltype = "0"
        Num = Labletemp(5)
        IsHot = Labletemp(6)
        IsElite = Labletemp(7)
        AuthorName = Labletemp(8)
        DateNum = Labletemp(9)
        OrderType = Labletemp(10)
        ShowType = Labletemp(11)
        TitleLen = Labletemp(12)
        ContentLen = Labletemp(13)
        ShowClassName = Labletemp(14)
        ShowPropertyType = Labletemp(15)
        ShowIncludePic = Labletemp(16)
        ShowAuthor = Labletemp(17)
        ShowDateType = Labletemp(18)
        ShowHits = Labletemp(19)
        ShowHotSign = Labletemp(20)
        ShowNewSign = Labletemp(21)
        ShowTips = Labletemp(22)
        ShowCommentLink = Labletemp(23)
        UsePage = Labletemp(24)
        OpenType = Labletemp(25)
        If UBound(Labletemp) = 26 Then
            Cols = Labletemp(26)
        End If
        If UBound(Labletemp) >= 29 Then
            Cols = Labletemp(26)
            CssNameA = Labletemp(27)
            CssName1 = Labletemp(28)
            CssName2 = Labletemp(29)
        End If

        If UBound(Labletemp) >= 30 Then
            IntervalLines = Labletemp(30)
        End If
     Case "GetPicArticle"
        ChannelShortName = "文章"
        imageproperty = "article"
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        Num = Labletemp(4)
        IsHot = Labletemp(5)
        IsElite = Labletemp(6)
        DateNum = Labletemp(7)
        OrderType = Labletemp(8)
        ShowType = Labletemp(9)
        ImgWidth = Labletemp(10)
        ImgHeight = Labletemp(11)
        TitleLen = Labletemp(12)
        ContentLen = Labletemp(13)
        ShowTips = Labletemp(14)
        Cols = Labletemp(15)
        ChannelShowType = "GetPic"
        ModuleType = 1
     Case "GetSlidePicArticle"
        ChannelShortName = "文章"
        imageproperty = "article"
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        Num = Labletemp(4)
        IsHot = Labletemp(5)
        IsElite = Labletemp(6)
        DateNum = Labletemp(7)
        OrderType = Labletemp(8)
        ImgWidth = Labletemp(9)
        ImgHeight = Labletemp(10)
        TitleLen = Labletemp(11)
        iTimeOut = Labletemp(12)
        effectID = Labletemp(13)
        ChannelShowType = "GetSlide"
        ModuleType = 1
     Case "GetSoftList"
        ChannelShortName = "软件"
        imageproperty = "Soft"
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        urltype = "0"
        Num = Labletemp(5)
        IsHot = Labletemp(6)
        IsElite = Labletemp(7)
        AuthorName = Labletemp(8)
        DateNum = Labletemp(9)
        OrderType = Labletemp(10)
        ShowType = Labletemp(11)
        TitleLen = Labletemp(12)
        ContentLen = Labletemp(13)
        ShowClassName = Labletemp(14)
        ShowPropertyType = Labletemp(15)
        ShowAuthor = Labletemp(16)
        ShowDateType = Labletemp(17)
        ShowHits = Labletemp(18)
        ShowHotSign = Labletemp(19)
        ShowNewSign = Labletemp(20)
        ShowTips = Labletemp(21)
        UsePage = Labletemp(22)
        OpenType = Labletemp(23)
        If UBound(Labletemp) >= 27 Then
            Cols = Labletemp(24)
            CssNameA = Labletemp(25)
            CssName1 = Labletemp(26)
            CssName2 = Labletemp(27)
        End If
        If UBound(Labletemp) >= 28 Then
            IntervalLines = Labletemp(28)
        End If
        ChannelShowType = "GetList"
        ModuleType = 2
     Case "GetPicSoft"
        ChannelShortName = "软件"
        imageproperty = "Soft"
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        Num = Labletemp(4)
        IsHot = Labletemp(5)
        IsElite = Labletemp(6)
        DateNum = Labletemp(7)
        OrderType = Labletemp(8)
        ShowType = Labletemp(9)
        ImgWidth = Labletemp(10)
        ImgHeight = Labletemp(11)
        TitleLen = Labletemp(12)
        ContentLen = Labletemp(13)
        ShowTips = Labletemp(14)
        Cols = Labletemp(15)
        ChannelShowType = "GetPic"
        ModuleType = 2
     Case "GetSlidePicSoft"
        ChannelShortName = "软件"
        imageproperty = "Soft"
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        Num = Labletemp(4)
        IsHot = Labletemp(5)
        IsElite = Labletemp(6)
        DateNum = Labletemp(7)
        OrderType = Labletemp(8)
        ImgWidth = Labletemp(9)
        ImgHeight = Labletemp(10)
        TitleLen = Labletemp(11)
        iTimeOut = Labletemp(12)
        effectID = Labletemp(13)
        ChannelShowType = "GetSlide"
        ModuleType = 2
     Case "GetPhotoList"
        ChannelShortName = "图片"
        imageproperty = "Photo"
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        urltype = "0"
        Num = Labletemp(5)
        IsHot = Labletemp(6)
        IsElite = Labletemp(7)
        AuthorName = Labletemp(8)
        DateNum = Labletemp(9)
        OrderType = Labletemp(10)
        ShowType = Labletemp(11)
        TitleLen = Labletemp(12)
        ContentLen = Labletemp(13)
        ShowClassName = Labletemp(14)
        ShowPropertyType = Labletemp(15)
        ShowAuthor = Labletemp(16)
        ShowDateType = Labletemp(17)
        ShowHits = Labletemp(18)
        ShowHotSign = Labletemp(19)
        ShowNewSign = Labletemp(20)
        ShowTips = Labletemp(21)
        UsePage = Labletemp(22)
        OpenType = Labletemp(23)
        If UBound(Labletemp) >= 27 Then
            Cols = Labletemp(24)
            CssNameA = Labletemp(25)
            CssName1 = Labletemp(26)
            CssName2 = Labletemp(27)
        End If
        If UBound(Labletemp) >= 28 Then
            IntervalLines = Labletemp(28)
        End If
        ChannelShowType = "GetList"
        ModuleType = 3
     Case "GetPicPhoto"
        ChannelShortName = "图片"
        imageproperty = "Photo"
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        Num = Labletemp(4)
        IsHot = Labletemp(5)
        IsElite = Labletemp(6)
        DateNum = Labletemp(7)
        OrderType = Labletemp(8)
        ShowType = Labletemp(9)
        ImgWidth = Labletemp(10)
        ImgHeight = Labletemp(11)
        TitleLen = Labletemp(12)
        ContentLen = Labletemp(13)
        ShowTips = Labletemp(14)
        Cols = Labletemp(15)
        ChannelShowType = "GetPic"
        ModuleType = 3

     Case "GetSlidePicPhoto"
        ChannelShortName = "图片"
        imageproperty = "Photo"
        ChannelID = Labletemp(0)
        ClassID = Labletemp(1)
        If InStr(Labletemp(1), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(2)
        SpecialID = Labletemp(3)
        Num = Labletemp(4)
        IsHot = Labletemp(5)
        IsElite = Labletemp(6)
        DateNum = Labletemp(7)
        OrderType = Labletemp(8)
        ImgWidth = Labletemp(9)
        ImgHeight = Labletemp(10)
        TitleLen = Labletemp(11)
        iTimeOut = Labletemp(12)
        effectID = Labletemp(13)
        ChannelShowType = "GetSlide"
        ModuleType = 3
     Case "GetProductList"
        ChannelShortName = "商品"
        imageproperty = "Product"
        ChannelID = 1000
        ClassID = Labletemp(0)
        If InStr(Labletemp(0), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(1)
        SpecialID = Labletemp(2)
        Num = Labletemp(3)
        ProductType = Labletemp(4)
        IsHot = Labletemp(5)
        IsElite = Labletemp(6)
        DateNum = Labletemp(7)
        OrderType = Labletemp(8)
        ShowType = Labletemp(9)
        TitleLen = Labletemp(10)
        ContentLen = Labletemp(11)
        ShowClassName = Labletemp(12)
        ShowPropertyType = Labletemp(13)
        ShowDateType = Labletemp(14)
        ShowHotSign = Labletemp(15)
        ShowNewSign = Labletemp(16)
        UsePage = Labletemp(17)
        OpenType = Labletemp(18)
        If UBound(Labletemp) >= 39 Then
            IntervalLines = Labletemp(19)
            Cols = Labletemp(20)
            ShowTableTitle = Labletemp(21)
            TableTitleStr = Labletemp(22)
            ShowProductModel = Labletemp(23)
            ShowProductStandard = Labletemp(24)
            ShowUnit = Labletemp(25)
            ShowStocksType = Labletemp(26)
            ShowWeight = Labletemp(27)
            ShowPrice_Market = Labletemp(28)
            ShowPrice_Original = Labletemp(29)
            ShowPrice = Labletemp(30)
            ShowPrice_Member = Labletemp(31)
            ShowDiscount = Labletemp(32)
            ShowButtonType = Labletemp(33)
            ButtonStyle = Labletemp(34)

            CssNameTable = Labletemp(35)
            CssNameTitle = Labletemp(36)
            CssNameA = Labletemp(37)
            CssName1 = Labletemp(38)
            CssName2 = Labletemp(39)
        End If
        urltype = "0"
        ChannelShowType = "GetList"
        ModuleType = 5
    Case "GetPicProduct"
        ChannelShortName = "商品"
        imageproperty = "Product"
        ChannelID = 1000
        ClassID = Labletemp(0)
        If InStr(Labletemp(0), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(1)
        SpecialID = Labletemp(2)
        Num = Labletemp(3)
        ProductType = Labletemp(4)
        IsHot = Labletemp(5)
        IsElite = Labletemp(6)
        DateNum = Labletemp(7)
        OrderType = Labletemp(8)
        ShowType = Labletemp(9)
        ImgWidth = Labletemp(10)
        ImgHeight = Labletemp(11)
        TitleLen = Labletemp(12)
        Cols = Labletemp(13)
        If UBound(Labletemp) >= 18 Then
            ShowPriceType = Labletemp(14)
            ShowDiscount = Labletemp(15)
            ShowButtonType = Labletemp(16)
            ButtonStyle = Labletemp(17)
            OpenType = Labletemp(18)
        End If
        ChannelShowType = "GetPic"
        ModuleType = 5
    Case "GetSlidePicProduct"
        ChannelID = 1000
        ChannelShortName = "商品"
        imageproperty = "Product"
        ClassID = Labletemp(0)
        If InStr(Labletemp(0), "|") > 0 Then
            NClassID = True
        Else
            NClassID = False
        End If
        IncludeChild = Labletemp(1)
        SpecialID = Labletemp(2)
        Num = Labletemp(3)
        IsHot = Labletemp(4)
        IsElite = Labletemp(5)
        DateNum = Labletemp(6)
        OrderType = Labletemp(7)
        ImgWidth = Labletemp(8)
        ImgHeight = Labletemp(9)
        TitleLen = Labletemp(10)
        iTimeOut = Labletemp(11)
        effectID = Labletemp(12)
        If UBound(Labletemp) >= 13 Then
            OpenType = Labletemp(13)
        End If
        ChannelShowType = "GetSlide"
        ModuleType = 5
    Case "GetPositionList"
        ChannelShortName = "职位"
        ChannelShowType = "GetPositionList"
        imageproperty = "Job"
        ModuleType = 8
        PositionNum = Labletemp(0)
        IsUrgent = Labletemp(1)
        DateNum = Labletemp(2)
        OrderType = Labletemp(3)
        ShowType = Labletemp(4)
        TitleLen = Labletemp(5)
        WorkPlaceNameLen = Labletemp(6)
        SubCompanyNameLen = Labletemp(7)
        PShowPoints = Labletemp(8)
        WShowPoints = Labletemp(9)
        SShowPoints = Labletemp(10)
        ShowDateType = Labletemp(11)
        ShowPositionID = Labletemp(12)
        ShowPositionName = Labletemp(13)
        ShowWorkPlaceName = Labletemp(14)
        ShowSubCompanyName = Labletemp(15)
        ShowPositionNum = Labletemp(16)
        ShowPositionStatus = Labletemp(17)
        ShowValidDate = Labletemp(18)
        If Labletemp(4) = 2 Or Labletemp(4) = 3 Then
            ShowUrgentSign = False
        Else
            ShowUrgentSign = Labletemp(19)
        End If
        If Labletemp(4) = 1 Or Labletemp(4) = 3 Then
            ShowNewSign = False
        Else
            ShowNewSign = Labletemp(20)
        End If
        If Labletemp(4) = 1 Or Labletemp(4) = 2 Then
            UsePage = False
        Else
            UsePage = Labletemp(21)
        End If
        OpenType = Labletemp(22)
    Case "GetSearchResult"
        ChannelShortName = "职位"
        ChannelShowType = "GetSearchResult"
        imageproperty = "Job"
        ModuleType = 8
        ShowNum = Labletemp(0)
        OrderType = Labletemp(1)
        TitleLen = Labletemp(2)
        WorkPlaceNameLen = Labletemp(3)
        SubCompanyNameLen = Labletemp(4)
        PShowPoints = Labletemp(5)
        WShowPoints = Labletemp(6)
        SShowPoints = Labletemp(7)
        ShowDateType = Labletemp(8)
        ShowPositionID = Labletemp(9)
        ShowPositionName = Labletemp(10)
        ShowWorkPlaceName = Labletemp(11)
        ShowSubCompanyName = Labletemp(12)
        ShowPositionNum = Labletemp(13)
        ShowPositionStatus = Labletemp(14)
        ShowValidDate = Labletemp(15)
        If Labletemp(4) = 1 Or Labletemp(4) = 2 Then
            UsePage = False
        Else
            UsePage = Labletemp(16)
        End If
        OpenType = Labletemp(17)
    Case Else
        Response.Write "<center><br><font color=red>您选择的不是标签</font></center>"
        Response.End
    End Select
End Sub

Private Sub CellNclass()
    Response.Write "    if (document.myform.NClassChild.checked==true){ " & vbCrLf
    Response.Write "        var Nclassidzhi=""""" & vbCrLf
    Response.Write "        for(var i=0;i<document.myform.ClassID.length;i++){" & vbCrLf
    Response.Write "            if (document.myform.ClassID.options[i].selected==true){" & vbCrLf
    Response.Write "                if (document.myform.ClassID.options[i].value==""rsClass_arrChildID""||document.myform.ClassID.options[i].value==""ClassID""||document.myform.ClassID.options[i].value==""arrChildID""||document.myform.ClassID.options[i].value==""0""){" & vbCrLf
    Response.Write "                    alert(""您在多选中选择了红色部分，多选栏目中是不能包含那部分的。"");" & vbCrLf
    Response.Write "                    return false" & vbCrLf
    Response.Write "                }else{" & vbCrLf
    Response.Write "                    if (Nclassidzhi==""""){" & vbCrLf
    Response.Write "                        Nclassidzhi+=document.myform.ClassID.options[i].value;" & vbCrLf
    Response.Write "                    }else{" & vbCrLf
    Response.Write "                        Nclassidzhi+=""|""+document.myform.ClassID.options[i].value;" & vbCrLf
    Response.Write "                    }" & vbCrLf
    Response.Write "                }" & vbCrLf
    Response.Write "            }" & vbCrLf
    Response.Write "        }" & vbCrLf
    Response.Write "        strJS+=Nclassidzhi;" & vbCrLf
    Response.Write "    }else{" & vbCrLf
    Response.Write "        strJS+=document.myform.ClassID.value;" & vbCrLf
    Response.Write "    }" & vbCrLf
End Sub

Private Sub strJS()
    Response.Write "<script language=""JavaScript"" type=""text/JavaScript"">" & vbCrLf
    Response.Write "function makejs(Parameter,Type)" & vbCrLf
    Response.Write "{" & vbCrLf
    If ModuleType <> 8 Then
        Response.Write "    if (document.myform.ClassID.value==''){" & vbCrLf
        Response.Write "        alert('" & ChannelShortName & "所属栏目不能指定为外部栏目！');" & vbCrLf
        Response.Write "        document.myform.ClassID.focus();" & vbCrLf
        Response.Write "        return false;" & vbCrLf
        Response.Write "    }" & vbCrLf
    End If
    Response.Write "    var strJS;" & vbCrLf
    If editLabel = "" And InsertTemplate = 0 Then
        If ModuleType <> 8 Then
            Response.Write "    if (document.myform.lableExplain.value !=""""){" & vbCrLf
            Response.Write "        strJS=""{$--""+document.myform.lableExplain.value+""--}"";" & vbCrLf
            Response.Write "    }else{" & vbCrLf
            Response.Write "        strJS="""";" & vbCrLf
            Response.Write "    }" & vbCrLf
        Else
            Response.Write "    strJS="""";" & vbCrLf
        End If
        Response.Write "    strJS+=""<IMG  SRC='editor/images/label.gif' BORDER='0' "";" & vbCrLf
        Response.Write "    strJS+=""zzz='\""\""{$""+Parameter+""("";" & vbCrLf
    Else
        Response.Write "    strJS=""{$""+Parameter+""("";" & vbCrLf
    End If
    Response.Write "  switch(Type){" & vbCrLf
    Response.Write "  case ""GetList"":" & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+=document.myform.ChannelID.value;" & vbCrLf
        Response.Write "    strJS+="",""" & vbCrLf
    End If

    Call CellNclass

    Response.Write "    strJS+="",""+document.myform.IncludeChild.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.SpecialID.value;   " & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+="",0""" & vbCrLf
    End If
    Response.Write "    strJS+="",""+document.myform.Num.value;" & vbCrLf
    If ModuleType = 5 Then
        Response.Write "    strJS+="",""+document.myform.ProductType.value;" & vbCrLf
    End If
    Response.Write "    strJS+="",""+document.myform.IsHot.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.IsElite.checked;" & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+="",""+""\""""+document.myform.AuthorName.value+""\"""";" & vbCrLf
    End If
    Response.Write "    strJS+="",""+document.myform.DateNum.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf

    Response.Write "    strJS+="",""+document.myform.ShowType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ContentLen.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowClassName.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowPropertyType.value;" & vbCrLf
    If ModuleType = 1 Then
        Response.Write "    strJS+="",""+document.myform.ShowIncludePic.checked; //A" & vbCrLf
    End If
    If ModuleType <> 5 Then
        Response.Write "    strJS+="",""+document.myform.ShowAuthor.checked;" & vbCrLf
    End If
    Response.Write "    strJS+="",""+document.myform.ShowDateType.value;" & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+="",""+document.myform.ShowHits.checked;" & vbCrLf
    End If
    Response.Write "    strJS+="",""+document.myform.ShowHotSign.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowNewSign.checked;" & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+="",""+document.myform.ShowTips.checked;" & vbCrLf
    End If
    If ModuleType = 1 Then
        Response.Write "    strJS+="",""+document.myform.ShowCommentLink.checked; //A" & vbCrLf
    End If
    Response.Write "    strJS+="",""+document.myform.UsePage.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+="",""+document.myform.Cols.value;            //A" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.CssNameA.value;        //A" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.CssName1.value;        //A" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.CssName2.value;        //A" & vbCrLf
    End If
    If ModuleType = 5 Then
        Response.Write "    strJS+="",""+document.myform.IntervalLines.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.Cols.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowTableTitle.checked;" & vbCrLf
        Response.Write "    var TableTitleStr=""""" & vbCrLf
        Response.Write "    for(var i=1;i<14;i++){" & vbCrLf
        Response.Write "        if (i==13){" & vbCrLf
        Response.Write "            TableTitleStr+=eval(""document.myform.TableTitleStr""+i+"".value"")" & vbCrLf
        Response.Write "        }else{" & vbCrLf
        Response.Write "            TableTitleStr+=eval(""document.myform.TableTitleStr""+i+"".value"")+""|""" & vbCrLf
        Response.Write "        }" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    strJS+="",""+TableTitleStr" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowProductModel.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowProductStandard.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowUnit.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowStocksType.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowWeight.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowPrice_Market.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowPrice_Original.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowPrice.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowPrice_Member.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowDiscount.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowButtonType.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ButtonStyle.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.CssNameTable.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.CssNameTitle.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.CssNameA.value;        //A" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.CssName1.value;        //A" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.CssName2.value;        //A" & vbCrLf
    End If
    Response.Write "    break;" & vbCrLf

    Response.Write "   case ""GetPic"":" & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+=document.myform.ChannelID.value;" & vbCrLf
        Response.Write "    strJS+="",""" & vbCrLf
    End If
    Call CellNclass
    Response.Write "    strJS+="",""+document.myform.IncludeChild.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.SpecialID.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.Num.value;" & vbCrLf
    If ModuleType = 5 Then
        Response.Write "    strJS+="",""+document.myform.ProductType.value;" & vbCrLf
    End If
    Response.Write "    strJS+="",""+document.myform.IsHot.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.IsElite.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.DateNum.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ShowType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ImgWidth.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ImgHeight.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+="",""+document.myform.ContentLen.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowTips.checked;" & vbCrLf
    End If
    Response.Write "    strJS+="",""+document.myform.Cols.value;" & vbCrLf
    If ModuleType = 5 Then
        Response.Write "    strJS+="",""+document.myform.ShowPriceType.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowDiscount.checked;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowButtonType.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ButtonStyle.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
    End If
    Response.Write "    break;" & vbCrLf
    
    Response.Write "   case ""GetSlide"":" & vbCrLf
    If ModuleType <> 5 Then
        Response.Write "    strJS+=document.myform.ChannelID.value;" & vbCrLf
        Response.Write "    strJS+="",""" & vbCrLf
    End If
    Call CellNclass
    Response.Write "    strJS+="",""+document.myform.IncludeChild.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.SpecialID.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.Num.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.IsHot.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.IsElite.checked;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.DateNum.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ImgWidth.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.ImgHeight.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.iTimeOut.value;" & vbCrLf
    Response.Write "    strJS+="",""+document.myform.effectID.value;" & vbCrLf
    'If ModuleType = 5 Then
    '    Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
    'End If
    Response.Write "    break;" & vbCrLf

    If ModuleType = 8 Then
        Response.Write "  case ""GetPositionList"":" & vbCrLf
        Response.Write "    strJS+=document.myform.PositionNum.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.IsUrgent.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.DateNum.value;   " & vbCrLf
        Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowType.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.WorkPlaceNameLen.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.SubCompanyNameLen.value;" & vbCrLf
        Response.Write "    if (document.myform.PShowPoints.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.PShowPoints.checked;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.WShowPoints.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.WShowPoints.checked;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.SShowPoints.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.SShowPoints.checked;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowDateType.value;" & vbCrLf
        Response.Write "    if (document.myform.ShowPositionID.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowPositionID.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowPositionName.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowPositionName.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowWorkPlaceName.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowWorkPlaceName.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowSubCompanyName.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowSubCompanyName.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowPositionNum.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowPositionNum.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowPositionStatus.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowPositionStatus.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowValidDate.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowValidDate.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowUrgentSign.checked ==false||document.myform.ShowType.value==2||document.myform.ShowType.value==3){" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowUrgentSign.value;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowNewSign.checked ==false||document.myform.ShowType.value==1||document.myform.ShowType.value==3){" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowNewSign.value;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowType.value==1||document.myform.ShowType.value==2){" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.UsePage.value;" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
        Response.Write "    break;" & vbCrLf

        Response.Write "  case ""GetSearchResult"":" & vbCrLf
        Response.Write "    strJS+=document.myform.ShowNum.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.OrderType.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.TitleLen.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.WorkPlaceNameLen.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.SubCompanyNameLen.value;" & vbCrLf
        Response.Write "    if (document.myform.PShowPoints.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.PShowPoints.checked;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.WShowPoints.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.WShowPoints.checked;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.SShowPoints.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.SShowPoints.checked;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""false"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.ShowDateType.value;" & vbCrLf
        Response.Write "    if (document.myform.ShowPositionID.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowPositionID.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowPositionName.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowPositionName.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowWorkPlaceName.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowWorkPlaceName.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowSubCompanyName.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowSubCompanyName.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowPositionNum.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowPositionNum.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowPositionStatus.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowPositionStatus.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    if (document.myform.ShowValidDate.checked ==true){" & vbCrLf
        Response.Write "        strJS+="",""+document.myform.ShowValidDate.value;" & vbCrLf
        Response.Write "    }else{" & vbCrLf
        Response.Write "        strJS+="",""+""0"";" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.UsePage.value;" & vbCrLf
        Response.Write "    strJS+="",""+document.myform.OpenType.value;" & vbCrLf
        Response.Write "    break;" & vbCrLf
    End If
    Response.Write "    default:" & vbCrLf
    Response.Write "        alert(""错误参数调用！"");" & vbCrLf
    Response.Write "        break;" & vbCrLf
    Response.Write "   }" & vbCrLf
    If editLabel = "" And InsertTemplate = 0 Then
        Response.Write "   strJS+="")}' >"";" & vbCrLf
    Else
        Response.Write "   strJS+="")}"";" & vbCrLf
    End If
    Response.Write "   window.returnValue = strJS;" & vbCrLf
    Response.Write "   window.close();" & vbCrLf
    Response.Write "}" & vbCrLf
    Response.Write "</script>" & vbCrLf

End Sub

%>
<script Language="JavaScript">
function NClassIDChild(){
    if (document.myform.NClassChild.checked==true){
        document.myform.ClassID.size=2;
        document.myform.ClassID.style.height=250;
        document.myform.ClassID.style.width=400;
        document.myform.ClassID.multiple=true;
        for(var i=0;i<document.myform.ClassID.length;i++){
            if (document.myform.ClassID.options[i].value=="rsClass_arrChildID"||document.myform.ClassID.options[i].value=="ClassID"||document.myform.ClassID.options[i].value=="arrChildID"||document.myform.ClassID.options[i].value=="0"){
                document.myform.ClassID.options[i].style.background="red";
            }
        }
    }else{
        document.myform.ClassID.size=1;
        document.myform.ClassID.style.width=200;
        document.myform.ClassID.multiple=false;
        for(var i=0;i<document.myform.ClassID.length;i++){
            if (document.myform.ClassID.options[i].value=="rsClass_arrChildID"||document.myform.ClassID.options[i].value=="ClassID"||document.myform.ClassID.options[i].value=="arrChildID"||document.myform.ClassID.options[i].value=="0"){
                document.myform.ClassID.options[i].style.background="";
            }
        }
    }
}
function change_item(element){
    if(element.selectedIndex!=-1)
    var selectednumber = element.options[element.selectedIndex].value;

    if(selectednumber==1){
        objFiles.style.display="";
        <%
        If ModuleType = 5 Then
        %>
            document.myform.common.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_common.gif"
            document.myform.elite.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_elite.gif"
            document.myform.ontop.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_ontop.gif"
        <%
        Else
        %>
            document.myform.common.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_common.gif"
            document.myform.elite.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_elite.gif"
            document.myform.ontop.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_ontop.gif"
        <%
        End If
        %>
    }
    else if (selectednumber==0)
    {
        objFiles.style.display="none";
    }
    else if (selectednumber==2)
    {
        objFiles.style.display="none";
    }
    else if (selectednumber>=3)
    {
        selectednumber = selectednumber - 1
        objFiles.style.display="";
        <%
        If ModuleType = 5 Then
        %>
            document.myform.common.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_common" + selectednumber + ".gif"
            document.myform.elite.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_elite" + selectednumber + ".gif"
            document.myform.ontop.src = "<%=InstallDir%>Shop/images/<%=imageproperty%>_ontop" + selectednumber + ".gif"
        <%
        Else
        %>
            document.myform.common.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_common" + selectednumber + ".gif"
            document.myform.elite.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_elite" + selectednumber + ".gif"
            document.myform.ontop.src = "<%=InstallDir & imageproperty%>/images/<%=imageproperty%>_ontop" + selectednumber + ".gif"
        <%
        End If
        %>
    }
}
</script>

