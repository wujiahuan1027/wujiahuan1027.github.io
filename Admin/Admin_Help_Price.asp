<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>商品最终价格计算公式表</title>
<style type="text/css">
<!--
td {
    font-size: 9pt;
}
.style2 {color: #0000FF}
-->
</style>
</head>

<body>
<table width="760"  border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000066">
  <tr align="center" bgcolor="#FFFFFF">
    <td height="20" colspan="5"><strong>商品最终价格计算公式表</strong></td>
  </tr>
  <tr align="center" bgcolor="#FFFFFF">
    <td height="20" align="center">&nbsp;</td>
    <td height="20">正常销售</td>
    <td>涨价商品</td>
    <td height="20">降价商品</td>
    <td height="20">附赠/换购的礼品</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="40" align="center"><span class="style2">当前零售价</span></td>
    <td align="center"><span class="style2">等于原始零售价</span></td>
    <td align="center"><span class="style2">手工指定</span></td>
    <td><span class="style2">IF 在优惠期限内 THEN<br>
&nbsp;&nbsp;&nbsp;&nbsp;系统自动根据原始零售价及折扣率进行计算，也可以手工指定<br>
ELSE<br>
&nbsp;&nbsp;&nbsp;&nbsp;等于原始零售价（显示为正常销售）<br>
END IF</span></td>
    <td align="center"><span class="style2">手工指定</span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td height="50" align="center">一般客户<br>
    （游客）</td>
    <td colspan="3" align="center">实际价格＝当前零售价</td>
    <td rowspan="2">实际价格＝当前零售价</td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="center">会员</td>
    <td colspan="2">IF 指定了会员价 THEN<br>
&nbsp;&nbsp;&nbsp;&nbsp;会员价＝指定会员价<br>
ELSE<br>
&nbsp;&nbsp;&nbsp;&nbsp;会员价＝当前零售价×会员折扣率<br>
END IF</td>
    <td><p>IF 指定了会员价 THEN（不再享受折上折优惠）<br>
      &nbsp;&nbsp;&nbsp;&nbsp;IF 指定会员价≤当前零售价 THEN<br>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;会员价＝指定会员价<br>
&nbsp;&nbsp;&nbsp;&nbsp;ELSE<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;会员价=当前零售价<br>
&nbsp;&nbsp;&nbsp;&nbsp;END IF <br>
        ELSE<br>
      &nbsp;&nbsp;&nbsp;&nbsp;IF 享有折上折优惠 OR 不在优惠期限内 THEN<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;会员价＝当前零售价×会员折扣率<br>
      &nbsp;&nbsp;&nbsp;&nbsp;ELSE<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;IF 原始零售价×会员折扣率≥当前零售价 THEN<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;会员价＝当前零售价<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ELSE<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;会员价＝原始零售价×会员折扣率<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;END IF<br>
&nbsp;&nbsp;&nbsp;&nbsp;END IF<br>
END IF</p>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="center">代理商</td>
    <td colspan="4">IF 指定了代理价 THEN<br>
&nbsp;&nbsp;&nbsp;&nbsp;代理价＝指定代理价<br>
ELSE<br>
&nbsp;&nbsp;&nbsp; 代理价＝当前零售价×代理折扣率<br>
END IF </td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td align="center">批量购买</td>
    <td colspan="4">IF 一次性购买量≥起批数量1 并且 此产品允许批发 THEN<br>
&nbsp;&nbsp;&nbsp;&nbsp;IF 一次性购买量&lt;起批数量2 THEN<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;实际价格＝批发价1<br>
&nbsp;&nbsp;&nbsp;&nbsp;ELSE<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;IF 一次性购买量&lt;起批数量3 THEN<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;实际价格＝批发价2<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ELSE<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;实际价格＝批发价3<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;END IF<br>
&nbsp;&nbsp;&nbsp;&nbsp;END IF<br>
ELSE<br>
&nbsp;&nbsp;&nbsp;&nbsp;实际价格按零售公式计算<br>
END IF </td>
  </tr>
</table>
</body>
</html>
