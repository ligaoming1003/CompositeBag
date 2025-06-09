<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>∷嘉美管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>

<SCRIPT language=javascript  src="js/selectcity.js"></script>
<script language="JavaScript" src="js/validate.js" type="text/JavaScript"></script>

<SCRIPT language=javascript src="css/init.js"></SCRIPT>

<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 9pt}
body {  font-family: "宋体"; font-size: 9pt}
select {  font-family: "宋体"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "宋体"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
-->
</style>




<SCRIPT LANGUAGE=javascript>
<!--
function Juge(myform)
{

	
}


function SelectAll() {
	for (var i=0;i<document.selform.selBigClass.length;i++) {
		var e=document.selform.selBigClass[i];
		e.checked=!e.checked;
	}
}
//-->
</script>


<!--media=print 这个属性可以在打印时有效--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style>
</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0>       

        <table width="700" border="0" cellpadding="0" style="border-collapse: collapse" >

 <tr>
 <td colspan="3" align=center valign=middle height=107><font size="3">安徽嘉美包装有限公司</font><b><font size="5"><br>
	客户发货对账单</font></b><hr width="200" ></td>
</tr>
<tr>
<td height=23 width="20%"  align=left><font size=2>&nbsp;&nbsp;日期：<%=date()%> </font>
</td>

<td height=23 width="40%" align=left>
<p align="right">　<b><font size=4 color="#FF0000">№:<%
Response.Write(year(date()) & "" & month(date()) & "" & day(date()) & "" & "" & mid("",weekday(date()),1))
%><%Function gen_key(digits) 
dim char_array(50) 
char_array(0) = "0" 
char_array(1) = "1" 
char_array(2) = "2" 
char_array(3) = "3" 
char_array(4) = "4" 
char_array(5) = "5" 
char_array(6) = "6" 
char_array(7) = "7" 
char_array(8) = "8" 
char_array(9) = "9" 
randomize 
do while len(output) < digits 
nums = char_array(Int((9 - 0 + 1) * Rnd + 0)) 
output = output + nums 
loop 
gen_key = output 
End Function 
response.write "" & gen_key(3) & "" & vbcrlf 
%> 
</font></b>&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>

<tr>
<td colspan="3">
<table width="700" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse">

<tr>
<td height="30" align=center width="10%"><b><font size="3">时间</font></b></td>
<td height="30" align=center width="10%"><b><font size="3">类别</font></b></td>
<td height="30" align=center width="26%"><b><font size="3">品名</font></b></td>
<td height="30" align=center width="11%"><b><font size="3">规格</font></b></td>
<td height="30" align=center width="12%"><b><font size="3">数量</font></b></td>
<td height="30" align=center width="14%"><b><font size="3">单价</font></b></td>
<td height="30" align=center width="16%"><b><font size="3">金额</font></b></td>
</tr>
<% 

Sum = trim(request.form("selBigClass").count)
if Sum =0 then 
response.write "<script language=javascript>alert('未选定记录！');history.go(-1);</script>"
response.end


else
for i = 1 to Sum
selBigClass= request.form("selBigClass")(i)
set rs=server.createobject("adodb.recordset")
sql="select * from chengping_out_store where id in ("&selBigClass&")"
rs.open sql,conn,1,1

  %>

<tr>
<td align=center><%=rs("uptime")%>　</td>
<td height="25" align=center><%=rs("fenlei")%></td>
	<td height="25" align=center><%=rs("pinming")%><%if rs("content")<>"" then%><font face="华文楷体" size="1">(<%=rs("content")%>)</font><%end if%></td>
<td align=center><%=rs("guige")%>*<%=rs("kuang")%></td>
<td align=center><%=Formatnumber(rs("Loan_num"),2,-1,0,0)%><%=rs("unit")%></td>
<td align=center ><%=Formatnumber(rs("loan_price"),4,-1,0,0)%>元/<%=rs("unit")%></td>
<td align=center><%=Formatnumber(rs("Loan_Amount"),2,-1,,0)%>元</td>

</tr>

<%
num=num+rs("Loan_Amount")
next
end if

%>

<tr>
<td height="30" align=left width="100%" colspan="7"  >
<p><b><font size="2">&nbsp; 金额:</font></b><font size="2">&nbsp;(大写)
<b>
<%     
  dim   num   '要转换成大写的金额     
  dim   atoc   '转换之后的值     
  Dim   String1   '如下定义     
  Dim   String2   '如下定义     
  Dim   String3   '从原A值中取出的值     
  Dim   I   '循环变量     
  Dim   J   'A的值乘以100的字符串长度     
  Dim   Ch1   '数字的汉语读法     
  Dim   Ch2   '数字位的汉字读法     
  Dim   nZero   '用来计算连续的零值是几个     

  String1   =   "零壹贰叁肆伍陆柒捌玖"     
  String2   =   "万仟佰拾亿仟佰拾万仟佰拾元角分"     
  nZero   =   0     
    
  If   InStr(1,   CStr(num   *   100),   ".")   <>   0   Then     
  err.Raise   5000,   ,   "此函数(   AtoC()   )只能转换小数点后有两位以内的数！"     
  End   If     
    
  J   =   Len(CStr(num   *   100))     
  String2   =   Right(String2,   J)   '取出对应位数的STRING2的值     
    
  For   I   =   1   To   J     
  String3   =   Mid(num   *   100,   I,   1)   '取出需转换的某一位的值     
    
  If   I   <>   (J   -   3)   +   1   And   I   <>   (J   -   7)   +   1   And   I   <>   (J   -   11)   +   1   And   I   <>(J   -   15)   +   1   Then     
  If   String3   =   0   Then     
  Ch1   =   ""     
  Ch2   =   ""     
  nZero   =   nZero   +   1     
  ElseIf   String3   <>   0   And   nZero   <>   0   Then     
  Ch1   =   "零"   &   Mid(String1,   clng(String3)   +   1,   1)     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   0     
  Else     
  Ch1   =   Mid(String1,   clng(String3)   +   1,   1)     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   0     
  End   If     
  Else   '该位是万亿，亿，万，元位等关键位     
  If   String3   <>   0   And   nZero   <>   0   Then     
  Ch1   =   "零"   &   Mid(String1,   clng(String3)   +   1,   1)     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   0     
  ElseIf   String3   <>   0   And   nZero   =   0   Then     
  Ch1   =   Mid(String1,   clng(String3)   +   1,   1)     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   0     
  ElseIf   String3   =   0   And   nZero   >=   3   Then     
  Ch1   =   ""     
  Ch2   =   ""     
  nZero   =   nZero   +   1     
  Else     
  Ch1   =   ""     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   nZero   +   1     
  End   If     
    
  If   I   =   (J   -   11)   +   1   Or   I   =   (J   -   3)   +   1   Then   '如果该位是亿位或元位，则必须写上     
  Ch2   =   Mid(String2,   I,   1)     
  End   If     
    
  End   If     
  AtoC   =   AtoC   &   Ch1   &   Ch2     
    
  If   I   =   J   And   String3   =   0   Then   '最后一位（分）为0时，加上“整”     
  AtoC   =   AtoC   &   "整"     
  End   If     
    
  Next     
  if   num=0   then     
  atoc="零元整"     
  end   if    
response.write atoc 
  %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;￥:<u><%=Formatnumber((num),2,-1,,0)%>元</u></b></font></td>




</font>

</td>
</tr>




</table>



</td>
</tr>
<tr>
 <td colspan="3" align=right valign=bottom height=30>
	<p align="left"><font size=2><b>&nbsp;&nbsp; 主管经理：_______________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;经 办 人：</b><%=name%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>

	</table> 			
	<table><tr><td height="30" colspan="8" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=打印 onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=直接打印 onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=页面设置 onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=打印预览 onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>

				
				</table>    
		
</BODY></HTML>

