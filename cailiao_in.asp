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
td {
	font-family: "宋体";
	font-size: 9pt
}
body {
	font-family: "宋体";
	font-size: 9pt
}
select {
	font-family: "宋体";
	font-size: 9pt
}
A {
	text-decoration: none;
	color: #336699;
	font-family: "宋体";
	font-size: 9pt
}
A:hover {
	text-decoration: underline;
	color: #FF0000;
	font-family: "宋体";
	font-size: 9pt
}
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
.Noprint {
	display:none;
}
.PageNext {
	page-break-after: always;
}
</style>
</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0>
<table width="700" border="0" cellpadding="0" style="border-collapse: collapse" >
<tr>
  <td colspan="2" align=center valign=middle height=50><font size="3">安徽嘉美包装有限公司</font><b><font size="5"><br>
    材 料 入 库 单</font></b></td>
</tr>
<tr>
  <td colspan="2" height=23 align=left><p>&nbsp;&nbsp;&nbsp;&nbsp; 日期：<%=date()%>&nbsp;&nbsp;&nbsp;&nbsp;<b><font size=4 color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; №:
      <%Response.Write(year(date()) & "" & month(date()) & "" & day(date()) & "" & "" & mid("",weekday(date()),1))%>
      <%Function gen_key(digits) 
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
<td>
<table width="700" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
  <tr>
    <td height="30" align=center width="10%"><b><font size="3">品名</font></b></td>
    <td height="30" align=center width="13%"><b><font size="3">规格</font></b></td>
    <td height="30" align=center width="15%"><b><font size="3">数量</font></b></td>
    <td height="30" align=center width="7%"><b><font size="3">单位</font></b></td>
    <td height="30" align=center width="10%"><b><font size="3">单价</font></b></td>
    <td height="30" align=center width="15%"><b><font size="3">金额</font></b></td>
    <td height="30" align=center width="12%"><b><font size="3">供货单位</font></b></td>
    <td height="30" align=center width="18%"><b><font size="3">备注</font></b></td>
  </tr>
  <% 

Sum = trim(request.form("selBigClass").count)
if Sum =0 then 
response.write "<script language=javascript>alert('未选定记录！');history.go(-1);</script>"
response.end
elseif sum>6 then
response.write "<script language=javascript>alert('最多只能选六项！');history.go(-1);</script>"
response.end
else
for i = 1 to Sum
selBigClass= request.form("selBigClass")(i)
set rs=server.createobject("adodb.recordset")
sql="select * from cailiao_in_store where id in ("&selBigClass&")"
rs.open sql,conn,1,1

  %>
  <tr>
    <td height="22" align=center><%=rs("pinming")%></td>
    <td height="22" align=center><%if rs("guige")=0 then%>
      &nbsp;
      <%else%>
      <%=rs("guige")%>*<%=rs("houdou")%>
      <%end if%></td>
    <td height="22" align=center ><%=rs("use_num")%></td>
    <td height="22" align=center ><%=rs("unit")%></td>
    <td height="22" align=center ><%=Formatnumber(rs("danjia"),2,-1,,0)%></td>
    <td height="22" align=center ><%=Formatnumber(rs("jinge"),2,-1,,0)%></td>
    <td height="22" align=center ><%=rs("gonghuodanwei")%></td>
    <td height="22" align=center ><%if rs("content")<>"" then%>
      <%=rs("content")%>
      <%end if%></td>
  </tr>
  <%
num=num+rs("jinge")
next
end if

%>
  <tr>
    <td height="30" colspan="8"><b><font size="2">&nbsp; 金额:</b>(大写)<b>
      <% response.write rmb(num)%>
      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;￥:<u><%=Formatnumber(num,2,-1,,0)%>元</u></b></font></td>
      </font>
      </b>
    </td>
  
    </tr>
  
</table>
</td>
</tr>
<tr>
  <td colspan="8" align=right valign=bottom height=30><p align="left"><font size=2><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </b></font> <font size=1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      白联：存根&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 红联：客户&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      黄联：财务 </font><font size=2><b>&nbsp; <br>
      &nbsp; 仓库主管：_______________&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;经 办 人：</b><%=name%></font><br>
  </td>
</tr>
<table>
  <tr>
    <td height="30" colspan="8" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT>
      </OBJECT>
      <input type=button value=打印 onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT">
      <input type=button value=直接打印 onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT">
      <input type=button value=页面设置 onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT">
      <input type=button value=打印预览 onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT"></td>
  </tr>
</table>
</BODY>
</HTML>

<%
Function rmb(num)
'Dim num
Dim numList 
Dim rmbList 
Dim numLen
Dim numChar
Dim numstr
Dim n 
Dim n1, n2 
Dim hz
numList = "零壹贰叁肆伍陆柒捌玖"
rmbList = "分角元拾佰仟万拾佰仟亿拾佰仟万"

num = FormatNumber(num, 2)
If num > 9999999999999.99 Then
    rmb = "超出范围的人民币值"
    Exit Function
End If

num=CStr(num)
if instr(num,"-")>0then
Tonum=Mid(num,2)
'rmb=Tonum
'rmb="是负数"
'exit Function
'end if


numstr = CStr(Tonum * 100)
else
numstr=CStr(num*100)
end if
numLen = Len(numstr)
n = 1
Do While n <= numLen
    numChar = CInt(Mid(numstr, n, 1))
    n1 = Mid(numList, numChar + 1, 1)
    n2 = Mid(rmbList, numLen - n + 1, 1)
    If Not n1 = "零" Then
        hz = hz + CStr(n1) + CStr(n2)
    Else
        If n2 = "亿" Or n2 = "万" Or n2 = "元" Or n1 = "零" Then
          Do While Right(hz, 1) = "零"
            hz = Left(hz, Len(hz) - 1)
            Loop
        End If
        If (n2 = "亿" Or (n2 = "万" And Right(hz, 1) <> "亿") Or n2 = "元") Then
            hz = hz + CStr(n2)
        Else
            If Left(Right(hz, 2), 1) = "零" Or Right(hz, 1) <> "亿" Then
             hz = hz + n1
            End If
        End If
    End If
    n = n + 1
Loop
Do While Right(hz, 1) = "零"
    hz = Left(hz, Len(hz) - 1)
Loop
 If Right(hz, 1) = "元" Then
    hz = hz + "整"
 End If
 if instr(num,"-")>0then
rmb ="负"& hz
else
rmb=hz
end if
End Function
%>