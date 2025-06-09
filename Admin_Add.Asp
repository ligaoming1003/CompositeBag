<!--#include file="Conn.Asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript  src="js/selectcity.js"></script>
<script language="JavaScript" src="js/validate.js" type="text/JavaScript"></script>
<SCRIPT language=javascript src="css/init.js"></SCRIPT>
<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 9pt}
body {  font-family: "宋体"; font-size: 9pt}
select {  font-family: "宋体"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "宋体"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
-->
</style>
<title>照片管理系统</title>
<script language="JavaScript">
function checkform(form)
{
if (form.Photo_Name.value=='')
  {
    window.alert('请输入员工姓名！');
    return false;
   }
 
 return true
} 
  
</script>
</head>
<body bgcolor="#E3E3E3">
<!--#include file="top.asp"-->
<%
If Request("Action")=1 then

function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
	fString = replace(fString, """", "&quot;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end function

'------------------检查数据是否为空---------------------------
Photo_Name=trim(Request("Photo_Name"))
ClassName=trim(Request("ClassName"))
if Photo_Name="" or ClassName=""  then
   Response.Write "<script>alert(""系统异常！请勿非法操作！"");location.href=""javascript:history.go(-1)"";</script>"
   response.end()
end if
Photo_author=trim(Request("Photo_author"))
Photo_Time=trim(Request("Photo_Time"))
Photo_Time1=trim(Request("Photo_Time1"))
Photo_Address=trim(Request("Photo_Address"))
shenfenzhenghao=trim(Request("shenfenzhenghao"))
dianhua=trim(Request("dianhua"))
nianling=trim(Request("nianling"))
minzu=trim(Request("minzu"))
jiguan=trim(Request("jiguan"))
zhuzhi=trim(Request("zhuzhi"))
bianhao=trim(Request("bianhao"))
zhiwu=trim(Request("zhiwu"))
gongzhong=trim(Request("gongzhong"))
hunfou=trim(Request("hunfou"))
yajin=trim(Request("yajin"))
baoxian=trim(Request("baoxian"))
dangan=trim(Request("dangan"))

qingkuang=HTMLEncode(Request("qingkuang"))
Photo_Content=HTMLEncode(Request("Photo_Content"))

'-----------------将内容写入数据库--------------------------------------

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from ImgData"
rs.open sql,conn,1,3
rs.addnew
rs("Photo_Name")=Photo_Name
rs("ClassName")=ClassName
rs("Photo_author")=Photo_author
rs("Photo_Time")=Photo_Time
rs("Photo_Time1")=Photo_Time1
rs("Photo_Address")=Photo_Address
rs("shenfenzhenghao")=shenfenzhenghao
rs("dianhua")=dianhua
rs("nianling")=nianling
rs("minzu")=minzu
rs("jiguan")=jiguan
rs("zhuzhi")=zhuzhi
rs("bianhao")=bianhao
rs("zhiwu")=zhiwu
rs("gongzhong")=gongzhong
rs("hunfou")=hunfou
rs("yajin")=yajin
rs("baoxian")=baoxian
rs("dangan")=dangan

rs("Photo_Content")=Photo_Content
rs("qingkuang")=qingkuang
rs("intime")=now()
rs.Update
rs.Close
Set rs=Nothing

'-----------------写入数据库结束----------------------------------------

'获取ID
sqlp="select top 1 id from [ImgData] order by id desc"
set rsp = createobject("ADODB.recordset")
rsp.open sqlp,conn
if not rsp.eof then
   Response.Redirect "Admin_Upload.Asp?id="&rsp("id")&""
   response.end()
end if
rsp.close
set rsp=nothing
End if
%>
<table border="1"  cellpadding="0" cellspacing="0" bordercolor=#0055E6 width="700">
<form ACTION="Admin_Add.Asp" METHOD="POST" onSubmit="return checkform(this)" name="form">

<tr>
<td align=center height=26 background="images/background.gif">
<b><font color="#ffffff">个 人 资 料</font></b>
</td>
<td align=center background="images/background.gif">
<b><font color="#ffffff">厂 部 资 料</font></b>
</td>
</tr>


<tr>
<td valign=top>
<table border="3" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#0055E6" width="350" id="AutoNumber1">
<tr>
<td width="60" align="right" bgcolor="#EFE6DE">员工姓名：</td>
<td width="420" bgcolor="#EFE6DE"><input type="text" name="Photo_Name" size="15">
<font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#D3CFC2">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
			<td width="420" bgcolor="#D3CFC2">
			<select name="Photo_author" size="1">
        <option value="男" selected>男</option>
        <option value="女">女</option>
			</select>
			</td>
</tr>

<tr>
<td width="60" align="right" bgcolor="#EFE6DE">年&nbsp;&nbsp;&nbsp;&nbsp;龄：</td>
<td width="420" bgcolor="#EFE6DE"><input type="text" name="nianling" size="30">
</td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#D3CFC2">民&nbsp;&nbsp;&nbsp;&nbsp;族：</td>
<td width="420" bgcolor="#D3CFC2"><input type="text" name="minzu" size="30" value="汉族">
</td>
</tr>

<tr>
<td width="60" align="right" bgcolor="#EFE6DE">籍&nbsp;&nbsp;&nbsp;&nbsp;贯：</td>
<td width="420" bgcolor="#EFE6DE"><input type="text" name="jiguan" size="30" value="">
</td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#D3CFC2">住&nbsp;&nbsp;&nbsp;&nbsp;址：</td>
<td width="420" bgcolor="#D3CFC2"><input type="text" name="zhuzhi" size="40" value="">
</td>
</tr>


<tr>
<td width="60" align="right" bgcolor="#EFE6DE">婚&nbsp;&nbsp;&nbsp;&nbsp;况：</td>
<td width="420" bgcolor="#EFE6DE">
			<select name="hunfou" size="1">
        <option value="已婚" selected>已婚</option>
        <option value="未婚">未婚</option>
			</select>
</td>
</tr>

<tr>
<td width="60" align="right" bgcolor="#D3CFC2">学&nbsp;&nbsp;&nbsp;&nbsp;历：</td>

			<td width="420" bgcolor="#D3CFC2">
			<select name="Photo_Address" size="1">
        <option value="小学" selected>小学</option>
        <option value="初中">初中</option>
        <option value="高中">高中</option>
        <option value="中专">中专</option>
        <option value="大专">大专</option>
        <option value="本科">本科</option>
			</select>
			</td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#EFE6DE">身份证号：</td>
<td width="420" bgcolor="#EFE6DE"><input type="text" name="shenfenzhenghao" size="30">
</td>
</tr>

<tr>
<td width="60" align="right" height="32">&nbsp;</td>
<td width="420">&nbsp;
</td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#EFE6DE">备&nbsp;&nbsp;&nbsp;&nbsp;注：</td>
<td width="420" bgcolor="#EFE6DE"><textarea rows="9" name="Photo_Content" cols="30"></textarea></td>
</tr>


</table>
</td>
<td valign=top>
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#0055E6" width="350" id="AutoNumber1">
<tr>
<td width="60" align="right" bgcolor="#EFE6DE">员工编号：</td>

<td width="420" bgcolor="#EFE6DE"><input type="text" name="bianhao" size="30">
</td>
</tr>

<tr>
<td width="60" align="right" bgcolor="#D3CFC2">所属部门：</td>
<td width="420" bgcolor="#D3CFC2">
<select name="ClassName" id="select2">
<option value="" selected>-&gt;请选择</option>
<%
sql2="select * from [department]"
set rs2=conn.execute(sql2)
if not rs2.bof then
Do While Not rs2.eof %>
<option value="<%=rs2(2)%>"><%=rs2(2)%></option>
<% i=i+1
rs2.MoveNext
Loop
end if
rs2.close
%></select><font color="#FF0000"> *</font></td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#EFE6DE">入厂时间：</td>

<td width="420" bgcolor="#EFE6DE">

<input name="Photo_Time" onFocus="show_cele_date(Photo_Time,'','',Photo_Time)" type="text"  size="10"  value='<% response.write date()%>'>
</td>
</tr>

<tr>
<td width="60" align="right" bgcolor="#D3CFC2">联系电话：</td>
<td width="420" bgcolor="#D3CFC2"><input type="text" name="dianhua" size="30">
</td>
</tr>


<tr>
<td width="60" align="right" bgcolor="#EFE6DE">职&nbsp;&nbsp;&nbsp;&nbsp;务：</td>
<td width="420" bgcolor="#EFE6DE"><input type="text" name="zhiwu" size="30" value="">
</td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#D3CFC2">工&nbsp;&nbsp;&nbsp;&nbsp;种：</td>
<td width="420" bgcolor="#D3CFC2"><input type="text" name="gongzhong" size="30" value="">
</td>
</tr>

<tr>
<td width="60" align="right" bgcolor="#EFE6DE">押&nbsp;&nbsp;&nbsp;&nbsp;金：</td>
<td width="420" bgcolor="#EFE6DE"><input type="text" name="yajin" size="30" value="">
</td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#D3CFC2">保&nbsp;&nbsp;&nbsp;&nbsp;险：</td>
<td width="420" bgcolor="#D3CFC2"><input type="text" name="baoxian" size="30" value="">
</td>
</tr>
<tr>
<td width="60" align="right" bgcolor="#EFE6DE">档&nbsp;&nbsp;&nbsp;&nbsp;案：</td>
<td width="420" bgcolor="#EFE6DE">
			<select name="dangan" size="1">
        <option value="有" selected>有</option>
        <option value="无">无</option>
			</select>
</td>
</tr>

<tr>
<td width="60" align="right" bgcolor="#D3CFC2">离厂时间：</td>

<td width="420" bgcolor="#D3CFC2">

<input name="Photo_Time1" onFocus="show_cele_date(Photo_Time1,'','',Photo_Time1)" type="text"  size="10"  value='<% response.write date()%>'>
</td>
</tr>




<tr>
<td width="60" align="right" bgcolor="#EFE6DE">工作变动情况：</td>
<td width="420" bgcolor="#EFE6DE"><textarea rows="9" name="qingkuang" cols="30"></textarea></td>
</tr>


</table>
</td>
</tr>
</table>



<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="700" id="AutoNumber1">
<tr>
<td width="60" align="right" height=30>　</td>
<td width="420"><input type="submit" value="添加照片" name="B1">注：</font>添加完基本信息后再上传照片</td>
</tr>
<tr>
<td width="60" align="right">　</td>
<td width="420"><font color="#FF0000"></td>
</tr><input type="hidden" name="action" value="1">
</table>
</td>
</tr>

</form>
</table>

<P>
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <TBODY>
  <TR vAlign=center>
    <TD align=middle width="100%"><!--#include file="footer.htm"--></TD>
  </TR></TBODY></TABLE></P>
</body>
</html>