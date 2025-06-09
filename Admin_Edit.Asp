<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<%
Function SafeRequest(ParaName,ParaType)
	Dim ParaValue
	ParaValue=Request(ParaName)
	If ParaType=1 Then
		If Not isNumeric(ParaValue) Then
			response.write"<body  bgcolor=#E3E3E3><a href='javascript:history.go(-1)'>你非法操作了！返回！</a>"
		End If
	Else
		ParaValue=replace(ParaValue,"'","''")
	End if
	SafeRequest=ParaValue
End function
ID=SafeRequest("id",1)
%>
<html>
<head>
<title>∷嘉美管理系统:.</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
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
ID=SafeRequest("id",1)
Photo_Name=trim(Request("Photo_Name"))
ClassName=Request.form("ClassName")
if Photo_Name="" or id="" or ClassName="" then
   Response.Write "<script>alert(""系统异常！请勿非法操作！"");location.href=""javascript:history.go(-1)"";</script>"
   response.end()
end if

Photo_author=trim(Request("Photo_author"))
Photo_Time=trim(Request("Photo_Time"))
Photo_Address=trim(Request("Photo_Address"))
shenfenzhenghao=trim(Request("shenfenzhenghao"))
dianhua=trim(Request("dianhua"))
nianling=trim(Request("nianling"))
minzu=trim(Request("minzu"))

bianhao=trim(Request("bianhao"))
jiguan=trim(Request("jiguan"))
zhuzhi=trim(Request("zhuzhi"))

zhiwu=trim(Request("zhiwu"))
gongzhong=trim(Request("gongzhong"))
hunfou=trim(Request("hunfou"))
yajin=trim(Request("yajin"))
baoxian=trim(Request("baoxian"))
dangan=trim(Request("dangan"))

qingkuang=HTMLEncode(Request("qingkuang"))
Photo_Content=HTMLEncode(Request("Photo_Content"))


'------------------检查结束------------------------------------

'-----------------将内容写入数据库--------------------------------------

Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from ImgData where id="&id
rs.open sql,conn,1,3

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

rs("bianhao")=bianhao
rs("jiguan")=jiguan
rs("zhuzhi")=zhuzhi


rs("zhiwu")=zhiwu
rs("gongzhong")=gongzhong
rs("hunfou")=hunfou
rs("yajin")=yajin
rs("baoxian")=baoxian
rs("dangan")=dangan

rs("Photo_Content")=Photo_Content
rs("qingkuang")=qingkuang

rs.Update
rs.Close
Set rs=Nothing
'-----------------写入数据库结束----------------------------------------
Response.Write "<script>alert(""图片修改成功！"");location.href=""Admin_All.asp"";</script>"
response.end()
end if




'读取信息
sql="select * from [ImgData] where id="&id
set rs = createobject("ADODB.recordset")
rs.open sql,conn
if not rs.eof then

   Photo_Name=rs("Photo_Name")
   ClassName=rs("ClassName")

Photo_author=rs("Photo_author")
Photo_Time=rs("Photo_Time")
Photo_Time1=rs("Photo_Time1")
Photo_Address=rs("Photo_Address")
shenfenzhenghao=rs("shenfenzhenghao")
dianhua=rs("dianhua")
nianling=rs("nianling")
minzu=rs("minzu")

bianhao=rs("bianhao")
zhuzhi=rs("zhuzhi")
jiguan=rs("jiguan")

zhiwu=rs("zhiwu")
gongzhong=rs("gongzhong")
hunfou=rs("hunfou")
yajin=rs("yajin")
baoxian=rs("baoxian")
dangan=rs("dangan")


qingkuang=rs("qingkuang")
Photo_Content=rs("Photo_Content")

end if
rs.close
set rs=nothing
%> 
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="522" id="AutoNumber1">

<form ACTION="Admin_Edit.asp" METHOD="POST" onSubmit="return checkform(this)" name="form">
  <tr>
    <td width="66" align="right">员工姓名：</td>
    <td width="436" colspan="2">
    <input type="text" name="Photo_Name" size="31" value="<%=Photo_Name%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">员工编号：</td>
    <td width="436" colspan="2">
    <input type="text" name="bianhao" size="31" value="<%=bianhao%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">所属部门：</td>
    <td width="436" colspan="2">
<select name="ClassName" id="select2">
<%
sql2="select * from [department]"
set rs2=conn.execute(sql2)
if not rs2.bof then
Do While Not rs2.eof %>
<option value="<%=rs2("depname")%>" <%if rs2("depname")=ClassName then response.write"Selected" end if%>><%=rs2("depname")%></option>
<% i=i+1
rs2.MoveNext
Loop
end if
rs2.close
%></select>
</td>
  </tr>
  <tr>
    <td width="66" align="right">性&nbsp;&nbsp;&nbsp;&nbsp;别：</td>
    <td width="436" colspan="2">


			<select name="Photo_author" size="1"  value="<%=Photo_author%>">
        <option value="男" <%if rs2("Photo_author")=Photo_author then response.write"Selected" end if%>>男</option>
        <option value="女" >女</option>
			</select>





</td>
  </tr>
  <tr>
    <td width="66" align="right">学&nbsp;&nbsp;&nbsp;&nbsp;历：</td>
    <td width="436" colspan="2">
    <input type="text" name="Photo_Address" size="20" value="<%=Photo_Address%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">入厂时间：</td>
    <td width="436" colspan="2">
    <input type="text" name="Photo_Time" size="20" value="<%=Photo_Time%>"></td>
  </tr>

  <tr>
    <td width="66" align="right">年&nbsp;&nbsp;&nbsp;&nbsp;龄：</td>
    <td width="436" colspan="2">
    <input type="text" name="nianling" size="20" value="<%=nianling%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">民&nbsp;&nbsp;&nbsp;&nbsp;族：</td>
    <td width="436" colspan="2">
    <input type="text" name="minzu" size="20" value="<%=minzu%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">身份证号：</td>
    <td width="436" colspan="2">
    <input type="text" name="shenfenzhenghao" size="20" value="<%=shenfenzhenghao%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">联系电话：</td>
    <td width="436" colspan="2">
    <input type="text" name="dianhua" size="20" value="<%=dianhua%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">家庭住址：</td>
    <td width="436" colspan="2">
    <input type="text" name="zhuzhi" size="31" value="<%=zhuzhi%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">籍&nbsp;&nbsp;&nbsp;&nbsp;贯：</td>
    <td width="436" colspan="2">
    <input type="text" name="jiguan" size="31" value="<%=jiguan%>"></td>
  </tr>

<tr>
<td width="60" align="right">职&nbsp;&nbsp;&nbsp;&nbsp;务：</td>
<td width="420"><input type="text" name="zhiwu" size="30"  value="<%=zhiwu%>">
</td>
</tr>
<tr>
<td width="60" align="right">工&nbsp;&nbsp;&nbsp;&nbsp;种：</td>
<td width="420"><input type="text" name="gongzhong" size="30"  value="<%=gongzhong%>">
</td>
</tr>
<tr>
<td width="60" align="right">婚&nbsp;&nbsp;&nbsp;&nbsp;否：</td>
<td width="420">
<input type="text" name="hunfou" size="20" value="<%=hunfou%>">

</td>
</tr>
<tr>
<td width="60" align="right">押&nbsp;&nbsp;&nbsp;&nbsp;金：</td>
<td width="420">
<input type="text" name="yajin" size="20" value="<%=yajin%>">
</td>
</tr>
<tr>
<td width="60" align="right">保&nbsp;&nbsp;&nbsp;&nbsp;险：</td>
<td width="420">
<input type="text" name="baoxian" size="20" value="<%=baoxian%>">
</td>
</tr>
<tr>
<td width="60" align="right">档&nbsp;&nbsp;&nbsp;&nbsp;案：</td>
<td width="420">
<input type="text" name="dangan" size="20" value="<%=dangan%>">
</td>
</tr>
  <tr>
    <td width="66" align="right">离厂时间：</td>
    <td width="436" colspan="2">
    <input type="text" name="Photo_Time1" size="20" value="<%=Photo_Time1%>"></td>
  </tr>
  <tr>
    <td width="66" align="right">更换照片：</td>
    <td width="253">
    <iframe name="I1" src="Admin_Upload.asp?ID=<%=id%>" marginwidth="1" marginheight="1" height="90" width="242" scrolling="no" border="0" frameborder="0"></iframe></td>
    <td width="173"><a target=blank href="view.Asp?id=<%=id%>" title="浏览原详细资料及照片">
                   查看原照片</a></td>
  </tr>
  <tr>
    <td width="66" align="right">工作变动情况：</td>
    <td width="436" colspan="2">
    <textarea rows="9" name="qingkuang" cols="52"><%=qingkuang%></textarea></td>
  </tr>
  <tr>
    <td width="66" align="right">备&nbsp;&nbsp;&nbsp;&nbsp;注：</td>
    <td width="436" colspan="2">
    <textarea rows="9" name="Photo_Content" cols="52"><%=Photo_Content%></textarea></td>
  </tr>
  <tr>
    <td width="66" align="right">　</td>
    <td width="436" colspan="2"><input type="submit" value="修改资料" name="B1"></td>
  </tr>
  <tr>
    <td width="66" align="right">　</td>
    <td width="436" colspan="2">　</td>
  </tr>
<input type="hidden" name="action" value="1">
  <input type="hidden" name="id" value="<%=id%>"></form>
</table>

<P>
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <TBODY>
  <TR vAlign=center>
    <TD align=middle width="100%"><!--#include file="footer.htm"--></TD>
  </TR></TBODY></TABLE></P>
</body>
</html>