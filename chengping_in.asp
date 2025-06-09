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




<% dim sql,rs

  select case request("action")
  case "add"
       call SaveAdd()
  case "modify"
       call SaveModify()
  case "del"
       call delCate()
  case "edit"
       isEdit=True
       call myform(isEdit)
  case "view"
       isEdit=True
       call myform(isview)
  case else
       isEdit=False
       call myform(isEdit) 
  end select

  
 
sub SaveModify   
set rs=server.createobject("adodb.recordset") 
sql="select * from chengping_in_store where id="&request.Form("id")
rs.open sql,conn,1,3
rs("id") = trim(request.Form("id"))
rs("uptime") = trim(request.Form("uptime"))
rs("fenlei") = trim(request.Form("fenlei"))
rs("pinming") = trim(request.Form("pinming"))
rs("guige") = trim(request.Form("guige"))
rs("Unit") = trim(request.Form("Unit"))
rs("Supplier") = trim(request.Form("Supplier"))
rs("use_price") = trim(request.Form("use_price"))
rs("use_num") = trim(request.Form("use_num"))
rs("use_Amount") = trim(request.Form("use_Amount"))
rs("gonghuodanwei") = trim(request.Form("gonghuodanwei"))




rs("content") = trim(request.Form("content"))
rs.update

rs.close
set rs=nothing
end sub   
 

  %> <% sub myform(isEdit) %> 	 

 
	 
	
	<%
	    set rs=server.createobject("adodb.recordset")
	   if isedit then
		   rs.open "select * from chengping_in_store where id=" & request("id"),conn,1,1
id=rs("id")
	   end if %>


        <table width="600" border="0" cellpadding="0" style="border-collapse: collapse" >
        <tr>
 <td colspan="2" align=center valign=middle height=57><u><font size="5"><b>成品入库单</b></font></u></td>
</tr>
<tr>
<td colspan="2" height=23 align=left><font size=2>&nbsp;&nbsp; 日期：<%=date()%>&nbsp;&nbsp;<b>车间:</b><%=rs("use_dep")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><b><font size=4 color="#FF0000">№:<%
Response.Write(year(date()) & "" & month(date()) & "" & day(date()) & "" & "" & mid("",weekday(date()),1))
%><%=rs("id")%></font></b></td>
</tr>

<tr>
<td>
<table width="600" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<td height="40" align=center width="20%" ><b><font size="3">品名</font></b></td>

<td height="40" align=center width="20%"  ><b><font size="3">规格</font></b></td>

<td height="40" align=center width="40%"><b><font size="3">数量</font></b></td>

<td height="40" align=center width="20%"  ><b><font size="3">单位</font></b></td>

</tr>


<tr>

<td height="40" align=center><%=rs("pinming")%></td>

<td height="40" align=center><%=rs("guige")%>*<%=rs("kuang")%></td>

<td height="40" align=center><%=Formatnumber(rs("use_num"),2,-1,0,0)%></td>

<td height="40" align=center><%=rs("unit")%></td>
</tr>





</table>






		          
<%end sub%>


</td>
</tr>
<tr>
 <td colspan="8" align=right valign=bottom height=30><br><font size=2><b>主管经理：_______________&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;经 办 人：</b><%=name%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
</tr>

				
	<table><tr><td height="30" colspan="8" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=打印 onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=直接打印 onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=页面设置 onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=打印预览 onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>

				
				</table>     
		
</BODY></HTML>

