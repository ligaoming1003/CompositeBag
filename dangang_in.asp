<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>

<SCRIPT language=javascript  src="js/selectcity.js"></script>
<script language="JavaScript" src="js/validate.js" type="text/JavaScript"></script>

<SCRIPT language=javascript src="css/init.js"></SCRIPT>

<style type="text/css">
<!--
td {  font-family: "����"; font-size: 9pt}
body {  font-family: "����"; font-size: 9pt}
select {  font-family: "����"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "����"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "����"; font-size: 9pt} 
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


<!--media=print ������Կ����ڴ�ӡʱ��Ч--> 
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

  
 
  
 

  %> <% sub myform(isEdit) %> 	 

 

	
	<%
	    set rs=server.createobject("adodb.recordset")
	   if isedit then
		   rs.open "select * from dangang where id=" & request("id"),conn,1,1
id=rs("id")
	   end if %>


        <table width="420" border="0" cellpadding="0" style="border-collapse: collapse" >
        <tr>
 <td colspan="2" align=center valign=middle height=147>��<p>��</p>
	<p><u><font size="5"><b>��������</b></font></u></td>
</tr>
<tr>
<td colspan="2" height=23 align=left><font size=2>&nbsp;</font><font size=3>&nbsp; �ͻ�:<b><%=rs("kehu")%></b><%=rs("tel")%>&nbsp;
&nbsp;&nbsp; </font><font size=4 color="#FF0000"><b>��:
</b>[</font><font size=2 color="#FF0000"></font><font size=4 color="#FF0000"><b><%=rs("id")%></b>]</font></td>
</tr>

<tr>
<td>
<table width="420" border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr>
<td height="40" align=center width="18%" ><b><font size="3">�� ��</font></b></td>

<td height="40" align=center width="65%"  ><b><font size="3">Ʒ��</font></b></td>

<td height="40" align=center width="17%"><b><font size="3">�� ��</font></b></td>

</tr>


<tr>

<td height="40" align=center><font size="3"><%=rs("fenlei")%></font></td>

<td height="40" align=center><font size="3"><%=rs("pinming")%></font></td>

<td height="40" align=center><font size="3"><%=rs("chang")%>*<%=rs("kuang")%></font></td>

</tr>





<tr>

<td height="40" align=center><b><font size="3">��ӡ</font></b></td>

<td height="40" align=center colspan="2">
<p align="left"><font size="3">&nbsp;<br>
<b>��Ĥ</b>��<%if rs("name1")="0" or rs("name1")="" then %>&nbsp;<%else%><%=rs("name1")%>&nbsp;<%=rs("nianbang")%>*<%=rs("cunshu")%>&nbsp;&nbsp;��ţ�<%=rs("pihao")%><%end if%></font><p align="left">
<font size="3"><b>˵��</b>��<%=rs("chati")%><br>
��</font></td>

</tr>





<tr>

<td height="40" align=center><b><font size="3">����</font></b></td>

<td height="40" align=center colspan="2">
<p align="left"><font size="3"><br>
<%if rs("name2")="0" or rs("name2")="" then %><%else%><b>��Ĥ</b>��<%=rs("name2")%>&nbsp;<%=rs("nianbang2")%>*<%=rs("cunshu2")%>&nbsp;&nbsp;<%end if%>
<%if rs("name3")="0" or rs("name3")="" then %><%else%>&nbsp;</font><p align="left">
<font size="3"><b> ԧĤ</b>��<%=rs("name3")%>&nbsp;<%=rs("nianbang3")%>*<%=rs("cunshu3")%>&nbsp;&nbsp;<%end if%><b><br>
<br>
 ��Ĥ</b>��<%if rs("name4")="0" or rs("name4")="" then %>&nbsp;<%else%><%=rs("name4")%>&nbsp;<%=rs("nianbang4")%>*<%=rs("cunshu4")%><%end if%>&nbsp;&nbsp;<p align="left">
<b>˵��</b>��<%=rs("chati4")%><br>
��</font></td>

</tr>








<tr>

<td height="80" align=center><b><font size="3">��ƷҪ��</font></b></td>

<td height="80" align=center colspan="2">
<p align="left"><font size="3"> 
<b>˵��</b>��<%=rs("chati2")%></font></td>

</tr>





<tr>

<td height="40" align=center><font size="3"><b>��װҪ��</b></font></td>

<td height="40" align=center colspan="2">
<p align="left"><b><font size="3">˵��</font></b><font size="3">��<%=rs("chati3")%></font></td>

</tr>





</table>







		          
<%end sub%>


</td>
</tr>
<tr>
 <td colspan="8" align=right valign=bottom height=30>
	<p align="left"><br><font size=2><b>���ܾ���:<u><%=name%></u></b></font></td>
</tr>

<table><tr><td height="30" colspan="8" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=��ӡ onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=ֱ�Ӵ�ӡ onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=ҳ������ onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=��ӡԤ�� onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>

				
				</table>  		
</BODY></HTML>
