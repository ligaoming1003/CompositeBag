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




<% dim sql,rs,hedu

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
		   rs.open "select * from dindang where id=" & request("id"),conn,1,1
id=rs("id")
	   end if %>


        <table width="450" border="0" cellpadding="0" style="border-collapse: collapse" >
        <tr>
 <td align=center valign=middle height=147>��<p>��</p>
	<p align="left"><b><font size="7">&nbsp;&nbsp; <font face="����"> <u><%=rs("fenlei")%></u></font></font></b></td>
</tr>

<tr>
<td>
<table width="450" border="0" cellpadding="0" cellspacing="0">
<tr>
<td height="50" align=center width="22%" ><b><font size="5">������</font></b></td>

<td height="50" align=left width="78%" colspan="2" ><b><font size="4"><%=rs("kehu")%></font></b></td>

</tr>





<tr>
<td height="46" align=center width="22%" ><b><font size="5">Ʒ����</font></b></td>

<td height="46" align=left width="78%" colspan="2" ><b><font size="4"><%=rs("pinming")%></font></b></td>

</tr>





<tr>

<td height="46" align=center width="22%" ><b><font size="5">���ϣ�</font></b></td>

<td height="46" align=center width="59%" >��<p align="left"><b><font size="4"><%=rs("name1")%><%if rs("name2")="0" or rs("name2")="" then %><%else%>/<%=rs("name2")%><%end if%><%if rs("name3")="0" or rs("name3")="" then %><%else%>/<%=rs("name3")%><%end if%><%if rs("name4")="0" or rs("name4")="" then %><%else%>/<%=rs("name4")%><%end if%>
&nbsp;<%=rs("nianbang")%>*<%=rs("cunshu")+rs("cunshu2")+rs("cunshu3")+rs("cunshu4")%>S</font></b></p>
<p>��</td>

<td height="46" align=center width="19%" ><%if rs("doupian")="" then%>&nbsp;<%else%><img src="<%=rs("doupian")%>"width="85"><%end if%></td>

</tr>





<tr>
<td height="46" align=center width="22%" ><b><font size="5">���</font></b></td>

<td height="46" align=center width="78%" colspan="2" >
<p align="left"><b><font size="4"><%=rs("chang")%>*<%=rs("kuang")%>cm</font></b></td>

</tr>





<tr>
<td height="46" align=center width="22%" ><b><font size="5">Ҫ��</font></b></td>

<td height="46" align=center width="78%" colspan="2" ><p align="left"><b><font size="4"><%=rs("chati2")%>��</font></b></p>
</td>

</tr>





<tr>
<td height="46" align=center width="22%" ><b><font size="5">ʱ�䣺</font></b>��</td>

<td height="46" align=center width="78%" colspan="2" >��</td>

</tr>





<tr>
<td height="46" align=center width="22%" ><b><font size="5">��װ��</font></b>��</td>

<td height="46" align=center width="78%" colspan="2" ><b><font size="4">
<p align="left"><%=rs("chati3")%></font></b>����</td>

</tr>





</table>







		          
<%end sub%>


</td>
</tr>
<tr>
 <td align=right valign=bottom height=30>
	<p><br><font size=2><b>���ܾ���:<u><%=name%></u>&nbsp;&nbsp;</b></font></td>
</tr>

				</table> 
<table><tr><td height="30" colspan="8" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=��ӡ onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=ֱ�Ӵ�ӡ onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=ҳ������ onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=��ӡԤ�� onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>

				
				</table>  		
</BODY></HTML>