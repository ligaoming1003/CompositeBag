<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
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

	if (myform.fenleinumber.value == "")
	{
		alert("����Ų���Ϊ�գ�");
		myform.fenleinumber.focus();
		return (false);
	}
	if (myform.fenleiname.value == "")
	{
		alert("������Ʋ���Ϊ�գ�");
		myform.fenleiname.focus();
		return (false);
	}
}


function SelectAll() {
	for (var i=0;i<document.selform.selBigClass.length;i++) {
		var e=document.selform.selBigClass[i];
		e.checked=!e.checked;
	}
}
//-->
</script>
</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0>       
<!--#include file="top.asp"--><% dim sql,rs

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
  case else
       isEdit=False
       call myform(isEdit) 
  end select

  
  sub SaveAdd   
set rs=server.createobject("adodb.recordset") 
sql="select * from fenlei"
rs.open sql,conn,1,3
rs.addnew
rs("fenleinumber") = trim(request.Form("fenleinumber"))
rs("fenleiname") = trim(request.Form("fenleiname"))
rs("content") = trim(request.Form("content"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('��ӳɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=fenlei.asp"">"
response.end
end sub

 
sub SaveModify   
set rs=server.createobject("adodb.recordset") 
sql="select * from fenlei where id="&request.Form("id")
rs.open sql,conn,1,3
rs("fenleiname") = trim(request.Form("fenleiname"))
rs("fenleinumber") = trim(request.Form("fenleinumber"))
rs("content") = trim(request.Form("content"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('�޸ĳɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=fenlei.asp"">"
response.end
rs.close
set rs=nothing
end sub   
 
  sub delCate()
        conn.execute("delete from fenlei where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('ɾ���ɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=fenlei.asp"">"
response.end
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=600 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin"  then%>
	
	<%
	    set rs=server.createobject("adodb.recordset")
	   if isedit then
		   rs.open "select * from fenlei where id=" & request("id"),conn,1,1
%>

<TABLE width=600 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�� �� �� �� �� ��</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="600" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��� �� �� �� ��</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="600" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
          <tr class="but">
			<td width="200" height="18" align="center" class="but" onMouseDown="this.fenleiname='tddown'" onMouseUp="this.fenleiname='but'" onMouseOut="this.fenleiname='but'">�����</td>
            <td width="200" height="18" align="center" class="but" onMouseDown="this.fenleiname='tddown'" onMouseUp="this.fenleiname='but'" onMouseOut="this.fenleiname='but'">�������</td> 


			<td width="200" height="18" align="center" class="but" onMouseDown="this.fenleiname='tddown'" onMouseUp="this.fenleiname='but'" onMouseOut="this.fenleiname='but'" colspan="2">��ע</td>
			
          </tr>
		  <form name="reg" method="post" action="fenlei.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"><input name="fenleinumber" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("fenleinumber")
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"><input name="fenleiname" type="text"  size="25" value='<% if isedit then
		                                                         response.write rs("fenleiname")
																  end if %>'></td>
  
 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"><input name="content" type="text"  size="30" value='<% if isedit then
		                                                         response.write rs("content")
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"><input type="submit" class="button" value="<% if isedit then
		                                                         response.write "�༭"
																 else
																 response.write "���"
																  end if %>"> </td>
          </tr>
		  </form>
		  
           </table>
        <%end if%>

</td>
</tr>
</table>
<br><br>
<TABLE width=600 align=center border="0" cellspacing="0" cellpadding="0">
<tr>
<td >


			<form action="fenlei.asp" method="post" name="selform" >
  <table width="600"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="600" align="center" valign="top">
	        <table width="600" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�� �� �� �� �� ��</font></b>

</td>
</tr>
</table>
	  
         <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="100" height="18" align="center" class="but01" onMouseDown="this.fenleiname='tddown'" onMouseUp="this.fenleiname='but'" onMouseOut="this.fenleiname='but01'">�����</td>
            <td width="200" height="18" align="center" class="but01" onMouseDown="this.fenleiname='tddown'" onMouseUp="this.fenleiname='but'" onMouseOut="this.fenleiname='but01'">�������</td>
			<td width="250" height="18" align="center" class="but01" onMouseDown="this.fenleiname='tddown'" onMouseUp="this.fenleiname='but'" onMouseOut="this.fenleiname='but01'">�� ע</td>
            <td width="50" height="18" align="center" class="but01" onMouseDown="this.fenleiname='tddown'" onMouseUp="this.fenleiname='but'" onMouseOut="this.fenleiname='but01'">�� ��</td>
          </tr>
		 <%
set rs=server.CreateObject("ADODB.RecordSet") 
rs.Source="select* from fenlei"
rs.Open rs.Source,conn,1,1
if not rs.EOF then
do while not rs.eof
%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("fenleinumber")%></td>
<td height="18" align="center"><a href="fenlei.asp?id=<%=rs("ID")%>&action=edit"><%=rs("fenleiname")%></a><%if rs("fenleiname")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("content")%><%if rs("content")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
<tr><td align="right" colspan="10" height="22"> <%if oskey="supper" then%>
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> ѡ��/��ѡ
              <input onClick="{if(confirm('�˲�����ɾ������Ϣ��\n\nȷ��Ҫִ�д��������')){this.document.selform.submit();return true;}return false;}" type=submit value=ɾ�� name=action2> 
              <input type="Hidden" name="action" value='del'><%end if%></td></tr>
        </table>

	  
</td>
    </tr>

  </table>

</form>
			</TD>
              </TR>
          </table>
		          

  </TBODY>
</TABLE>
<%end sub%>
<P>
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <TBODY>
  <TR vAlign=center>
    <TD align=middle width="100%"><!--#include file="footer.htm"--></TD>
  </TR></TBODY></TABLE></P></BODY></HTML>


