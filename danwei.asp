<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>∷嘉美管理系统∷</title>
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
<SCRIPT LANGUAGE=javascript>
<!--
function Juge(myform)
{

	if (myform.danweiname.value == "")
	{
		alert("名称不能为空！");
		myform.danweiname.focus();
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
sql="select * from danwei"
rs.open sql,conn,1,3
rs.addnew
rs("danweiname") = trim(request.Form("danweiname"))
rs("danweinumber") = trim(request.Form("danweinumber"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('添加成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=danwei.asp"">"
response.end
rs.close
set rs=nothing
end sub

 
sub SaveModify   
set rs=server.createobject("adodb.recordset") 
sql="select * from danwei where id="&request.Form("id")
rs.open sql,conn,1,3
rs("danweinumber") = trim(request.Form("danweinumber"))
rs("danweiname") = trim(request.Form("danweiname"))
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=danwei.asp"">"
response.end
rs.close
set rs=nothing
end sub   
 
  sub delCate()
        conn.execute("delete from danwei where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=danwei.asp"">"
response.end
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=600 align=center border="0" cellspacing="0" bordercolor="#D6D3CE">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"><%if oskey="supper" or oskey="admin"  then%> 
	<fieldset style="width:98%"><legend>
	<%
	    set rs=server.createobject("adodb.recordset")
	   if isedit then
		   rs.open "select * from danwei where id=" & request("id"),conn,1,1
	       response.write "编 辑 单 位 信 息"
	   else
	       response.write "添 加 单 位 信 息"
	   end if %></legend>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
          <tr class="but"> 
            <td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">部门编号</td>
			<td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">部门名称</td>

			<td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2">&nbsp;</td>
			
          </tr>
		  <form name="myform" method="post" action="danwei.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
           <td align="center"><input name="danweinumber" type="text"  size="25" value='<% if isedit then
		                                                         response.write rs("danweinumber")
																  end if %>'></td>        
    <td height="18" align="center"><input name="danweiname" type="text"  size="30"  value='<% if isedit then
		                                                         response.write rs("danweiname")
																  end if %>'></td>
 
			<td width="148" align="center"><input type="submit" class="button" value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>"> </td>
          </tr>
		  </form>
		  
           </table>
        </fieldset><br><br><br>
<fieldset style="width:98%">
<legend>单 位 信 息 列 表</legend>
			<form action="danwei.asp" method="post" name="selform" >
  <table width="100%"  border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8" align="center">
    <tr> 
      
      <td width="100%" align="center" valign="top">
	
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
          <tr class="but">
            <td width="200" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">部门编号</td> 
			<td width="200" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">部门名称</td>

            <td width="200" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">操 作</td>
          </tr>
		 <%
set rs=server.CreateObject("ADODB.RecordSet") 
rs.Source="select* from danwei"
rs.Open rs.Source,conn,1,1
if not rs.EOF then
do while not rs.eof
%>
<tr>
<td height="18" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><%=rs("danweinumber")%><%if rs("danweinumber")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><a href="danwei.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("danweiname")%></a></td>

<td height="18" align="center" class="but1" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but1'"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.movenext
loop
end if
rs.close
set rs=nothing
%>
<tr><td align="right" colspan="10" height="22"><%if oskey="supper" or oskey="admin"  then%> 
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input onClick="{if(confirm('此操作将删除该信息！\n\n确定要执行此项操作吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
              <input type="Hidden" name="action" value='del'><%end if%></td></tr>
        </table>

	  <%end if%>
</td>
    </tr>

  </table></fieldset>
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


