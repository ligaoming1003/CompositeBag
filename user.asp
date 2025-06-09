<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<!--#include file="md5.asp" -->
<%
if oskey<>"supper" then
response.Write "<script language=javascript>alert('您无权访问！');window.close();</script>"
response.end
end if
%>
<html>
<head>
<title>∷嘉美管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">

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
<SCRIPT LANGUAGE=javascript>
<!--
function Juge(powersearch)
{

	if (powersearch.username.value == "")
	{
		alert("用户名称不能为空！");
		powersearch.username.focus();
		return (false);
	}
	if (powersearch.pwd.value == "")
	{
		alert("用户密码不能为空！");
		powersearch.pwd.focus();
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
<!--#include file="top.asp"-->
<%
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
sql="select * from userinfo"
rs.open sql,conn,1,3
rs.addnew
rs("name") = trim(request.Form("username"))
rs("fullname") = trim(request.Form("fullname"))
rs("depname") = trim(request.Form("depname"))
rs("oskey") = trim(request.Form("oskey"))
rs("pwd") = md5(request.Form("pwd"))
rs.update
response.Write "<script language=javascript>alert('添加成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=user.asp"">"
response.end
rs.close
set rs=nothing
end sub

 
sub SaveModify
passwd=request.form("pwd")
passwd1=md5(trim(request.form("pwd")))   
set rs=server.createobject("adodb.recordset") 
sql="select * from userinfo where userid="&request.Form("userid")
rs.open sql,conn,1,3
rs("name") = trim(request.Form("username"))
rs("fullname") = trim(request.Form("fullname"))
rs("depname") = trim(request.Form("depname"))
rs("oskey") = trim(request.Form("oskey"))
if passwd<>rs("pwd") then
rs("pwd")=passwd1
end if
rs.update
response.Write "<script language=javascript>alert('修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=user.asp"">"
response.end
rs.close
set rs=nothing
end sub   
 
  sub delCate()
        conn.execute("delete from userinfo where userid in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=user.asp"">"
response.end
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> 
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from userinfo where userid=" & request("userid"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编 辑 权 限 信 息</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添 加 权 限 信 息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">用户名</td>
			<td width="5%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">密 码</td>
			<td width="15%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">登陆次数</td>
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">上次登陆</td>
			<td width="15%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">上次时间</td>
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">部 门</td>
			<td width="18%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">权 限</td>
			
          </tr>
		  <form name="powersearch" method="post" action="user.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="userid" value="<%=rs("userid")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="username"  type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("name")
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="pwd" onmouseover="this.focus()"  onfocus="this.select()" type="password" size="15"  value='<% if isedit then
		                                                         response.write rs("pwd")
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"><input readonly name="logins" type="text"  size="10" value='<% if isedit then
		                                                         response.write rs("logins")
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input readonly name="lastlogin" type="text"  size="20" value='<% if isedit then
		                                                         response.write rs("lastlogin")
																  end if %>'>
			</td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"><input readonly name="lastip" type="text"  size="20" value='<% if isedit then
		                                                         response.write rs("lastip")
																  end if %>'>
			
			</td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="depname"  size="1" style='width:80'>
			<%		
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from department",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if trim(rs3("depname"))=trim(rs("depname")) then
			%> selected="selected"<%end if%><%end if%>><%=rs3("depname")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>
			<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name="oskey"  size="1" style="width:104; height:19">
			<option value="supper" <% if isedit then
			if trim(rs("oskey"))="supper" then
			%> selected="selected"<%end if%><%end if%>>超级管理员</option>

			<option value="admin" <% if isedit then
			if trim(rs("oskey"))="admin" then
			%> selected="selected"<%end if%><%end if%>>一般管理员</option>


			<option value="cailiao" <% if isedit then
			if trim(rs("oskey"))="cailiao" then
			%> selected="selected"<%end if%><%end if%>>材料仓库管理员</option>


			<option value="chejian" <% if isedit then
			if trim(rs("oskey"))="chejian" then
			%> selected="selected"<%end if%><%end if%>>生产车间管理员</option>

			<option value="chengpin" <% if isedit then
			if trim(rs("oskey"))="chengpin" then
			%> selected="selected"<%end if%><%end if%>>成品仓库管理员</option>

			<option value="xiaoshou" <% if isedit then
			if trim(rs("oskey"))="xiaoshou" then
			%> selected="selected"<%end if%><%end if%>>产品销售管理员</option>





			</select>
			</td>
          </tr>
		  	<tr>
			<td colspan="7" align="center"><input type="submit"  value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        </td>
</tr>
</table>
<br><br>
<TABLE width=820 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6">
<tr>
<td >

			<form action="user.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">权 限 信 息 列 表</font></b>

</td>
</tr>
</table>
	 
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="20%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">用户名</td>
			<td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">部 门</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">权 限</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">登陆次数</td>
			<td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">上次登陆</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">上次IP</td>
            <td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">操 作</td>
          </tr>
		 <%

set rs=server.CreateObject("ADODB.RecordSet") 
rs.Source="select* from userinfo order by -userid"
rs.Open rs.Source,conn,1,1
if not rs.EOF then
do while not rs.eof
%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><a href="user.Asp?userid=<%=rs("userid")%>&action=edit"><%=rs("name")%></a></td>
<td height="18" align="center">&nbsp;<%=rs("depname")%></td>
<td height="18" align="center">&nbsp;
<%if rs("oskey")="supper" then%>超级管理员<%else%>
<%if rs("oskey")="admin" then%>一般管理员<%else%>
<%if rs("oskey")="cailiao" then%>材料仓库管理员<%else%>
<%if rs("oskey")="chejian" then%>生产车间管理员<%else%>
<%if rs("oskey")="chengpin" then%>成品仓库管理员<%else%>
<%if rs("oskey")="xiaoshou" then%>产品销售管理员<%else%>
<%end if%><%end if%><%end if%><%end if%><%end if%><%end if%></td>

<td height="18" align="center"><%=rs("logins")%></td>
<td height="18" align="center"><%=rs("lastlogin")%></td>
<td height="18" align="center"><%=rs("lastip")%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("userid")%>"></td>
</tr>
<%
rs.MoveNext
loop
end if
rs.close
set rs=nothing
%>
<tr><td align="right" colspan="7" height="22">

        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input onClick="{if(confirm('此操作将删除该信息！\n\n确定要执行此项操作吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
              <input type="Hidden" name="action" value='del'></td></tr>
        </table>

	  
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

