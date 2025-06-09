
<html>
<head>
<title>∷嘉美管理系统∷.</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript src="css/init.js">
</SCRIPT>
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 9pt}
body {  font-family: "宋体"; font-size: 9pt}
select {  font-family: "宋体"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "宋体"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
-->
</style>

</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0> 
<div style="position:absolute; top:185px; height:30px;border:0px solid white; width:232px; left:481px; margin-right:auto; margin-bottom:0">
<%
Response.Buffer = True
Response.Write "<img id=""wait"" src=""images/Run.gif""><!--" & String(1024, "/") & "-->"
Response.Flush 
%>
</div>      
<p style="text-align: left">      
<tr><br><br>　　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 

    <TD width="40%"  align="center"> 
      <script language=JavaScript src="css/date.js">
              </script>
      <script language=JavaScript src="css/time.js">
              </script>
      <script language=JavaScript src="css/init.js">
              </script>
      <span id=thename style="COLOR: navy; FONT-SIZE: 10pt" 
      align="center" name="thename"></span> 
      <script language=javascript>
  showtime(); 
</script>
    </td>
    <TD width="20%"  align="center"> 
    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    <a href="user_edit.asp">修改个人资料</a> &nbsp;<a onclick="window.parent.location.href='default.asp'" href='#'>注销登陆</a>
    </td></tr><tr>
	<td  align="center" colspan="3"></p>
<hr width="90%"  >　</td>
  </TR>

<font color="#FF0000"></font> 
<form method="post" name=myform>  
<p style="text-align: left">  
<%if action="restore" then%><INPUT TYPE="hidden" name="action" value="restore">  
<%elseif action="backup" then%><INPUT TYPE="hidden" name="action" value="backup">
<%else%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 选择操作：  <INPUT TYPE="radio" name="action" id="act_backup" value="backup"><label for=act_backup>备份</label>   
<INPUT TYPE="radio" name="action" id="act_restore" value="restore"><label for=act_restore>恢复</label>
<%end if%>  
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 数据库名：<INPUT TYPE="text" name="databasename" value="请输入数据库名" onfocus="if(this.value=='请输入数据库名')this.value=''" onblur="if(this.value=='')this.value='请输入数据库名'" size="20" >  
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 文件路径：<INPUT TYPE="text" name="bak_file" value="backup" readOnly="true"><br>  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
<input type="submit" value="确定">  </p>
<p style="text-align: left">  
　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<font color="#FF0000">数据库恢复必须在服务器端进行,先断开SQL连接（打开SQL企业管理器－右健组下面的数据库－断开）！</font></p>
</form>  
<%  
'SQL Server 数据库的备份与恢复!  
 
dim sqlserver,sqlname,sqlpassword,sqlLoginTimeout,databasename,bak_file,act  
sqlserver = "192.168.1.88" 'sql服务器  
sqlname = "sa" '用户名  
sqlpassword = "Lgm731003" '密码  
sqlLoginTimeout = 15 '登陆超时  
databasename = trim(request("databasename"))  
bak_file = trim(request("bak_file"))  
bak_file=Server.MapPath("backup/"&bak_file)
act = lcase(request("action"))   
if databasename = "" or databasename = "请输入数据库名" then  
response.write ""
  
else  
if act = "backup" then  
Set srv = Server.CreateObject("SQLDMO.SQLServer")  
srv.LoginTimeout = sqlLoginTimeout  
srv.Connect sqlserver,sqlname,sqlpassword  

Set bak = Server.CreateObject("SQLDMO.Backup")  
bak.Database = databasename  
bak.Devices = Files  
bak.Files = bak_file  
bak.Action = 0  
bak.Initialize = 1  
'bak.ReplaceDatabase = True  
bak.SQLBackup srv  
if err.number>0 then  
response.write err.number&"<font color=red><br>"  
response.write err.description&"</font>"  
end if  
Response.write "<font color=green>备份成功!</font>"  
srv.disconnect  
Set srv = nothing   
Set bak = nothing   
elseif act = "restore" then  
'恢复时要在没有使用数据库时进行！  
Set srv=Server.CreateObject("SQLDMO.SQLServer")  
srv.LoginTimeout = sqlLoginTimeout  
srv.Connect sqlserver,sqlname,sqlpassword  
Set rest = Server.CreateObject("SQLDMO.Restore")  
rest.Action = 0 ' full db restore  
rest.Database = databasename  
rest.Devices = Files  
rest.Files = bak_file  
rest.ReplaceDatabase = True 'Force restore over existing database  
if err.number>0 then  
response.write err.number&"<font color=red><br>"  
response.write err.description&"</font>"  
end if  
rest.SQLRestore srv   
Response.write "<font color=green>恢复成功!</font>"  
srv.disconnect  
Set srv = nothing  
Set rest = nothing  
else  
Response.write "<font color=red>没有选择操作</font>"  
end if  
end if  %> <br> 
<%
Response.Write "<script language=""javascript"">document.getElementById('wait').style.display='none';</script>"  '处理完成,将图片隐藏
%>
<!--#include file="footer.htm"--></BODY></HTML>