
<html>
<head>
<title>�˼�������ϵͳ��.</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript src="css/init.js">
</SCRIPT>
<style type="text/css">
<!--
td {  font-family: "����"; font-size: 9pt}
body {  font-family: "����"; font-size: 9pt}
select {  font-family: "����"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "����"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "����"; font-size: 9pt} 
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
<tr><br><br>����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 

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
    <a href="user_edit.asp">�޸ĸ�������</a> &nbsp;<a onclick="window.parent.location.href='default.asp'" href='#'>ע����½</a>
    </td></tr><tr>
	<td  align="center" colspan="3"></p>
<hr width="90%"  >��</td>
  </TR>

<font color="#FF0000"></font> 
<form method="post" name=myform>  
<p style="text-align: left">  
<%if action="restore" then%><INPUT TYPE="hidden" name="action" value="restore">  
<%elseif action="backup" then%><INPUT TYPE="hidden" name="action" value="backup">
<%else%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ѡ�������  <INPUT TYPE="radio" name="action" id="act_backup" value="backup"><label for=act_backup>����</label>   
<INPUT TYPE="radio" name="action" id="act_restore" value="restore"><label for=act_restore>�ָ�</label>
<%end if%>  
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���ݿ�����<INPUT TYPE="text" name="databasename" value="���������ݿ���" onfocus="if(this.value=='���������ݿ���')this.value=''" onblur="if(this.value=='')this.value='���������ݿ���'" size="20" >  
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �ļ�·����<INPUT TYPE="text" name="bak_file" value="backup" readOnly="true"><br>  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
<input type="submit" value="ȷ��">  </p>
<p style="text-align: left">  
��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<font color="#FF0000">���ݿ�ָ������ڷ������˽���,�ȶϿ�SQL���ӣ���SQL��ҵ���������ҽ�����������ݿ⣭�Ͽ�����</font></p>
</form>  
<%  
'SQL Server ���ݿ�ı�����ָ�!  
 
dim sqlserver,sqlname,sqlpassword,sqlLoginTimeout,databasename,bak_file,act  
sqlserver = "192.168.1.88" 'sql������  
sqlname = "sa" '�û���  
sqlpassword = "Lgm731003" '����  
sqlLoginTimeout = 15 '��½��ʱ  
databasename = trim(request("databasename"))  
bak_file = trim(request("bak_file"))  
bak_file=Server.MapPath("backup/"&bak_file)
act = lcase(request("action"))   
if databasename = "" or databasename = "���������ݿ���" then  
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
Response.write "<font color=green>���ݳɹ�!</font>"  
srv.disconnect  
Set srv = nothing   
Set bak = nothing   
elseif act = "restore" then  
'�ָ�ʱҪ��û��ʹ�����ݿ�ʱ���У�  
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
Response.write "<font color=green>�ָ��ɹ�!</font>"  
srv.disconnect  
Set srv = nothing  
Set rest = nothing  
else  
Response.write "<font color=red>û��ѡ�����</font>"  
end if  
end if  %> <br> 
<%
Response.Write "<script language=""javascript"">document.getElementById('wait').style.display='none';</script>"  '�������,��ͼƬ����
%>
<!--#include file="footer.htm"--></BODY></HTML>