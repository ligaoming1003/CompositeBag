<!--#include file="checkuser.asp"-->
<%
userid=session("userid")
logins=session("logins")
oskey=session("oskey")
%>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<table width="100%">
  <tr>　　 
    <TD width="40%"  align="center"><%=name%>&nbsp;您的权限为:
	<%if oskey="supper" then%>超级管理员<%end if%>
<%if oskey="admin" then%>一般管理员<%end if%>

<%if oskey="cailiao" then%>材料仓库管理员<%end if%>

<%if oskey="chejian" then%>生产车间管理员<%end if%>
<%if oskey="chengpin" then%>成品仓库管理员<%end if%>
<%if oskey="xiaoshou" then%>产品销售管理员<%end if%>

&nbsp;登陆：<%=logins%> 次 </td>
    <TD width="40%"  align="center"> 
      <script language=JavaScript src="css/date.js">
              </script>
      <script language=JavaScript src="css/time.js">
              </script>
      <script language=JavaScript src="css/init.js">
              </script>
      <span id=thename style="COLOR: navy; FONT-SIZE: 10pt" 
      align="center" name="thename">dafdsfasd</span> 
      <script language=javascript>
  showtime(); 
</script>
    </td>
    <TD width="20%"  align="center"> 
    <a href="user_edit.asp">修改个人资料</a> &nbsp;<input type="submit" name="button" id="button" onclick="window.location.href='default.asp'" value="退出" />
    </td></tr>
	</table>
   
    

