<!--#include file="checkuser.asp"-->
<%
userid=session("userid")
logins=session("logins")
oskey=session("oskey")
%>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK href="images/style.css" type=text/css rel=stylesheet>
<table width="100%">
  <tr>���� 
    <TD width="40%"  align="center"><%=name%>&nbsp;����Ȩ��Ϊ:
	<%if oskey="supper" then%>��������Ա<%end if%>
<%if oskey="admin" then%>һ�����Ա<%end if%>

<%if oskey="cailiao" then%>���ϲֿ����Ա<%end if%>

<%if oskey="chejian" then%>�����������Ա<%end if%>
<%if oskey="chengpin" then%>��Ʒ�ֿ����Ա<%end if%>
<%if oskey="xiaoshou" then%>��Ʒ���۹���Ա<%end if%>

&nbsp;��½��<%=logins%> �� </td>
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
    <a href="user_edit.asp">�޸ĸ�������</a> &nbsp;<input type="submit" name="button" id="button" onclick="window.location.href='default.asp'" value="�˳�" />
    </td></tr>
	</table>
   
    

