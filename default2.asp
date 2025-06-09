<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<HTML><HEAD>
<title>∷嘉美管理系统∷</title>

<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<!--#include file="conn.asp"-->
<LINK href="images/dcportal.css" rel=stylesheet type=text/css>
<script Language="JavaScript">
 function form_check(){
   if(document.form1.username.value==""){window.alert("登录名不能为空");document.form1.username.focus();return (false);}
   if(document.form1.pwd.value==""){window.alert("登录密码不能为空");document.form1.pwd.focus();return (false);}
      }
function loginclick( ) {
	if(form_Check())
		thisForm.submit();
}
</script>
</HEAD>
<BODY>
<form name="form1" method="post" action="valuser.asp" onSubmit="return form_check();">
<FONT face=宋体>
<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" id=Table1 
width="100%">
  <TBODY>
  <TR>
    <TD background="images/login_topBg.jpg" height=90><FONT 
      face=宋体>
      <TABLE border=0 cellPadding=0 cellSpacing=0 id=Table2 width="100%">
        <TBODY>
        <TR>
          <TD><img border=0 id=Image1 src="images/login_topPic1.jpg"></TD>
          <TD align=right><IMG border=0 id=Image2 
            src="images/login_topPic2.jpg"></TD></TR></TBODY></TABLE></FONT></TD></TR>
  <TR>
    <TD bgColor=#07359b height=4><FONT face=宋体>
      <TABLE border=0 cellPadding=0 cellSpacing=0 id=Table3 width="100%">
        <TBODY>
        <TR>
          <TD><IMG border=0 id=Image3 
            src="images/login_line1.jpg"></TD>
          <TD align=right><IMG border=0 id=Image4 
            src="images/login_line1_r.jpg"></TD></TR></TBODY></TABLE></FONT></TD></TR>
  <TR>
    <TD align="center" 
    style="BACKGROUND-COLOR: #7e9fd8; BACKGROUND-POSITION: left top; BACKGROUND-REPEAT: no-repeat"><FONT 
      face=宋体>
      <TABLE border=0 cellPadding=5 cellSpacing=0 id=Table4 width=260>
        <TBODY>
        <TR>
          <TD height=25><FONT face=宋体>
            <TABLE border=0 cellPadding=0 cellSpacing=0 id=Table5 
              width="100%"><TBODY>
              <TR>
                <TD align=right><IMG border=0 id=Image7 
                  src="images/login_arrow.gif"></TD>
                <TD align="center"><img border=0 id=Image8 src="images/login_font.jpg" alt="∷嘉美管理系统∷"></TD>
              </TR></TBODY></TABLE></FONT></TD></TR>
        <TR>
          <TD>
            <TABLE bgColor=#9cb5e1 border=0 cellPadding=5 cellSpacing=0 
            id=Table7 width="100%">
              <TBODY>
              <TR>
                <TD align=right class=loginBorder height=3><FONT face=宋体><SPAN 
                  id=Label3 style="COLOR: white">登录系统：</SPAN></FONT></TD>
                <TD class=loginBorder height=3><span class="style1"><FONT face=宋体>∷嘉美管理系统∷</FONT></span></TD>
              </TR>
              <TR>
                <TD align=right class=loginBorder><FONT face=宋体><SPAN 
                  id=Label1 style="COLOR: white">用户名：</SPAN></FONT></TD>
                <TD class=loginBorder><FONT face=宋体>
				<INPUT 
                  name=username id=txtUser 
                  style="COLOR: maroon; FONT-FAMILY: Tahoma; FONT-WEIGHT: bold; WIDTH: 150px" onmouseover=this.select() value="" size="20" >
                </FONT></TD></TR>
              <TR>
                <TD align=right class=loginBorderB><FONT face=宋体><SPAN 
                  id=Label2 style="COLOR: white">密  码：</SPAN></FONT></TD>
                <TD class=loginBorderB>
				<INPUT name=pwd 
                  type=password id=txtPassword 
                  style="COLOR: maroon; FONT-FAMILY: Tahoma; FONT-WEIGHT: bold; WIDTH: 150px" value="" size="20" ></TD></TR>
              <TR>
                <TD align=right height=3><FONT face=宋体></FONT></TD>
                <TD height=3></TD></TR></TBODY></TABLE></TD></TR>
        <TR>
          <TD height=35>
            <TABLE border=0 cellPadding=0 cellSpacing=0 id=Table6 
              width="100%"><TBODY>
              <TR>
                <TD align="center" width="50%"><FONT face=宋体><INPUT border=0 
                  id=btnEnter name=btnEnter 
                  onmouseout="this.src='Images/login_btnLogin.jpg'" 
                  onmouseover="this.src='Images/login_btnLogin_hover.jpg'" 
                  src="images/login_btnLogin.jpg" 
                  type="image" onClick="loginclick( )"></FONT></TD>
                <TD align=middle><INPUT border=0 id=btnExit name=btnExit 
                  onclick=window.close() 
                  onmouseout="this.src='Images/login_btnCancel.jpg'" 
                  onmouseover="this.src='Images/login_btnCancel_hover.jpg'" 
                  src="images/login_btnCancel.jpg" 
              type=image></TD>







</TR></TBODY></TABLE></TD></TR></TBODY></TABLE></FONT>
			  <table width="50%"  border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>　</td>
  </tr>
  <tr>
    <td></td>
  </tr>
  <tr>
    <td>　</td>
  </tr>
</table>

	    </TD></TR>
  <TR>
    <TD bgColor=#07359b height=4><FONT face=宋体><IMG border=0 id=Image6 
      src="images/login_line2.jpg"></FONT></TD></TR>
  <TR>
    <TD align=middle bgColor=#07359b height=40><FONT face=宋体>
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center"><FONT face=宋体 color="#FFFFFF"><font color="#ffffff">嘉美包装</font> 版权所有  </font><FONT face=宋体><SPAN id=Label4 style="COLOR: white">技术支持：
	汪文武</span></FONT> <img src="images/ico.gif"></span></FONT><br>
	<SPAN id=Label1 style="COLOR: white"><br>　</td>
  </tr>
</table>
</FONT></TD></TR></TBODY></TABLE></FONT></FORM></BODY></HTML>

