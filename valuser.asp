<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->

<%
checkbox1=request.form("checkbox1")
checkbox2=request.form("checkbox2")
checkbox3=request.form("checkbox3")
name=request.form("username")
pwd=md5(trim(request("pwd")))

if  checkbox1="" and checkbox2="" and checkbox3=""  then
response.Write "<script language=javascript>alert('请选择部门！');history.go(-1);</script>"
response.End
end if



if  checkbox2="复合车间" and checkbox3="办公室" and checkbox1="印刷车间"  then
response.Write "<script language=javascript>alert('部门只能选择一个！');history.go(-1);</script>"
response.End
end if


if  checkbox2="复合车间" and checkbox3="办公室" then
response.Write "<script language=javascript>alert('部门只能选择一个！');history.go(-1);</script>"
response.End
end if


if  checkbox1="印刷车间" and checkbox3="办公室" then
response.Write "<script language=javascript>alert('部门只能选择一个！');history.go(-1);</script>"
response.End
end if


if  checkbox1="印刷车间" and checkbox2="复合车间" then
response.Write "<script language=javascript>alert('部门只能选择一个！');history.go(-1);</script>"
response.End
end if

if  checkbox1="印刷车间" then
set rs=server.CreateObject("ADODB.RecordSet")
sql="select * from userinfo where name='"&name&"' and pwd='"&pwd&"'and depname='"&checkbox1&"'"
on error resume next
rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.Write "<script language=javascript>alert('用户名、密码、部门\n\n出现错误！');history.go(-1);</script>"
response.End
else
rs("logins")=rs("logins")+1
rs("lastlogin")=now()
rs("lastip")=Request.ServerVariables("REMOTE_ADDR")
rs.update
session.Timeout=20
session("userid")=rs("userid")
session("name")=rs("name")
session("oskey")=rs("oskey")
session("depid")=rs("depid")
session("logins")=rs("logins")
rs.update
rs.close
conn.close
set rs=nothing
set conn=nothing
response.redirect "cy.asp"
response.end
end if
end if

if checkbox2="复合车间" then
set rs=server.CreateObject("ADODB.RecordSet")
sql="select * from userinfo where name='"&name&"' and pwd='"&pwd&"'and depname='"&checkbox2&"'"
on error resume next
rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.Write "<script language=javascript>alert('用户名、密码、部门\n\n出现错误！');history.go(-1);</script>"
response.End
else
rs("logins")=rs("logins")+1
rs("lastlogin")=now()
rs("lastip")=Request.ServerVariables("REMOTE_ADDR")
rs.update
session.Timeout=20
session("userid")=rs("userid")
session("name")=rs("name")
session("oskey")=rs("oskey")
session("depid")=rs("depid")
session("logins")=rs("logins")
rs.update
rs.close
conn.close
set rs=nothing
set conn=nothing
response.redirect "fh.asp"
response.end
end if
end if


if checkbox3="办公室" then
set rs=server.CreateObject("ADODB.RecordSet")
sql="select * from userinfo where name='"&name&"' and pwd='"&pwd&"'and depname='"&checkbox3&"'"
on error resume next
rs.open sql,conn,1,3
if rs.eof or rs.bof then
response.Write "<script language=javascript>alert('用户名或密码错误！');history.go(-1);</script>"
response.End
else
rs("logins")=rs("logins")+1
rs("lastlogin")=now()
rs("lastip")=Request.ServerVariables("REMOTE_ADDR")
rs.update
session.Timeout=20
session("userid")=rs("userid")
session("name")=rs("name")
session("oskey")=rs("oskey")
session("depid")=rs("depid")
session("logins")=rs("logins")
rs.update
rs.close
conn.close
set rs=nothing
set conn=nothing
response.redirect "index.asp"
response.end
end if
end if
%>
	
