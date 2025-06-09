<!--#include file="conn.asp"-->
<%
name=session("name")
depid=session("depid")
oskey=session("oskey")
userid=session("userid")

if userid<>"" then
set rs=server.CreateObject("ADODB.RecordSet")
sql="select * from department where id="&depid
rs.open sql,conn,1,3
if rs.eof or rs.bof then
depname=rs("depname")
end if
rs.close
set rs=nothing
end if
if session("name")="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if%>
