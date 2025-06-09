<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->

<SCRIPT LANGUAGE=javascript>
<!--
function Juge(myform)
{

	
}


function SelectAll() {
	for (var i=0;i<document.selform.selBigClass.length;i++) {
		var e=document.selform.selBigClass[i];
		e.checked=!e.checked;
	}
}
//-->
</script>
<%
 
Sum = trim(request.form("selBigClass").count)
if Sum =0 then 
response.write "<script language=javascript>alert('未选定记录！');history.go(-1);</script>"
response.end
else
for i = 1 to Sum
selBigClass= request.form("selBigClass")(i)
set rs=server.createobject("adodb.recordset")
sql="select * from dindang where id in ("&selBigClass&")"
rs.open sql,conn,1,3

rs("zhuantai")=""

rs.update 
next 
rs.close
 response.Write "<script language=javascript>alert('所有材料齐全，已安排妥当！');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=dindang1.asp"">"
        response.end
       
end if
        %>

