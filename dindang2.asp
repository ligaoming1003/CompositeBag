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
response.write "<script language=javascript>alert('δѡ����¼��');history.go(-1);</script>"
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
 response.Write "<script language=javascript>alert('���в�����ȫ���Ѱ����׵���');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=dindang1.asp"">"
        response.end
       
end if
        %>

