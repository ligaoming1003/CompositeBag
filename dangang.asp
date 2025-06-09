<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
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
A:hover 
<!--
.thumbnail{position:relative;z-index:0;}
.thumbnail:hover{background-color:transparent;z-index:50;}
.thumbnail span{position:absolute;background-color:#FFFFE0;left:-1000px;border:1px dashed gray;visibility:hidden;color:#000;text-decoration:none;padding:5px;}
.thumbnail span img{border-width:0;padding:2px;}
.thumbnail:hover span{visibility:visible;top:20px;left:50px;}
p{margin-top:5px}
-->
{text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
-->
</style>
<SCRIPT LANGUAGE=javascript>
<!--
function Juge(powersearch)
{

	if (powersearch.pinming.value == "")
	{
		alert("品名不能为空！");
		powersearch.pinming.focus();
		return (false);
	}
	if (powersearch.chang.value == "")
	{
		alert("规格不能为空！");
		powersearch.chang.focus();
		return (false);
	}
	if (powersearch.bian.value == "")
	{
		alert("折边不能为空！");
		powersearch.bian.focus();
		return (false);
	}
    if (powersearch.di.value == "")
	{
		alert("折底不能为空！");
		powersearch.di.focus();
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
<% dim sql,rs
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

fenlei=request.Form("fenlei")  
pinming=trim(request.Form("pinming"))
chang=trim(request.Form("chang"))
kuang=trim(request.Form("kuang"))
bian=trim(request.Form("bian"))
di=trim(request.Form("di"))
kehu=request.Form("kehu")
chengping=request.Form("file1")
pihao=request.form("pihao")
tel=request.form("tel")
name1=request.form("name1")
nianbang=request.form("nianbang")
cunshu=request.form("cunshu")
diangjia=trim(request.form("diangjia"))
chati=request.form("chati")
name2=request.form("name2")
nianbang2=request.form("nianbang2")
cunshu2=request.form("cunshu2")

chati2=request.form("chati2")


set rs=server.createobject("adodb.recordset") 
sql="select * from dangang"
rs.open sql,conn,1,3
rs.addnew
rs("pihao") = pihao
rs("fenlei") = fenlei
rs("chang") = chang 
rs("kuang")=kuang
rs("pinming") = pinming
rs("kehu")=kehu
rs("bian")=bian
rs("di")=di
rs("tel")=tel
rs("name1")=name1
rs("nianbang")=nianbang
rs("cunshu")=cunshu

rs("chati")=chati
rs("name2")=name2
rs("nianbang2")=nianbang2
rs("cunshu2")=cunshu2

rs("chati2")=chati2
rs("name3")=request.form("name3")
rs("nianbang3")=request.form("nianbang3")
rs("cunshu3")=request.form("cunshu3")

rs("chati3")=request.form("chati3")
rs("name4")=request.form("name4")
rs("nianbang4")=request.form("nianbang4")
rs("cunshu4")=request.form("cunshu4")

rs("chati4")=request.form("chati4")
rs("chengping")=chengping
rs("diangjia")=diangjia
rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('档案添加成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=dangang.asp"">"
response.end

end sub

 
sub SaveModify
fenlei=request.Form("fenlei")  
pinming=trim(request.Form("pinming"))
diangjia=trim(request.form("diangjia"))
chang=trim(request.Form("chang"))
kuang=trim(request.Form("kuang"))
bian=trim(request.Form("bian"))
di=trim(request.Form("di"))
kehu=request.Form("kehu")
checkbox=trim(request.Form("checkbox"))
pihao=request.form("pihao")
chengping=request.Form("file1")
wang=year(date()) & "" & month(date()) & "" & day(date()) & "" & "" & mid("",weekday(date()),1)

if checkbox="yess" then
set rs=server.createobject("adodb.recordset") 
sql="select * from dindang"
rs.open sql,conn,1,3
rs.addnew 
rs("pihao") = pihao
rs("fenlei") = fenlei
rs("chang") = chang 
rs("kuang")=kuang
rs("pinming") = pinming+wang
rs("kehu")=kehu
'rs("uptime")=date()
rs("tel")=request.form("tel")
rs("name1")=request.form("name1")
rs("nianbang")=request.form("nianbang")
rs("cunshu")=request.form("cunshu")
rs("bian")=bian
rs("di")=di
rs("chati")=request.form("chati")
rs("name2")=request.form("name2")
rs("nianbang2")=request.form("nianbang2")
rs("cunshu2")=request.form("cunshu2")
rs("doupian")=chengping
rs("chati2")=request.form("chati2")
rs("name3")=request.form("name3")
rs("nianbang3")=request.form("nianbang3")
rs("cunshu3")=request.form("cunshu3")

rs("chati3")=request.form("chati3")
rs("name4")=request.form("name4")
rs("nianbang4")=request.form("nianbang4")
rs("cunshu4")=request.form("cunshu4")

rs("chati4")=request.form("chati4")
rs("diangjia")=diangjia
rs.update
rs.close
set rs=nothing
end if

set rs=server.createobject("adodb.recordset") 
sql="select * from dangang where id="&request.Form("id")
rs.open sql,conn,1,3
rs("pihao") = pihao
rs("fenlei") = fenlei
rs("chang") = chang 
rs("kuang")=kuang
rs("pinming") = pinming
rs("kehu")=kehu
rs("bian")=bian
rs("di")=di
rs("tel")=request.form("tel")
rs("name1")=request.form("name1")
rs("nianbang")=request.form("nianbang")
rs("cunshu")=request.form("cunshu")

rs("chati")=request.form("chati")
rs("name2")=request.form("name2")
rs("nianbang2")=request.form("nianbang2")
rs("cunshu2")=request.form("cunshu2")

rs("chati2")=request.form("chati2")
rs("name3")=request.form("name3")
rs("nianbang3")=request.form("nianbang3")
rs("cunshu3")=request.form("cunshu3")

rs("chati3")=request.form("chati3")
rs("name4")=request.form("name4")
rs("nianbang4")=request.form("nianbang4")
rs("cunshu4")=request.form("cunshu4")

rs("chati4")=request.form("chati4")
rs("chengping")=chengping
rs("diangjia")=diangjia
rs.update
rs.close
set rs=nothing
if checkbox="yess" then
response.Write "<script language=javascript>alert('订单添加成功！');</script>"
else
response.Write "<script language=javascript>alert('档案修改成功！');</script>"
end if
response.write "<meta http-equiv=""refresh"" content=""0;url=dangang.asp"">"
response.end

end sub   
 
  sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=dangang.asp"">"
  response.end
else
        conn.execute("delete from dangang where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"

		response.write "<meta http-equiv=""refresh"" content=""0;url=dangang.asp"">"
		end if
response.end
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from dangang where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑档案信息</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加档案信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="9%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">版号</td>
            <td width="8%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">客户</td>
            <td width="9%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">电话</td>
			<td width="18%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">品 名</td>
			<td width="4%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">长 度</td>			
			<td width="4%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">宽 度</td>			
			<td width="5%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			折边</td>			
			<td width="5%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			折底</td>			
			<td width="4%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			单价</td>			
			<td width="8%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">类 别</td>			
			
          </tr>
		  <form name="powersearch" method="post" action="dangang.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="pihao" type="text"  size="12" value='<% if isedit then
		                                                         response.write rs("pihao")
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<script language = "javascript">
var i,j;
j=0;
goaler = new Array();
<%
set rs_p=server.createobject("adodb.recordset")
rs_p.open "select * from kehu",conn,1,1
if rs_p.eof then%>
goaler[0] = new Array("","");
<%else
i=0
do while not rs_p.eof%>
goaler[<%=i%>] = new Array("<%=rs_p("tel")%>","<%=rs_p("kehu")%>","<%=rs_p("id")%>");
<%rs_p.movenext
i=i+1
loop
end if
rs_p.close
%>
j=<%=i%>;

function changelocation(id)//传递一级分类的值,从而改变二级分类
{
//document.getElementById("tel").value="";
var i;
for (i=0;i < j; i++)
{
if (goaler[i][1] ==id)
document.getElementById("tel").value = goaler[i][0];
}
}
</script>
			
			
			
			<select name="kehu" style='width:65; height:16' id="cname" onChange="changelocation(this.options[this.selectedIndex].value)">

			<% if isedit then%>
			<option selected="selected"><%=rs("kehu")%></option>
			<%else%>
			<option selected value="">所有客户</option>
            <%end if%>
			
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from kehu  order by kehu",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   kehu=rs2("kehu")
		   %>
		   <option value="<%=rs2("kehu")%>"><%=rs2("kehu")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="tel" type="text"  size="12" value='<% if isedit then
		                                                         response.write rs("tel")
																  end if %>'>				</td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="pinming" type="text" size="28" value='<% if isedit then
		                                                         response.write rs("pinming")
																  end if %>'maxlength=42></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="chang" type="text" style="ime-mode:disabled"  size="5" value='<% if isedit then
		                                                         response.write rs("chang")
																  end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="kuang" type="text" style="ime-mode:disabled"  size="5" value='<% if isedit then
		                                                         response.write rs("kuang")
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="bian" type="text" style="ime-mode:disabled"  size="5" value='<% if isedit then
		                                                         response.write rs("bian")
		                                                         else response.write 0
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="di" type="text" style="ime-mode:disabled"  size="5" value='<% if isedit then
		                                                         response.write rs("di")
		                                                         else response.write 0
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="diangjia" type="text" style="ime-mode:disabled"  size="6" value='<% if isedit then
		                                                         response.write rs("diangjia")
		                                                         else response.write 0
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name="fenlei" size="1" style='width:73; height:18'>
			<% if isedit then%>
			<option><%=rs("fenlei")%></option><%end if%>
			
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from fenlei1",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option <% if isedit then
			if trim(rs1("fenlei"))=trim(rs("fenlei")) then
			%> selected="selected"<%end if%><%end if%>><%=rs1("fenlei")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td>


          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<font color="#FFFFFF">彩印面膜</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<font color="#FFFFFF">面膜宽度</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<font color="#FFFFFF">面膜厚度</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="3">
<input class="iText" type="text" name="file1" size="30" value='<% if isedit then
		                                                         response.write rs("chengping")
																 else 
																 response.write "files/2106.jpg"
																  end if %>'>
<input class="iButton" type="button" onclick="window.open('Upload.asp','_blank','Width=480 Height=50 top=300 left=300 status=1');" value="上传" /></td>


<td height="12" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="4">
<font color="#FFFFFF">彩印要求</font></td>


          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='name1'  style='width:55; height:16'>
			<% if isedit then%>
			<option selected value='0'>请选择</option>
			<option selected="selected"><%=rs("name1")%></option>
			<%else%>
			<option selected value='0'>请选择</option>
			<%end if%>
						<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from cailiao where fenlei='面膜' order by pinming",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("pinming")%>"><%=rs2("pinming")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="nianbang"  size="1" style='width:54; height:18'>
			<% if isedit then%>
			<option selected value='0'>请选择</option>
			<option ><%=rs("nianbang")%></option>
			<%else%>
			<option selected value='0'>请选择</option>
			<%end if%>
			<%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from guige order by guige",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if rs3("guige")=rs("nianbang") then
			%> selected="selected"<%end if%><%end if%>><%=rs3("guige")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="cunshu"  size="1" style='width:55; height:16'>
			<% if isedit then%><option selected value='0'>请选择</option><option ><%=rs("cunshu")%></option>
						<%else%>
			<option selected value='0'>请选择</option><%end if%>
			<%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option <% if isedit then
			if rs4("houdou")=rs("cunshu") then
			%> selected="selected"<%end if%><%end if%>><%=rs4("houdou")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
			%>
			</select></td>
<td height="126" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="3" rowspan="7">
			<% if isedit then%><font color="#FFFFFF">添加订单，请勾选</font><p><input type="checkbox" name="checkbox" value="yess"><%end if%></td>


<td height="36" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="4" rowspan="3">
<textarea name="chati" cols="33" rows="3"><% if isedit then
		                                                         response.write rs("chati")
																  end if %></textarea></td>


          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<font color="#FFFFFF">复合中膜</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<font color="#FFFFFF">中膜宽度</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<font color="#FFFFFF">中膜厚度</font></td>


          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='name2'  style='width:55; height:16'>
			<% if isedit then%><option selected value='0'>请选择</option>
			<option selected="selected"><%=rs("name2")%></option>
			<%else%>
			<option selected value='0'>请选择</option>
			<%end if%>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from cailiao where fenlei='中膜' order by pinming",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("pinming")%>"><%=rs2("pinming")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="nianbang2"  size="1" style='width:54; height:18'>
			<% if isedit then%><option selected value='0'>请选择</option><option ><%=rs("nianbang2")%></option>			<%else%>
			<option selected value='0'>请选择</option><%end if%>
			<%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from guige order by guige",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if rs3("guige")=rs("nianbang2") then
			%> selected="selected"<%end if%><%end if%>><%=rs3("guige")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="cunshu2"  size="1" style='width:55; height:16'>
			<% if isedit then%><option selected value='0'>请选择</option><option ><%=rs("cunshu2")%></option>			<%else%>
			<option selected value='0'>请选择</option><%end if%>
			<%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option <% if isedit then
			if rs4("houdou")=rs("cunshu2") then
			%> selected="selected"<%end if%><%end if%>><%=rs4("houdou")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
			%>
			</select></td>


          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<font color="#FFFFFF">复合鸳膜</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<font color="#FFFFFF">鸳膜宽度</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<font color="#FFFFFF">鸳膜厚度</font></td>
<td height="12" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="4">
<font color="#FFFFFF">复合要求</font></td>


          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='name3'  style='width:55; height:16'>
				<% if isedit then%><option selected value='0'>请选择</option>
			<option selected="selected"><%=rs("name3")%></option>
			<%else%>
			<option selected value='0'>请选择</option>
			<%end if%>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from cailiao where fenlei='中膜' order by pinming",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("pinming")%>"><%=rs2("pinming")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="nianbang3"  size="1" style='width:54; height:18'>
			<% if isedit then%><option selected value='0'>请选择</option><option ><%=rs("nianbang3")%></option>			<%else%>
			<option selected value='0'>请选择</option><%end if%>
			<%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from guige order by guige",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if rs3("guige")=rs("nianbang3") then
			%> selected="selected"<%end if%><%end if%>><%=rs3("guige")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="cunshu3"  size="1" style='width:55; height:16'>
			<% if isedit then%><option selected value='0'>请选择</option><option ><%=rs("cunshu3")%></option>			<%else%>
			<option selected value='0'>请选择</option><%end if%>
			<%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option <% if isedit then
			if rs4("houdou")=rs("cunshu3") then
			%> selected="selected"<%end if%><%end if%>><%=rs4("houdou")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
			%>
			</select></td>
<td height="36" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="4" rowspan="3">
<textarea name="chati4" cols="33" rows="3"><% if isedit then
		                                                         response.write rs("chati4")
																  end if %></textarea></td>


          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<font color="#FFFFFF">复合内膜</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<font color="#FFFFFF">内膜宽度</font></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<font color="#FFFFFF">内膜厚度</font></td>


          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='name4'  style='width:55; height:16'>
				<% if isedit then%><option selected value='0'>请选择</option>
			<option selected="selected"><%=rs("name4")%></option>
			<%else%>
			<option selected value='0'>请选择</option>
			<%end if%>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from cailiao where fenlei='内膜' order by pinming",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("pinming")%>"><%=rs2("pinming")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="nianbang4"  size="1" style='width:54; height:18'>
			<% if isedit then%><option selected value='0'>请选择</option><option ><%=rs("nianbang4")%></option>			<%else%>
			<option selected value='0'>请选择</option><%end if%>
			<%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from guige order by guige",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if rs3("guige")=rs("nianbang4") then
			%> selected="selected"<%end if%><%end if%>><%=rs3("guige")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="cunshu4"  size="1" style='width:55; height:16'>
			<% if isedit then%><option selected value='0'>请选择</option><option ><%=rs("cunshu4")%></option>			<%else%>
			<option selected value='0'>请选择</option><%end if%>
			<%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option <% if isedit then
			if rs4("houdou")=rs("cunshu4") then
			%> selected="selected"<%end if%><%end if%>><%=rs4("houdou")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
			%>
			</select></td>


          </tr>
           <tr> 
<td height="71" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<font color="#FFFFFF">成品要求</font></td>
<td height="71" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="5">
<textarea name="chati2" cols="49" rows="4"><% if isedit then
		                                                         response.write rs("chati2")
																  end if %></textarea></td>


<td height="71" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="4">
<font color="#FFFFFF">包装要求：</font><textarea name="chati3" cols="24" rows="2"><% if isedit then
		                                                         response.write rs("chati3")
																  end if %></textarea></td>


          </tr>

			<tr>
			<td colspan="10" align="center"><input type="submit"  value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        <%end if%>

</td>
</tr>
</table><br><br>
<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >

  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp; 档 案 列 表</font></b>

</td>
</tr>
</table>
			<form action="dangang.asp" method="post" name="selform">
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD" height="18" align="left">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">类 别</td>
			<td width="31%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">品 名</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">规 格</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">面膜</td>
			<td width="20%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">材料</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">客户</td>
            <td width="12%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			版号</td>
            <td width="4%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">操作</td>
          </tr>
		 <%
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 1000          '每页显示多少条
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet") 
search=request("search")

if search<>"" then 
akeyword=request("akeyword")
keyword=request("keyword")
rs.Source="select* from dangang where kehu like '%"&akeyword&"%' and pinming like '%"&keyword&"%' order by -id "
else
rs.Source="select* from dangang order by -id"
end if

rs.Open rs.Source,conn,1,1
if not rs.EOF then
rs.PageSize     = MyPageSize
MaxPages         = rs.PageCount
rs.absolutepage = MyPage
total            = rs.RecordCount
for i=1 to rs.PageSize
if not rs.EOF then
%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><a class="thumbnail" href="#"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%><span><img src="<%=rs("chengping")%>" height="400"></span></a></td>
<td height="18" align="center">&nbsp;<a href="dangang.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18" align="center"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%if rs("nianbang")=0 then%>找料<%else%><%=rs("nianbang")%><%end if%></td>
<td height="18" align="center"><p align="left">&nbsp;<%=rs("name1")%>│<%if rs("name2")="0" then%><%else%><%=rs("name2")%>│<%end if%><%if rs("name3")="0" then%><%else%><%=rs("name3")%>│<%end if%><%if rs("name4")="0" then%><%else%><%=rs("name4")%><%end if%></td>
<td height="18" align="center"><a href="dangang_in.asp?id=<%=rs("id")%>&action=edit"><%=rs("kehu")%></a></td>
<td height="18" align="center"><%=rs("pihao")%><%if rs("pihao")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="8" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="dangang.asp?keyword="&keyword&"&akeyword="&akeyword&"&search=search&"
			else
            url="dangang.asp?"
			end if				
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
rs.Close
set rs=nothing
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="cailiao"  or oskey="admin" then%>
     <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input onClick="{if(confirm('此操作将删除该信息！\n\n确定要执行此项操作吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
              <input type="Hidden" name="action" value='del'><%end if%></td></tr>
        </table>

	  
</td>
    </tr>
      </table>
</form>


</td>
</tr>
</table>
</td>
</tr></table>
<br>
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>


	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<font color="#FFFFFF"><b>档</b></font><b><font color="#ffffff">案搜索</font></b>

</td>
</tr>
</table>

        <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="dangang.asp">
		  <tr><input name="search" type="hidden" value="search">
          <td align="center" width="36%">按品名查找： 
				<input name="keyword" type="text"  size="23"></td>
                      
			<td align="center" width="13%" ><input type="submit"  value="查 找"> </td>
			<td align="center" width="35%">按客户查找： 
					<select name='akeyword'  style='width:130'>
			<option selected value="">请选择客户</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(kehu) from dangang order by kehu",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("kehu")%>"><%=rs2("kehu")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
                      
			<td align="center" width="17%" ><input type="submit"  value="查 找"> </td>

          </tr></form>		   
		</td></tr>
	   </table>
		<table width="100%" border="0" cellspacing="0" cellpadding="4">
      
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
