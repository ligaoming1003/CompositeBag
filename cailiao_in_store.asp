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
A:hover {text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
-->
</style>
<SCRIPT LANGUAGE=javascript>
<!--
function Juge(powersearch)
{

	if (powersearch.pinming.value == "")
	{
		alert("名称不能为空！");
		powersearch.pinming.focus();
		return (false);
	}
	if (powersearch.use_num.value == "")
	{
		alert("数量不能为空！");
		powersearch.use_num.focus();
		return (false);
	}
	if (powersearch.danjia.value == "")
	{
		alert("单价不能为空！");
		powersearch.danjia.focus();
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


<script>
  var po_ca_show = new Array();
  var po_ca_value = new Array();
  var po_detail_show = new Array();
  var po_detail_value = new Array();
  var funtype1;
//Modified by Ryan Gao 2006-03-30 16:01:09
<%
i=0
		
set rs1=server.createobject("adodb.recordset")
rs1.open "select * from fenlei",conn,1,1
if not rs1.eof then   
do while not rs1.eof
j=0	
%>
po_ca_show[<%=i%>] = '<%=rs1("fenleiname")%>';
po_ca_value[<%=i%>] = '<%=rs1("fenleiname")%>';
po_detail_show[<%=i%>] = new Array();
po_detail_value[<%=i%>] = new Array();

<%
fenleiname=rs1("fenleiname")
set rs2=server.createobject("adodb.recordset")
rs2.open "select * from cailiao where fenlei='"&fenleiname&"' order by pinming",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = ' <%=rs2("pinming")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("pinming")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '此类暂无材料';
po_detail_value[<%=i%>][<%=j%>] = '';
<%
end if
rs2.close
set rs2=nothing
%>

<%
rs1.movenext
i=i+1
loop
end if
rs1.close
set rs1=nothing
%>

</script>
<script Language="JavaScript">



var psid="";

function DoLoad(form,funtypev){
        var n;
        var i,j,k;
        var num;
        num= GetObjID('funtype[]');
        num1= GetObjID('pinming');

        if (!funtypev)
           return;
        k=0;
        for (i=0;i<po_ca_show.length;i++) {
         for(j = 0; j < po_detail_value[i].length; j++){
                if(funtypev.indexOf(po_detail_value[i][j])!=-1) {
                    NewOptionName = new Option(po_detail_show[i][j], po_detail_value[i][j]);
                    form.elements[num].options[k] = NewOptionName;
                    k++;
                    }
         }
        }
}

function Do_po_Change(form){
        var num,n, i, m;
        num= GetObjID('d_position1');
        m = document.powersearch.elements[num].selectedIndex-1;
        n = document.powersearch.elements[num + 1].length;
        for(i = n - 1; i >= 0; i--)
                document.powersearch.elements[num + 1].options[i] = null;

        if (m>=0) {
        for(i = 0; i < po_detail_show[m].length; i++){
                NewOptionName = new Option(po_detail_show[m][i], po_detail_value[m][i]);
                document.powersearch.elements[num + 1].options[i] = NewOptionName;
        }
        document.powersearch.elements[num + 1].options[0].selected = true;
        }
}

function InsertItem(ObjID, Location)
{
	len=document.powersearch.elements[ObjID].length;
	for (counter=len; counter>Location; counter--)
	{
		Value = document.powersearch.elements[ObjID].options[counter-1].value;
		Text2Insert  = document.powersearch.elements[ObjID].options[counter-1].text;
		document.powersearch.elements[ObjID].options[counter] = new Option(trimPrefixIndent(Text2Insert), Value);
	}
}
function GetLocation(ObjID, Value)
{
  total=document.powersearch.elements[ObjID].length;
  for (pp=0; pp<total; pp++)
     if (document.powersearch.elements[ObjID].options[pp].text == "---"+Value+"---")
     { return (pp);
       break;
     }
  return (-1);
}

function GetObjID(ObjName)
{
  for (var ObjID=0; ObjID < window.powersearch.elements.length; ObjID++)
    if ( window.powersearch.elements[ObjID].name == ObjName )
    {  return(ObjID);
       break;
    }
  return(-1);
}




</script>
</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0>       
<!--#include file="top.asp"-->
<%
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
fenlei1=request.Form("d_position1")  
pinming=request.Form("pinming")
guige=request.Form("guige")
houdou=request.Form("houdou")
use_num=request.Form("use_num")
gonghuodanwei=trim(request.Form("gonghuodanwei"))
'吹膜纸芯按每六十公斤一公斤计算
if gonghuodanwei="吹膜" then
set rs2=server.createobject("adodb.recordset") 
sql2="select * from fuliao_store"
rs2.open sql2,conn,2,3
rs2("number")=rs2("number")-(use_num-use_num/60)
rs2.update
rs2.close
set rs2=nothing
end if
set rs1=server.createobject("adodb.recordset") 
if guige="" then
sql1="select * from cailiao_store where pinming='"&pinming&"'and houdou="&houdou&" and fenlei='"&fenlei1&"'"
else
sql1="select * from cailiao_store where pinming='"&pinming&"' and guige="&guige&" and houdou="&houdou&" and fenlei='"&fenlei1&"'"
end if
rs1.open sql1,conn,1,3
if not rs1.eof and not rs1.bof then
rs1("number")=rs1("number")+request.Form("use_num")

else
rs1.addnew
rs1("fenlei") = trim(request.Form("d_position1"))
rs1("Unit") = trim(request.Form("unit"))
rs1("guige") = guige 
rs1("houdou")=houdou
rs1("pinming") = pinming
rs1("number")=request.Form("use_num")

end if
rs1.update
rs1.close
set rs1=nothing

set rs=server.createobject("adodb.recordset") 
sql="select * from cailiao_in_store"
rs.open sql,conn,1,3
rs.addnew
rs("fenlei") = trim(request.Form("d_position1"))
rs("guige") = guige
rs("houdou") = houdou
rs("pinming") = pinming
rs("uptime") = request.Form("uptime")
rs("Unit") = trim(request.Form("unit"))
rs("gonghuodanwei") = trim(request.Form("gonghuodanwei"))
rs("use_num") = request.Form("use_num")
rs("danjia") = request.Form("danjia")
rs("jinge") =request.Form("danjia")*request.Form("use_num")
rs("content") = trim(request.Form("content"))
rs("zhuangtai")="ok"
rs.update
response.Write "<script language=javascript>alert('入库成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_in_store.asp"">"
response.end
rs.close
set rs=nothing
end sub

 
sub SaveModify
use_num1=request.Form("use_num1")
use_num=request.Form("use_num")
danjia=request.Form("danjia")
fenlei1=request.Form("d_position1")  
pinming=request.Form("pinming")
pinming1=request.Form("pinming1")
guige=request.Form("guige")
guige1=request.Form("guige1")
unit=request.form("unit")
unit1=request.form("unit1")
houdou=request.form("houdou")
houdou1=request.form("houdou1")
use_num=request.Form("use_num")
gonghuodanwei=trim(request.Form("gonghuodanwei"))
num=use_num1-use_num


if pinming<>pinming1 or guige<>guige1 or unit<>unit1 or houdou<>houdou1 then
      response.Write "<script language=javascript>alert('对不起,只有数量、单价能修改！\n\n如果要修改其它项目,\n\n请先将数量修改为0,再重新输入.');</script>"
      response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_in_store.asp"">"
      response.end 
       
end if


set rs1=server.createobject("adodb.recordset") 
if guige="" then
sql1="select * from cailiao_store where pinming='"&pinming&"'and fenlei='"&fenlei1&"'"
else
sql1="select * from cailiao_store where pinming='"&pinming&"' and guige="&guige&" and houdou="&houdou&" and fenlei='"&fenlei1&"'"
end if
rs1.open sql1,conn,1,3
     
     
 if not rs1.eof then  
       if rs1("number")<num then
       response.Write "<script language=javascript>alert('对不起！原入库数已出库,库存数量不足！');</script>"
       response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_in_store.asp"">"
       response.end   
       else
       rs1("number")=rs1("number")-num
       end if
  end if
rs1.update
rs1.close
set rs1=nothing

if gonghuodanwei="吹膜"  then
set rs2=server.createobject("adodb.recordset") 
sql2="select * from fuliao_store"
rs2.open sql2,conn,2,3
rs2("number")=rs2("number")+(num-num/60)
rs2.update
rs2.close
set rs2=nothing
end if

set rs=server.createobject("adodb.recordset") 
sql="select * from cailiao_in_store where id="&request.Form("id")
rs.open sql,conn,1,3
rs("use_num") =use_num
rs("danjia")=danjia
rs("gonghuodanwei") = trim(request.Form("gonghuodanwei"))
rs("jinge")=use_num*danjia
rs.update
rs.close
set rs=nothing

response.Write "<script language=javascript>alert('修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_in_store.asp"">"
response.end

end sub   
 
  sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_in_store.asp"">"
  response.end
else
        conn.execute("delete from cailiao_in_store where id in ("&selBigClass&")and use_num=0")
		response.Write "<script language=javascript>alert('数量为0的已删除成功！\n\n数量不为0的没有被删除！');</script>"
		response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_in_store.asp"">"
		end if
response.end
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="cailiao" or oskey="admin" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from cailiao_in_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑入库信息</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加入库信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">入库日期</td>
            <td width="16%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">类 别</td>
			<td width="16%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">品 名</td>
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">宽&nbsp;&nbsp; 度</td>			
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">厚&nbsp;&nbsp; 度</td>			
			<td width="6%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">单价</td>			
			<td width="14%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">数量</td>			
			<td width="8%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">单位</td>
			
          </tr>
		  <form name="powersearch" method="post" action="cailiao_in_store.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:90'>
			<% if isedit then%>
			<option selected="selected"><%=rs("fenlei")%></option>
			<%else%>
			<option selected value='0000'>请选择类别</option>
			<%end if%>
			
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from fenlei",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("fenleiname")%>"><%=rs1("fenleiname")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='pinming' size=1 style='width:190'>
			<% if isedit then%>
			<option selected="selected" value="<%=rs("pinming")%>"><%=rs("pinming")%> </option>
			<input type="hidden" value="<%=rs("use_num")%>" name="use_num1">
			<input type="hidden" value="<%=rs("pinming")%>" name="pinming1">
			<input type="hidden" value="<%=rs("guige")%>" name="guige1">
			<input type="hidden" value="<%=rs("unit")%>" name="unit1">
			<input type="hidden" value="<%=rs("houdou")%>" name="houdou1">
			<%else%>
			<option>--请选择--</option>
			<%end if%>
			</select>
			</td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="guige"  size="1" style='width:54; height:18'>
			<% if isedit then%><option ><%=rs("guige")%></option><%end if%>
			<%
			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from guige order by guige",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option <% if isedit then
			if rs2("guige")=rs("guige") then
			%> selected="selected"<%end if%><%end if%>><%=rs2("guige")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%>
			</select><td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="houdou"  size="1" style='width:55; height:16'>
			<% if isedit then%><option ><%=rs("houdou")%></option><%end if%>
			<%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option <% if isedit then
			if rs4("houdou")=rs("houdou") then
			%> selected="selected"<%end if%><%end if%>><%=rs4("houdou")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
			%>
			</select><td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="danjia" type="text" style="ime-mode:disabled"  size="6" value='<% if isedit then
		                                                         response.write rs("danjia")
																  end if %>'></td>


			<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="use_num" type="text" style="ime-mode:disabled"  size="10" value='<% if isedit then
		                                                         response.write rs("use_num")
																  end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="unit"  size="1" style='width:67; height:18'>
			<% if isedit then%><option ><%=rs("unit")%></option><%end if%>
			<%
			
			set rs5=server.createobject("adodb.recordset")
		   rs5.open "select * from unit",conn,1,1
		   if not rs5.eof then
		   do while not rs5.eof
			%>
			<option <% if isedit then
			if trim(rs5("unit"))=trim(rs("unit")) then
			%> selected="selected"<%end if%><%end if%>><%=rs5("unit")%></option>
			<%
			rs5.movenext
			loop
			end if
			rs5.close
			set rs5=nothing
			%>
			</select></td>
			</td>
          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'" colspan="2">供货单位</td>
			<td  colspan="6" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">备  注</td>			
          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="2">
<select name="gonghuodanwei"  size="1" style='width:90; height:16'>
			<% if isedit then%>
			<option  selected="selected"><%=rs("gonghuodanwei")%></option>
			<%else%>
			<option selected value='0000'>请选择</option>
			<%end if%>

			<%
			
			set rs6=server.createobject("adodb.recordset")
		   rs6.open "select * from gonghuodanwei",conn,1,1
		   if not rs6.eof then
		   do while not rs6.eof
			%>
			<option <% if isedit then
			if trim(rs6("gonghuodanwei"))=trim(rs("gonghuodanwei")) then
			%> selected="selected"<%end if%><%end if%>><%=rs6("gonghuodanwei")%></option>
			<%
			rs6.movenext
			loop
			end if
			rs6.close
			set rs6=nothing
			%>
			</select>		</td>
<td height="18" colspan="6" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<textarea name="content" cols="40" rows="2"><% if isedit then
		                                                         response.write rs("content")
																  end if %></textarea></td></tr>
			<tr>
			<td colspan="8" align="center"><input type="submit"  value="<% if isedit then
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
<b><font color="#ffffff">入 库 材 料 列 表</font></b>

</td>
</tr>
</table>
			<form action="cailiao_in.asp" method="post" name="selform">
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">入库日期</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">类 别</td>
			<td width="16%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">品 名</td>
			<td width="9%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">规 格</td>
			<td width="14%" height="18" align="center"  class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">数量</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			单价</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">单 位</td>
			<td width="18%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			供货单位</td>
            <td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">操 作</td>
          </tr>
		 <%
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 200          '每页显示多少条
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet") 
search=request("search")

if search<>"" then 
gonghuodanwei1=request("gonghuodanwei")
keyword=request("keyword")
rs.Source="select* from cailiao_In_Store where gonghuodanwei like '%"&gonghuodanwei1&"%' and pinming like '%"&keyword&"%'and guige<>0 and use_num<>0 order by -id"
else
rs.Source="select* from cailiao_In_Store where guige<>0  and use_num<>0 order by -id"
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
<td height="18" align="center"><a href="cailiao_In_Store.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("uptime")%></a></td>
<td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center">&nbsp;<a href="cailiao_In_Store.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18" align="center"><%if rs("guige")=0 then%>&nbsp;<%else%><%=rs("guige")%>*<%=rs("houdou")%><%end if%></td>
<td height="18" align="center"><%=Formatnumber(rs("use_num"),2,-1,0,0)%><%if rs("use_num")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("danjia")%></td>
<td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%if rs("old_fenlei")<>"" then%>由<%=rs("old_fenlei")%><%=rs("end_time")%><%=rs("gonghuodanwei")%><%else%><%=rs("gonghuodanwei")%><%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="9" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="cailiao_in_store.asp?gonghuodanwei="&gonghuodanwei1&"&keyword="&keyword&"&search=search&"
			else
            url="cailiao_in_store.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="cailiao" or oskey="admin" then%>
     <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input type=submit value=打印 name=action> 
              


              
              <%end if%></td></tr>
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
<b><font color="#ffffff">入库搜索</font></b>

</td>
</tr>
</table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		           
<form name="search" method="post" action="cailiao_in_store.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">按品名查找： 
				
				<select name='keyword'  style='width:130'>
			<option selected value="">请选择品名</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pinming) from cailiao_in_store where guige<>0 order by pinming",conn,2,2
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
                      
			<td align="center" width="13%" ><input type="submit"  value="查 找"> </td>
			<td align="center" width="35%">按供应商查找： 
			<select name="gonghuodanwei"  style='width:130'>
			<option selected value="">请选择供货商</option>
			<%			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select distinct(gonghuodanwei) from cailiao_in_store where guige<>0 order by gonghuodanwei",conn,2,2
		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option value="<%=rs4("gonghuodanwei")%>"><%=rs4("gonghuodanwei")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
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
