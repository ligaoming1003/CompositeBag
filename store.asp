<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>∷管理系统∷</title>
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

	if (powersearch.Loan_price.value == "")
	{
		alert("单价不能为空！");
		powersearch.Loan_price.focus();
		return (false);
	}
	if (powersearch.Loan_num.value == "")
	{
		alert("数量不能为空！");
		powersearch.Loan_num.focus();
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
rs1.open "select * from fenlei1",conn,1,1
if not rs1.eof then   
do while not rs1.eof
j=0	
%>
po_ca_show[<%=i%>] = '<%=rs1("fenlei")%>';
po_ca_value[<%=i%>] = '<%=rs1("fenlei")%>';
po_detail_show[<%=i%>] = new Array();
po_detail_value[<%=i%>] = new Array();

<%
fenlei=rs1("fenlei")
set rs2=server.createobject("adodb.recordset")

rs2.open "select * from chengping_store where fenlei='"&fenlei&"'and number<>0 order by -id",conn,1,1

if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("pinming")%> |<%=rs2("guige")%>|<%=rs2("kuang")%>| <%=Formatnumber(rs2("number"),2,-1,0,0)%> | <%=rs2("unit")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("kuang")%>|<%=Formatnumber(rs2("number"),2,-1,0,0)%>|<%=rs2("unit")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '暂无产品出库';
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
Loan_num=request.Form("Loan_num")
Loan_price= request.Form("Loan_price")
pinming=trim(request.Form("pinming"))
use_dep=trim(request.Form("use_dep"))
names=Split(pinming,"|")
id=names(0)
pinming=names(1)
guige=names(2)
kuang=names(3)
num=names(4)
unit=names(5)

if num-Loan_num<0 then
response.Write "<script language=javascript>alert('对不起，库存不足，请调整您的出库量！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_out_store.asp"">"
response.end
end if
if Loan_price>Loan_num then 
response.Write "<script language=javascript>alert('单价比数量还大，可能是你填写错误！\n\n请核对后修改！');</script>"
end if
 

set rs2=server.createobject("adodb.recordset") 
sql2="select * from JS_out_store"
rs2.open sql2,conn,1,3
rs2.addnew
rs2("fenlei") = trim(request.Form("d_position1"))
rs2("guige") = guige
rs2("kuang") = kuang
rs2("pinming") = pinming
rs2("uptime") = request.Form("uptime")
rs2("Unit") = unit
rs2("use_dep") = trim(request.Form("use_dep"))
rs2("Loan_num") = Loan_num
rs2("Loan_price") = request.Form("Loan_price")
rs2("Loan_Amount") = request.Form("Loan_num") * request.Form("Loan_price")
rs2("content") = trim(request.Form("content"))

rs2.update
rs2.close
set rs2=nothing

          set rs3=server.createobject("adodb.recordset")
          sql3="select * from JS_huzong where kehu='"&use_dep&"'"
          rs3.open sql3,conn,2,3
            if not rs3.eof then
            rs3("Loan_Amount")=rs3("Loan_Amount")+(Loan_num*Loan_price)
            else 
            rs3.addnew
             rs3("Loan_Amount")=Loan_num*Loan_price
             rs3("kehu") = trim(request.Form("use_dep"))            
            end if
         rs3.update
         rs3.close
         set rs3=nothing


set rs=server.createobject("adodb.recordset") 
sql="select * from chengping_out_store"
rs.open sql,conn,1,3
rs.addnew
rs("fenlei") = trim(request.Form("d_position1"))
rs("guige") = guige
rs("kuang") = kuang
rs("pinming") = pinming
rs("uptime") = request.Form("uptime")
rs("Unit") = unit
rs("use_dep") = trim(request.Form("use_dep"))
rs("Loan_num") = Loan_num
rs("Loan_price") = request.Form("Loan_price")
rs("Loan_Amount") = request.Form("Loan_num") * request.Form("Loan_price")
rs("content") = trim(request.Form("content"))
rs.update
rs.close
set rs=nothing

set rs1=server.createobject("adodb.recordset") 
sql1="select * from chengping_store where id="&id&""
rs1.open sql1,conn,1,3
if not rs1.eof then
rs1("number")=rs1("number")-Loan_num
rs1.update
end if
response.Write "<script language=javascript>alert('出库成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_out_store.asp"">"
response.end
rs1.close
set rs1=nothing
end sub

 
sub SaveModify  
old_num=request.Form("old_num")
old_dep=request.Form("old_dep")
old_price=request.Form("old_price")
Loan_num=request.Form("Loan_num")
pinming=trim(request.Form("pinming"))
fenlei=trim(request.Form("d_position1"))
use_dep=trim(request.Form("use_dep"))
Loan_price= request.Form("Loan_price")
names=Split(pinming,"|")
id=names(0)
pinming=names(1)
guige=names(2)
kuang=names(3)
unit=names(5)

num= old_num-Loan_num
onum=old_price-loan_price

if old_dep<>use_dep then
response.Write "<script language=javascript>alert('对不起，客户名不能修改！\n\n只能在数量栏中输入负数清零后，重新输入！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_out_store.asp"">"
response.end
end if



set rs=server.createobject("adodb.recordset") 
sql="select * from chengping_store where pinming='"&pinming&"'and fenlei='"&fenlei&"'and guige="&guige&"and kuang="&kuang&""
rs.open sql,conn,1,3
if not rs.eof then
if rs("number")+num<0 then
response.Write "<script language=javascript>alert('对不起，库存不足，请调整您的出库量！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_out_store.asp"">"
response.end
else
rs("number")=rs("number")+num
rs.update
end if
end if
rs.close
set rs=nothing


set rs2=server.createobject("adodb.recordset")

sql2="select * from JS_out_store where pinming='"&pinming&"'"

rs2.open sql2,conn,1,3
if not rs2.eof then

rs2("Loan_num") = Loan_num
rs2("Loan_Amount") = request.Form("Loan_num") * request.Form("Loan_price")
rs2("Loan_price") = Loan_price
rs2.update
end if
rs2.close
set rs2=nothing



         set rs3=server.createobject("adodb.recordset")
          sql3="select * from JS_huzong where kehu='"&use_dep&"'"
          rs3.open sql3,conn,2,3
          rs3("Loan_Amount")=rs3("Loan_Amount")-(onum*old_num+Loan_price*num)          
         rs3.update
         rs3.close
         set rs3=nothing

  
set rs=server.createobject("adodb.recordset") 
sql="select * from chengping_out_store where id="&request.Form("id")
rs.open sql,conn,1,3
rs("Loan_num") = Loan_num
rs("Loan_Amount") = request.Form("Loan_num") * request.Form("Loan_price")
rs("Loan_price") = Loan_price
rs.update
response.Write "<script language=javascript>alert('修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_out_store.asp"">"
response.end
rs.close
set rs=nothing
end sub   
   sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=changping_out_store.asp"">"
  response.end
else
        conn.execute("delete from chengping_out_store where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_out_store.asp"">"
response.end
end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="xiaoshou" or oskey="admin" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from chengping_out_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑销售信息</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加销售信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="8%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">出库日期</td>
            <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">类 别</td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2">
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp; 品 名 | 规&nbsp; 格 | 余数|&nbsp; 单位</td>
			
          </tr>
		  <form name="powersearch" method="post" action="chengping_out_store.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="8"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:59; height:19'>
			<% if isedit then%>
				<option selected="selected"><%=rs("fenlei")%></option>
			<%else%>
		        <option selected value='0000'>请选择类别</option>
			<%end if%>
			
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from fenlei1",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("fenlei")%>"><%=rs1("fenlei")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="2">
			<p align="left">
			<select name='pinming' size=1 style='width:674; height:18'>
			<% if isedit then%>
			<option selected="selected" value="<%=rs("id")%>|<%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("kuang")%>|<%=rs("Loan_num")%>|<%=rs("unit")%>"><%=rs("pinming")%> | <%=rs("guige")%>|<%=rs("kuang")%> | <%=rs("unit")%></option>
			<input type="hidden" value="<%=rs("Loan_num")%>" name="old_num">
			<input type="hidden" value="<%=rs("use_dep")%>" name="old_dep">
            <input type="hidden" value="<%=rs("Loan_price")%>" name="old_price">

			<%else%>
			<option>--请选择--</option>
			<%end if%>
			</select>
			</td>

          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" >单价</td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" >数量</td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">客户名称</td>			
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">备  注</td>			
          </tr>
		  <tr> 
 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
 <input name="Loan_price" type="text"  size="5" value='<% if isedit then
		                                                         response.write rs("Loan_price")
																  end if %>'></td>



 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="Loan_num" type="text"  size="8" value='<% if isedit then
		                                                         response.write rs("Loan_num")
																  end if %>'></td>



<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
 <select name='use_dep' size="1" style='width:130; height:18'>
			<% if isedit then%>
				<option selected="selected"><%=rs("use_dep")%></option>

			<%else%>
		        <option selected value='0000'>请选择客户</option>
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
			%>			</select></td>



<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<textarea name="content" cols="30" rows="2"><% if isedit then
		                                                         response.write rs("content")
																  end if %></textarea></td></tr>
			<tr>
			<td colspan="4" align="center"><input type="submit" value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        <%end if%>


</td>
</tr>
</table>
<br><br>
<TABLE width=820 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6" height="143">
<tr>
<td >



			<form action="chengping_out.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">产 品 销 售 列 表</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">出库日期</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">类 别</td>
			<td width="21%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">品 名</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">规 格</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">数 量</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">单 位</td>
			<td width="9%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">单价</td>
			<td width="11%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">金额</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">客户名称</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">操 作</td>
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
keyword=request("keyword")
akeyword=request("akeyword")
rs.Source="select* from chengping_out_store where pinming like '%"&keyword&"%' and use_dep like '%"&akeyword&"%' and loan_num<>0 order by -id"
else
rs.Source="select* from chengping_out_store where loan_num<>0 order by -id"
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
<td height="18" align="center"><a href="chengping_out_store.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("uptime")%></a></td>
<td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center">&nbsp;<a href="chengping_out_store.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%><%if rs("content")<>"" then%>(<%=rs("content")%>)<%end if%></a></td>
<td height="18"  align="center"><%=rs("guige")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=Formatnumber(rs("Loan_num"),2,-1,0,0)%><%if rs("Loan_num")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=Formatnumber(rs("Loan_price"),4,-1,0,0)%><%if rs("Loan_price")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=Formatnumber(rs("Loan_Amount"),2,-1,0,0)%><%if rs("Loan_Amount")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("use_dep")%><%if rs("use_dep")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="center" colspan="10" height="22">
<p align="right">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="chengping_out_store.asp?keyword="&keyword&"&akeyword="&akeyword&"&search=search&"
			else
            url="chengping_out_store.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper"  then%>
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input onClick="{if(confirm('确定要执行删除吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
              <input type="Hidden" name="action" value='del'><%end if%></td></tr>
        </table>

	  
</td>
    </tr>

  </table>


</form>

</td>
</tr>
</table>
<br><br>
<TABLE width=820 align=center border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6" height="92">
<tr>
<td>
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">销售搜索</font></b>

</td>
</tr>
</table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="chengping_out_store.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">按品名查找： 
				<input name="keyword" type="text"  size="23"></td>
                      
			<td align="center" width="13%" ><input type="submit"  value="查 找"> </td>
			<td align="center" width="35%">按客户查找： 
			<input name="akeyword" type="text"  size="22"></td>
                      
			<td align="center" width="17%" ><input type="submit"  value="查 找"> </td>

          </tr></form>
			</TD>
              </TR>
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