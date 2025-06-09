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
		alert("品名不能为空！");
		powersearch.pinming.focus();
		return (false);
	}
	if (powersearch.number.value == "")
	{
		alert("数量不能为空！");
		powersearch.number.focus();
		return (false);
	}
	if (powersearch.kuang.value == "")
	{
		alert("规格不能为空！");
		powersearch.kuang.focus();
		return (false);
	}
	if (powersearch.guige.value == "")
	{
		alert("规格不能为空！");
		powersearch.guige.focus();
		return (false);
	}
	if (powersearch.pihao.value == "")
	{
		alert("来源不能为空！");
		powersearch.pihao.focus();
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
number=request.Form("use_num")
pinming=trim(request.Form("pinming"))

           set rs2=server.createobject("adodb.recordset") 
           sql2="select * from chengping_in_store"
           rs2.open sql2,conn,1,3
           rs2.addnew
           rs2("fenlei") =trim(request.Form("d_position1"))
           rs2("unit")=trim(request.Form("unit"))
           rs2("use_num")=number
           rs2("guige") =request.Form("guige")
           rs2("pinming") =pinming
           rs2("uptime") =request.Form("uptime")
           rs2("kuang")=request.Form("kuang")
           rs2("pihao")=trim(request.Form("pihao"))
           rs2.update
           rs2.close
           set rs2=nothing           


set rs=server.createobject("adodb.recordset") 
sql="select * from chengping_store"
rs.open sql,conn,1,3
rs.addnew
rs("fenlei") = trim(request.Form("d_position1"))
rs("pinming") = pinming
rs("uptime") = request.Form("uptime")
rs("pihao") = trim(request.Form("pihao"))
rs("number") = number
rs("kuang") = request.Form("kuang")
rs("guige") = request.Form("guige")
rs("unit")=trim(request.Form("unit"))
rs.update
response.Write "<script language=javascript>alert('结存入库成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_qita.asp"">"
response.end
rs.close
set rs=nothing
end sub

 
sub SaveModify  
number=request.Form("use_num")
pinming=trim(request.Form("pinming"))
old_num=request.Form("old_num")
fenlei=trim(request.Form("d_position1"))
old_fenlei=trim(request.Form("old_fenlei"))
old_pinming=trim(request.Form("old_pinming"))
num=number-old_num

if old_fenlei<>fenlei or old_pinming<>pinming then
response.Write "<script language=javascript>alert('类别、名称不能修改！\n\n\请先将数量改为0后，再重新录入！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_qita.asp"">"
  response.end

end if


           set rs2=server.createobject("adodb.recordset") 
           sql2="select * from chengping_in_store where id="&request.Form("id")
           rs2.open sql2,conn,1,3
           rs2("fenlei") =fenlei
           rs2("unit")=trim(request.Form("unit"))
           rs2("use_num")=number
           rs2("guige") =request.Form("guige")
           rs2("pinming") =pinming
           rs2("uptime") =request.Form("uptime")
           rs2("kuang")=request.Form("kuang")
           rs2("pihao")=trim(request.Form("pihao"))
           rs2.update
           rs2.close
           set rs2=nothing           


set rs=server.createobject("adodb.recordset") 
sql="select * from chengping_store where pinming='"&pinming&"'and fenlei='"&fenlei&"'"
rs.open sql,conn,1,3
           rs("fenlei") = fenlei
           rs("pinming") = pinming
           rs("uptime") = request.Form("uptime")
           rs("number")=rs("number")+num
           rs("kuang") = request.Form("kuang")
           rs("guige") = request.Form("guige")
           rs("unit")=trim(request.Form("unit"))
        
rs.update
response.Write "<script language=javascript>alert('结存入库修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_qita.asp"">"
response.end
rs.close
set rs=nothing
end sub   
   sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_qita.asp"">"
  response.end
else
        conn.execute("delete from chengping_store where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_qita.asp"">"
response.end
end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="chengpin" or oskey="admin"  then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from chengping_in_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑结存入库</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加</font></b><font color="#FFFFFF"><b>结存入库</b></font></td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			入库日期</td>
            <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">类 别</td>
			<td width="19%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			<p>品 名 </td>
			<td width="6%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			长度</td>
			
			<td width="5%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			宽度</td>
			
			<td width="13%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">数量</td>
			
			<td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			单位</td>
			
			<td width="25%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			来源</td>
			
          </tr>
		  <form name="powersearch" method="post" action="chengping_qita.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>">
		        <input type="hidden" value="<%=rs("use_num")%>" name="old_num"> 
		        <input type="hidden" value="<%=rs("fenlei")%>" name="old_fenlei">
		        <input type="hidden" value="<%=rs("pinming")%>" name="old_pinming">
		   <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:68; height:16'>
			<% if isedit then%>
				<option selected="selected"><%=rs("fenlei")%></option> 
				
			<%else%>
		       <option selected value='0000'>请选择</option>
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
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">

			<input name="pinming" type="text"  size="21" value='<% if isedit then
		                                                         response.write rs("pinming")
																  end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="kuang" type="text" style="ime-mode:disabled"  size="7" value='<% if isedit then
		                                                         response.write rs("kuang")
																  end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="guige" type="text" style="ime-mode:disabled"  size="7" value='<% if isedit then
		                                                         response.write rs("guige")
																  end if %>'>
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="use_num" type="text" style="ime-mode:disabled"  size="18" value='<% if isedit then
		                                                         response.write rs("use_num")
																  end if %>'></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="unit"  size="1" style='width:66; height:16'>
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
			</select>			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">

<input name="pihao" type="text"  size="21" value='<% if isedit then
		                                                         response.write rs("pihao")
																  end if %>'></td>
          </tr>
			<tr>
			<td colspan="8" align="center"><input type="submit" value="<% if isedit then
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
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">

<br><br>
<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">库 存 成 品 列 表</font></b>

</td>
</tr>
</table>
			<form action="chengping_qita.asp" method="post" name="selform" >
  <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">
	
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
          	<td width="10%" height="18" align="center"  class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">日 期</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">类 别</td>
			<td width="25%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">品 名</td>
			<td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">规 格</td>
			<td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">库存数量</td>
			<td width="10%" height="18" align="center"  class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">单 位 </td>

            <td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">操 作</td>
          </tr>
		 <%
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 20          '每页显示多少条
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet")
search=request("search")
if search<>"" then 
fenlei1=request("fenlei")
keyword=request("keyword")
rs.Source="select* from chengping_in_Store where use_num<>0 and fenlei like '%"&fenlei1&"%' and pinming like '%"&keyword&"%' and pihao<>''  order by -id"
else
rs.Source="select* from chengping_in_Store where use_num<>0 and pihao<>'' order by -id"
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
<td height="18" align="center"><%=rs("uptime")%><%if rs("uptime")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td height="18"  align="center"><a href="chengping_qita.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18" align="center"><%=rs("guige")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=Formatnumber(rs("use_num"),2,-1,0,0)%><%if rs("use_num")="0" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>

<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="10" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="chengping_qita.asp?fenlei="&fenlei1&"&keyword="&keyword&"&search=search&"
			else
            url="chengping_qita.asp?"
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
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">库存搜索</font></b>

</td>
</tr>
</table>

        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="chengping_qita.asp">
		  <tr><input name="search" type="hidden" value="search">
            <td  height="40" valign="middle" align="center" width="20%">
			</td>
            <td align="center" width="65%">品名关键字： <input name="keyword" type="text"  size="25"> 关键字为空则搜索所有</td>
                      
			<td align="center" width="15%" ><input type="submit"  value="查 找"> </td>
          </tr></form>
		  
           </table>
        
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
