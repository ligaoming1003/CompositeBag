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


Function gen_key(digits)
dim char_array(50)
char_array(0) = "0"
char_array(1) = "1"
char_array(2) = "2"
char_array(3) = "3"
char_array(4) = "4"
char_array(5) = "5"
char_array(6) = "6"
char_array(7) = "7"
char_array(8) = "8"
char_array(9) = "9"
randomize
do while len(output) < digits
cnum = char_array(Int((10 - 0 + 1) * Rnd + 0))
output = output + cnum
loop
gen_key = output
End Function


sub SaveAdd

Loan_Amount=request.Form("Loan_Amount")
pinming=trim(request.Form("pinming"))
use_dep=trim(request.Form("use_dep"))
comm=gen_key(9)

set rs2=server.createobject("adodb.recordset") 
sql2="select * from JS_out_store"
rs2.open sql2,conn,1,3
rs2.addnew

rs2("pinming") = pinming
rs2("uptime") = request.Form("uptime")
rs2("use_dep") = use_dep
rs2("comm") =comm
    if pinming="上期客户欠款" then
    rs2("Loan_Amount") =Loan_Amount
   else
    rs2("hkdz")=Loan_Amount
   end if
rs2.update
response.Write "<script language=javascript>alert('结算成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=JS_one.asp"">"
response.end
rs2.close
set rs2=nothing
end sub

 
sub SaveModify  
old_num=request.Form("old_num")
old_dep=request.Form("old_dep")
old_pinming=request.Form("old_pinming")
comm=request.Form("old_comm")
Loan_Amount=request.Form("Loan_Amount")
pinming=trim(request.Form("pinming"))
use_dep=trim(request.Form("use_dep"))

num= old_num-Loan_Amount

if old_dep<>use_dep then
response.Write "<script language=javascript>alert('对不起！客户名称不能修改！\n\n只能在数量栏中输入负数清零后，重新输入！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=JS_one.asp"">"
response.end
end if

if pinming="上期客户欠款" or pinming="上期客户长款" then
set rs2=server.createobject("adodb.recordset") 
sql2="select * from JS_out_store where use_dep='"&use_dep&"'and  comm="&comm&""
rs2.open sql2,conn,1,3
rs2("uptime") = request.Form("uptime")
rs2("pinming") = pinming
   if pinming="上期客户欠款" then
    rs2("Loan_Amount") =Loan_Amount
    rs2("hkdz")=0
   elseif pinming="上期客户长款" then
    rs2("hkdz")=Loan_Amount
    rs2("Loan_Amount") =0
   end if

rs2.update
response.Write "<script language=javascript>alert('结算修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=JS_one.asp"">"
response.end
rs2.close
set rs2=nothing
else
response.Write "<script language=javascript>alert('对不起！非结算项不能修改！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=JS_one.asp"">"
response.end
end if
end sub  
 
   sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=JS_one.asp"">"
  response.end
else
        conn.execute("delete from JS_out_store where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=JS_one.asp"">"
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
		   rs.open "select * from JS_out_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<font color="#ffffff">
<b>&nbsp;&nbsp; 编辑结算</b></font></td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;&nbsp; 添加</font></b><font color="#FFFFFF"><b>结算</b></font></td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			结算日期</td>
            <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			类 别</td>
			
			<td width="25%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">数量</td>
			
			<td width="25%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">客户名称</td>
			
          </tr>
		  <form name="powersearch" method="post" action="JS_one.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='pinming'  size="1" style='width:99; height:14'>
			<% if isedit then%>
		    			
            <option selected value='上期客户长款'>上期客户长款</option>	
            <option selected value='上期客户欠款'>上期客户欠款</option>
		    <option selected="selected"><%=rs("pinming")%></option>
				 <input type="hidden" value="<%=rs("Loan_Amount")%>" name="old_num">
			     <input type="hidden" value="<%=rs("use_dep")%>" name="old_dep">
                 <input type="hidden" value="<%=rs("comm")%>" name="old_comm"> 
                 <input type="hidden" value="<%=rs("pinming")%>" name="old_pinming">
                 			<%else%>
            <option selected value='上期客户长款'>上期客户长款</option>
            <option selected value='上期客户欠款'>上期客户欠款</option>			
            	

			<%end if%>

			
			
			</select></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="Loan_Amount" type="text" style="ime-mode:disabled"  size="19" value='<% if isedit then
                                                                 if rs("pinming")="上期客户长款" then
                                                                 response.write rs("hkdz")
                                                                 else
		                                                         response.write rs("Loan_Amount")
		                                                         end if
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
          </tr>
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



			<form action="JS.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;未 结 算 列 表</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="13%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">日期</td>
			<td width="33%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			款项类别</td>
			<td width="18%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			欠款金额</td>
			<td width="11%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			长款金额</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">客户名称</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">操 作</td>
          </tr>
		 <%
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 100          '每页显示多少条
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet") 
search=request("search")
if search<>"" then 
keyword=request("keyword")

rs.Source="select* from JS_out_store where pihao='ok' and use_dep like '%"&keyword&"%' order by -id"
else
rs.Source="select* from JS_out_store where pihao='ok' order by -id"
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
<td height="18" align="center"><a href="JS_one.asp?id=<%=rs("ID")%>&action=edit"><%=rs("uptime")%></a></td>
<td height="18" align="center">&nbsp;<a href="JS_one.asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18" align="center"><%if rs("Loan_Amount")="0" then%>&nbsp;<%else%><%=rs("Loan_Amount")%><%end if%></td>
<td height="18" align="center"><%if rs("hkdz")="0" then%>&nbsp;<%else%><%=rs("hkdz")%><%end if%></td>
<td height="18" align="center"><div title="点击打印"><a href="JS_in.asp?id=<%=rs("id")%>&action=edit"><%=rs("use_dep")%></a></div></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="center" colspan="6" height="22">
<p align="right">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="JS_one.asp?keyword="&keyword&"&search=search&"
			else
            url="JS_one.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="xiaoshou" or oskey="admin"  then%>
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input type=submit value=结算 name=action1> 
<%end if%></td></tr>
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
<b><font color="#ffffff">&nbsp; 结&nbsp; 算&nbsp; 搜&nbsp; 索</font></b>

</td>
</tr>
</table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="JS_one.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">按客户查找： 
				<select name="keyword" size="1">
			<option value="">请选择</option>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(use_dep) from JS_out_store order by use_dep",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   kehu=rs2("use_dep")
		   %>
		   <option value="<%=rs2("use_dep")%>"><%=rs2("use_dep")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
                      
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
