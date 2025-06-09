<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>∷嘉美管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript src="css/init.js"></SCRIPT>

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
	if (powersearch.bizhong.value == "")
	{
		alert("比重不能为空！");
		powersearch.bizhong.focus();
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
<!--#include file="top.asp"--><% dim sql,rs

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
set rs=server.createobject("adodb.recordset") 
sql="select * from cailiao where fenlei='"&trim(request.Form("fenlei"))&"' and pinming='"&trim(request.Form("pinming"))&"'"
rs.open sql,conn,1,3
if not rs.eof then
response.Write "<script language=javascript>alert('对不起，你现在添加的原料名称已经存在！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao.asp"">"
response.end
else
rs.addnew
rs("fenlei") = trim(request.Form("fenlei"))
rs("pinming") = trim(request.Form("pinming"))
rs("bizhong") = request.Form("bizhong")

rs.update
response.Write "<script language=javascript>alert('添加成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao.asp"">"
response.end
end if
rs.close
set rs=nothing
end sub

 
sub SaveModify   
set rs=server.createobject("adodb.recordset") 
sql="select * from cailiao where id="&request.Form("id")
rs.open sql,conn,1,3
rs("fenlei") = trim(request.Form("fenlei"))
rs("pinming") = trim(request.Form("pinming"))
rs("bizhong") = request.Form("bizhong")

rs.update
rs.close
set rs=nothing
response.Write "<script language=javascript>alert('修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao.asp"">"
response.end
end sub   
 
     sub delCate()
   selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao.asp"">"
  response.end
else
        conn.execute("delete from cailiao where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao.asp"">"
response.end
end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=400 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"><%if oskey="supper" or oskey="cailiao" or oskey="admin" then%> 

	<%
	    set rs=server.createobject("adodb.recordset")
	   if isedit then
		   rs.open "select * from cailiao where id=" & request("id"),conn,1,1
%>

    <table width="400" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编 辑 品 名 信 息</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="400" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添 加 品 名 信 息</font></b>

</td>
</tr>
</table>
<%end if %>
       <table width="400" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="100" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">类  别</td>
            <td width="150" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">品  名</td>


			
            <td width="150" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			比重</td>


			
          </tr>
		  <form name="powersearch" method="post" action="cailiao.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name="fenlei" onchange='Do_po_Change(this);' valign=top style='width:90'>
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
<input name="pinming" type="text"  size="19" value='<% if isedit then
		                                               response.write rs("pinming")
                                                       end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="bizhong" type="text"  size="18" value='<% if isedit then
		                                                         response.write rs("bizhong")
																  end if %>'></td>


          </tr>
		  <tr> 
            
			<td align="center" colspan="3" height="25"><input type="submit"  value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        <%end if%></td>
</tr>
</table>
<br><br>
<TABLE width=400 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6">
<tr>
<td >


			<form action="cailiao.asp" method="post" name="selform" >
  <table width="400"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="400" align="center" valign="top">
	        <table width="400" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">品 名 明 细 列 表</font></b>

</td>
</tr>
</table>
	 
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="100" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">类  别</td>
                        <td width="125" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">品  名</td>

                        <td width="125" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
						比重</td>

                        <td width="50" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">操 作</td>
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
rs.Source="select* from cailiao where fenlei like '%"&fenlei1&"%' and pinming like '%"&keyword&"%' order by id desc"
else
rs.Source="select* from cailiao order by id desc"
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
<td height="18" align="center"><%=rs("fenlei")%></td>
<td height="18" align="center">&nbsp;&nbsp;<a href="cailiao.asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>

<td height="18" align="center"><%=rs("bizhong")%></td>

<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if

%>
<tr><td align="right" colspan="4" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="cailiao.asp?fenlei="&fenlei1&"&keyword="&keyword&"&search=search&"
			else
            url="cailiao.asp?"
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
%>
&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" then%>
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
<br><br>
<TABLE width=500 align=center border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>

	        <table width="500" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">品 名 信 息 搜 索</font></b>

</td>
</tr>
</table>



        <table width="500" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="cailiao.asp">
		  <tr><input name="search" type="hidden" value="search">
            <td align="center" width="20%" height="40" valign="middle">
			<select name="fenlei" size="1">
			<option value="">所有品名</option>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from fenlei",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   fenleiname=rs2("fenleiname")
		   %>
		   <option value="<%=rs2("fenleiname")%>"><%=rs2("fenleiname")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
            <td align="center" width="50%">品名关键字： <input name="keyword" type="text"  size="25"></td>
                      
			<td align="center" width="30%" ><input type="submit"  value="查 找"> </td>
          </tr></form>
		 </fieldset> 
           </table>
       
			</TD>
              </TR>
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


