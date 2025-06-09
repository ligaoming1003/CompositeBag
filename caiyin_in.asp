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
  response.Write "<script language=javascript>alert('此页面只供修改！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
  response.end
end sub

 
sub SaveModify
old_dep=request.Form("old_dep")
old_num=request.Form("old_num")
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))
use_dep=trim(request.Form("use_dep"))
unit=trim(request.Form("unit"))
names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)
houdou=names(3)
pihao=names(4)
num=old_num-Loan_oan


if use_dep="复合车间" and old_dep="复合车间" then
         set rs=server.createobject("adodb.recordset") 
          sql="select * from fuhe where cname='"&cname&"'"
          rs.open sql,conn,2,3
          if not rs.eof then
         rs("guige") = guige
         rs("houdou")= houdou
         rs("m_Loan_num")=rs("m_Loan_num")-num
          end if
         rs.update
         rs.close
         set rs=nothing
elseif use_dep="熟化室"  and old_dep="熟化室" then  
             set rs5=server.createobject("adodb.recordset") 
             sql5="select * from matur where cname='"&cname&"'"
             rs5.open sql5,conn,2,3
             if not rs5.eof and not rs5.bof then
             rs5("number")=rs5("number")-num
             end if
             rs5.update 
             rs5.close
             set rs5=nothing
elseif use_dep="熟化室"  and old_dep="复合车间" then
         set rs=server.createobject("adodb.recordset") 
          sql="select * from fuhe where cname='"&cname&"'"
          rs.open sql,conn,2,3
          if not rs.eof then
         rs("guige") = guige
         rs("m_Loan_num")=rs("m_Loan_num")-old_num
          end if
         rs.update
         rs.close
         set rs=nothing
             set rs5=server.createobject("adodb.recordset") 
             sql5="select * from matur where cname='"&cname&"'"
             rs5.open sql5,conn,2,3
             if not rs5.eof and not rs5.bof then
             rs5("number")=rs5("number")+Loan_oan
             else
              rs5.addnew
              rs5("cname") = cname
              rs5("guige") = guige
              rs5("number")=loan_oan

              rs5("uptime")=trim(request.Form("uptime"))

             end if
             rs5.update 
             rs5.close
             set rs5=nothing
elseif use_dep="复合车间"  and old_dep="熟化室" then
         set rs=server.createobject("adodb.recordset") 
          sql="select * from fuhe where cname='"&cname&"'"
          rs.open sql,conn,2,3
         if not rs.eof and not rs.bof then
         rs("guige") = guige
         rs("m_Loan_num")=rs("m_Loan_num")+Loan_oan
         else 
         rs.addnew
         rs("m_Loan_num")=Loan_oan
         rs("fenlei") = fenlei
         rs("guige") = guige
         rs("houdou")= houdou
         rs("pinming") = pinming
         rs("uptime") = request.Form("uptime")         
         rs("cname")= cname


          end if
         rs.update
         rs.close
         set rs=nothing
             set rs5=server.createobject("adodb.recordset") 
             sql5="select * from matur where cname='"&cname&"'"
             rs5.open sql5,conn,2,3
             if not rs5.eof and not rs5.bof then
             rs5("number")=rs5("number")-old_num
             end if
             rs5.update 
             rs5.close
             set rs5=nothing

             
else 
             response.Write "<script language=javascript>alert('领用车间出错！');</script>"
             response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
             response.end 
            
    

end if
          set rs1=server.createobject("adodb.recordset") 
          sql1="select * from caiyin_in where id="&request.Form("id")
          rs1.open sql1,conn,1,3
          if not rs1.eof then
           rs1("unit")=unit
           rs1("loan_oan")=Loan_oan
           rs1("guige") =guige
           rs1("cname") =cname
           rs1("uptime") = request.Form("uptime") 
           rs1("houdou")=houdou

           rs1("use_dep")=use_dep
          end if  
           rs1.update
           rs1.close
           set rs1=nothing

          set rs3=server.createobject("adodb.recordset") 

           sql3="select * from caiyin where cname='"&cname&"'"

           rs3.open sql3,conn,1,3
          if not rs3.eof and not rs3.bof then
           rs3("loan_oan")=rs3("loan_oan")-num
           rs3("qianmi")=rs3("qianmi")+num
           rs3("use_dep")=use_dep
           end if
           rs3.update
             response.Write "<script language=javascript>alert('修改成功！');</script>"
             response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
             response.end 
           rs3.close
           set rs3=nothing 


 
 
end sub   
 
sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
  response.end
else
        conn.execute("delete from caiyin_in where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
response.end
end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin" then%>
	
	<%
	    
	   if isedit then

	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from caiyin_in where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑彩印车间出库信息</font></b>

</td>
</tr>
</table>
<%else%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加彩印车间出库信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="19%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">出库日期</td>
			<td width="60%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'"><p align="left">&nbsp;&nbsp;&nbsp;&nbsp; 品 名 | 规&nbsp; 格 | 余数|&nbsp; 单位|&nbsp; 批号</td>
			<td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">领用车间</td>
			
			<td width="9%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			单位</td>
			
          </tr>
		  <form name="powersearch" method="post" action="caiyin_in.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='cname' size=1 style='width:481; height:15'>
			<% if isedit then%>
			<option selected="selected" value="0|<%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("pihao")%>"><%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("pihao")%></option>
			<input type="hidden" value="<%=rs("Loan_oan")%>" name="old_num">
             <input type="hidden" value="<%=rs("use_dep")%>" name="old_dep">
			<%else%>
			<option>--请选择--</option>
			<%end if%>
			</select>
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="use_dep"  size="1" style='width:68; height:16'>
			<% if isedit then%><option ><%=rs("use_dep")%></option><%else%>
						<option selected value='熟化室'>熟化室</option>
			<option selected value='退库'>退库</option>
			<option selected value='复合车间'>复合车间</option><%end if%>
			</select></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="unit"  size="1" style='width:50; height:15'>
			<% if isedit then%><option ><%=rs("unit")%></option><%end if%>
			<option selected value='公斤'>公斤</option>
			<option selected value='千米'>千米</option>

			</select></td>
          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			数&nbsp;&nbsp; 量</td>
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="3" rowspan="2">
			<p align="left">说明:1、出库输入有误，点击“品名”链接进行修改。<br>
&nbsp;&nbsp;&nbsp;&nbsp; 2、退库输入有误只能输入负数清零后，再重新输入。</td>
          </tr>
		  <tr> 
 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
 <input name="Loan_oan" type="text" style="ime-mode:disabled" size="11" value='<% if isedit then
		                                                         response.write rs("Loan_oan")
																  end if %>'><font color="#FFFFFF">千米(公斤)</font></td>



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
<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >

  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">彩 印 出 库 列 表</font></b>

</td>
</tr>
</table>
			<form action="caiyin_in.asp" method="post" name="selform" >
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">入库日期</td>
			<td width="52%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">品 名</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">规 格</td>
			<td width="8%" height="18" align="center"  class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">数量</td>
			<td width="4%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">单 位</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">车间</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			员工</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">操作</td>
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
keyword=request("keyword")

rs.Source="select* from caiyin_in where cname like '%"&keyword&"%' order by -id "
else
rs.Source="select* from caiyin_In order by -id"
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
<td height="18" align="center"><%=rs("uptime")%></td>
<td height="18" align="center"><a href="caiyin_in.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("cname")%></a></td>
<td height="18" align="center">&nbsp;<%=rs("guige")%>*<%=rs("houdou")%></td>
<td height="18" align="center"><%if rs("loan_oan")="0" then%>&nbsp;<%else%><%=rs("loan_oan")%><%end if%></td>
<td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("use_dep")%></td>
<td height="18" align="center"><%=rs("pihao")%></td>
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
			url="caiyin_In.asp?keyword="&keyword&"&search=search&"
			else
            url="caiyin_In.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" then%>
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input onClick="{if(confirm('确定要删除该记录吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
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
<b><font color="#ffffff">入库搜索</font></b>

</td>
</tr>
</table>

        <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="caiyin_in.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="84%">按品名查找： 
			<input name='keyword' type="text"  size="22"></td>
                      
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