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


  %> <% sub myform(isEdit) %> 	  

<br><br>
<TABLE width=820 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6" height="143">
<tr>
<td >



			<form action="JS1.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;已 结 算 列 表</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="13%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">日期</td>
			<td width="33%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">款项类别</td>
			<td width="18%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">欠款金额</td>
			<td width="11%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">长款金额</td>
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

rs.Source="select* from JS_out_store where pihao<>'ok'and use_dep like '%"&keyword&"%' order by -id"
else
rs.Source="select* from JS_out_store where pihao<>'ok'  order by -id"
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
<td height="18" align="center"><a href="JS_one1.asp?id=<%=rs("ID")%>&action=edit"><%=rs("uptime")%></a></td>
<td height="18" align="center">&nbsp;<a href="JS_one1.asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18" align="center"><%if rs("Loan_Amount")="0" then%>&nbsp;<%else%><%=rs("Loan_Amount")%><%end if%></td>
<td height="18" align="center"><%if rs("hkdz")="0" then%>&nbsp;<%else%><%=rs("hkdz")%><%end if%></td>
<td height="18" align="center"><a href="JS_in.asp?id=<%=rs("id")%>&action=edit"><%=rs("use_dep")%></a></td>
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
			url="JS_one1.asp?keyword="&keyword&"&search=search&"
			else
            url="JS_one1.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="xiaoshou" or oskey="admin" then%>
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input type=submit value=返回 name=action1> 
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
<font color="#FFFFFF"><b>&nbsp;&nbsp;&nbsp; 已</b></font><b><font color="#ffffff">结算项搜索</font></b>

</td>
</tr>
</table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="JS_one1.asp">
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
