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

 
sub SaveAdd

end sub

sub SaveModify 
end sub   
  sub delCate() 

        conn.execute("delete from caiyin_in")
        conn.execute("delete from caiyin")
		response.Write "<script language=javascript>alert('已删除成功！');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_del.asp"">"
        response.end

    end sub

  %> <% sub myform(isEdit) %> 	  

<br><br>
<TABLE width=820 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6" height="143">
<tr>
<td >



			<form action="caiyin_del.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">印&nbsp; 刷&nbsp; 明&nbsp; 细&nbsp; 列&nbsp; 表</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="8%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">出库日期</td>
            <td width="4%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">类 别</td>
			<td width="31%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">产品名称</td>
			<td width="8%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">规 格</td>
			<td width="11%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" colspan="2">
			领&nbsp;&nbsp;&nbsp;&nbsp; 用</td>
			<td width="7%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">出库数</td>
			<td width="7%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">现存</td>
			<td width="5%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">退库</td>
			<td width="6%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">损耗率</td>
			<td width="7%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">领用车间</td>
            <td width="6%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">操 作</td>
          </tr>
          <tr class="but"> 
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			公斤</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			千米</td>
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

keyword=request("keyword")
rs.Source="select* from caiyin where cname like '%"&keyword&"%' order by -id"
else
rs.Source="select* from caiyin order by -id"
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
<td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("cname")%></td>
<td height="18" align="center"><%=rs("guige")%>*<%=rs("houdou")%></td>
<td height="18" align="center"><%if rs("Loan_num")="" then%>&nbsp;<%else%><%=rs("Loan_num")%><%end if%></td>
<td height="18" align="center"><%if rs("zhemi")="" then%>&nbsp;<%else%><%=Formatnumber(rs("zhemi"),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%if rs("Loan_oan")="0" then%><font color="#FF0000">在印</font><%else%><%=rs("Loan_oan")%>km<%end if%></td>
<td height="18" align="center" bgcolor="#FFFFCC"><%if rs("Loan_oan")="0" then%>&nbsp;<%else%><%=Formatnumber(rs("qianmi"),2,-1,0,0)%>km<%end if%></td>
<td height="18" align="center"><%if rs("oan_num")="0" then%>&nbsp;<%else%><%=rs("oan_num")%>kg<%end if%></td>
<td height="18" align="center"><%if rs("Loan_oan")="0" then%>&nbsp;<%elseif rs("zhemi")="0" then%>&nbsp;<%else%><%=Formatnumber(rs("qianmi")/rs("zhemi")*100,2,-1,0,0)%>%<%end if%></td>
<td height="18" align="center"><%if rs("use_dep")="" then%>&nbsp;<%else%><%=rs("use_dep")%><%end if%></td>
<td height="18" align="center"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="center" colspan="12" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="caiyin_del.asp?keyword="&keyword&"&search=search&"
			else
            url="caiyin_del.asp?"
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
        全部删除→
              <input onClick="{if(confirm('确定要删除吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
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
<TABLE width=820 align=center border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">出库搜索</font></b>

</td>
</tr>
</table>

        <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="caiyin_del.asp">
		  <tr><input name="search" type="hidden" value="search">
            <td height="40" valign="middle" align="center" width="20%">
			　</td>
            <td align="center" width="65%">品名关键字： <input name="keyword" type="text"  size="25"> 关键字为空则搜索所有</td>
                      
			<td align="center" width="15%" ><input type="submit"  value="查 找"> </td>
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
