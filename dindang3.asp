<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>∷嘉美管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>
<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp; 生 产 进 度</font></b>

</td>
</tr>
</table>
			
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">日期</td>
            <td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">客户</td>
			<td width="27%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">品 名</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">规 格</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">材料</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">规格</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">需要量</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">安排量</td>
            <td width="4%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			彩印</td>
            <td width="4%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			复合</td>
            <td width="4%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			成品</td>
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
akeyword=request("akeyword")
keyword=request("keyword")
rs.Source="select* from dindang where kehu like '%"&akeyword&"%' and pinming like '%"&keyword&"%' order by -id"
else
rs.Source="select* from dindang  order by -id"
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

<%if rs("name1")<>"0"  and rs("name2")<>"0" and rs("name3")<>"0"  and rs("name4")<>"0" then%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="72" align="center" rowspan="4"><%=rs("uptime")%></td>
<td height="72" align="center" rowspan="4"><%=rs("kehu")%><%if rs("kehu")="" then%>&nbsp;<%end if%></td>
<td height="72" align="center" rowspan="4"><%=rs("pinming")%></td>
<td height="72" align="center" rowspan="4"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=rs("name1")%></td>
<td height="18" align="center"><%=rs("nianbang")%>*<%=rs("cunshu")%></td>
<td height="18" align="center"><%=rs("nianti")%><%=rs("unit1")%></td>
<td height="18" align="center"><%if rs("anpai1")="0" then%>&nbsp;<%else%><%=rs("anpai1")%><%end if%></td>
<td height="72" align="center" rowspan="4" <%if rs("caiyin")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<td height="72" align="center" rowspan="4"<%if rs("fuhe")="ok" then%>bgcolor="#FFFF00"<%end if%> >　　</td>
<td height="72" align="center" rowspan="4"<%if rs("chengping")="ok" then%>bgcolor="#FFFF00"<%end if%> >　　</td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name2")%></td>
<td height="18" align="center"><%=rs("nianbang2")%>*<%=rs("cunshu2")%></td>
<td height="18" align="center"><%=rs("nianti2")%><%=rs("unit2")%></td>
<td height="18" align="center"><%if rs("anpai2")="0" then%>&nbsp;<%else%><%=rs("anpai2")%><%end if%></td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name3")%></td>
<td height="18" align="center"><%=rs("nianbang3")%>*<%=rs("cunshu3")%></td>
<td height="18" align="center"><%=rs("nianti3")%><%=rs("unit3")%></td>
<td height="18" align="center"><%if rs("anpai3")="0" then%>&nbsp;<%else%><%=rs("anpai3")%><%end if%></td>
</tr>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name4")%></td>
<td height="18" align="center"><%=rs("nianbang4")%>*<%=rs("cunshu4")%></td>
<td height="18" align="center"><%=rs("nianti4")%><%=rs("unit4")%></td>
<td height="18" align="center"><%if rs("anpai4")="0" then%>&nbsp;<%else%><%=rs("anpai4")%><%end if%></td>
			</tr>
</tr>
<%elseif rs("name1")<>"0"  and rs("name2")<>"0" and rs("name3")="0"  and rs("name4")<>"0" then%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="54" align="center" rowspan="3"><%=rs("uptime")%></td>
<td height="54" align="center" rowspan="3"><%=rs("kehu")%><%if rs("kehu")="" then%>&nbsp;<%end if%></td>
<td height="54" align="center" rowspan="3"><%=rs("pinming")%></td>
<td height="54" align="center" rowspan="3"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=rs("name1")%></td>
<td height="18" align="center"><%=rs("nianbang")%>*<%=rs("cunshu")%></td>
<td height="18" align="center"><%=rs("nianti")%><%=rs("unit1")%></td>
<td height="18" align="center"><%if rs("anpai1")="0" then%>&nbsp;<%else%><%=rs("anpai1")%><%end if%></td>
<td height="54" align="center" rowspan="3" <%if rs("caiyin")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<td height="54" align="center" rowspan="3" <%if rs("fuhe")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<td height="54" align="center" rowspan="3" <%if rs("chengping")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name2")%></td>
<td height="18" align="center"><%=rs("nianbang2")%>*<%=rs("cunshu2")%></td>
<td height="18" align="center"><%=rs("nianti2")%><%=rs("unit2")%></td>
<td height="18" align="center"><%if rs("anpai2")="0" then%>&nbsp;<%else%><%=rs("anpai2")%><%end if%></td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name4")%></td>
<td height="18" align="center"><%=rs("nianbang4")%>*<%=rs("cunshu4")%></td>
<td height="18" align="center"><%=rs("nianti4")%><%=rs("unit4")%></td>
<td height="18" align="center"><%if rs("anpai4")="0" then%>&nbsp;<%else%><%=rs("anpai4")%><%end if%></td>
			</tr>
</tr>
<%elseif rs("name1")<>"0"  and rs("name2")="0" and rs("name3")="0"  and rs("name4")<>"0" then%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="36" align="center" rowspan="2"><%=rs("uptime")%></td>
<td height="36" align="center" rowspan="2"><%=rs("kehu")%><%if rs("kehu")="" then%>&nbsp;<%end if%></td>
<td height="36" align="center" rowspan="2"><%=rs("pinming")%></td>
<td height="36" align="center" rowspan="2"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=rs("name1")%></td>
<td height="18" align="center"><%=rs("nianbang")%>*<%=rs("cunshu")%></td>
<td height="18" align="center"><%=rs("nianti")%><%=rs("unit1")%></td>
<td height="18" align="center"><%if rs("anpai1")="0" then%>&nbsp;<%else%><%=rs("anpai1")%><%end if%></td>
<td height="36" align="center" rowspan="2" <%if rs("caiyin")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<td height="36" align="center" rowspan="2" <%if rs("fuhe")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<td height="36" align="center" rowspan="2" <%if rs("chengping")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name4")%></td>
<td height="18" align="center"><%=rs("nianbang4")%>*<%=rs("cunshu4")%></td>
<td height="18" align="center"><%=rs("nianti4")%><%=rs("unit4")%></td>
<td height="18" align="center"><%if rs("anpai4")="0" then%>&nbsp;<%else%><%=rs("anpai4")%><%end if%></td>
			</tr>
</tr>
<%elseif rs("name1")<>"0"  and rs("name2")="0" and rs("name3")="0"  and rs("name4")="0" then%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("uptime")%></td>
<td height="18" align="center"><%=rs("kehu")%><%if rs("kehu")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("pinming")%></td>
<td height="18" align="center"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=rs("name1")%></td>
<td height="18" align="center"><%=rs("nianbang")%>*<%=rs("cunshu")%></td>
<td height="18" align="center"><%=rs("nianti")%><%=rs("unit1")%></td>
<td height="18" align="center"><%if rs("anpai1")="0" then%>&nbsp;<%else%><%=rs("anpai1")%><%end if%></td>
<td height="18" align="center" <%if rs("caiyin")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<td height="18" align="center" <%if rs("fuhe")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
<td height="18" align="center" <%if rs("chengping")="ok" then%>bgcolor="#FFFF00"<%end if%> >　</td>
			</tr>
<%end if%>


<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="11" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="dindang3.asp?keyword="&keyword&"&akeyword="&akeyword&"&search=search&"
			else
            url="dindang3.asp?"
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
     </td></tr>
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
         
		  <form name="search" method="post" action="dindang3.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">按品名查找： 
				<input name="keyword" type="text"  size="23"></td>
                      
			<td align="center" width="13%" ><input type="submit"  value="查 找"> </td>
			<td align="center" width="35%">按客户查找： 
			<input name="akeyword" type="text"  size="22"></td>
                      
			<td align="center" width="17%" ><input type="submit"  value="查 找"> </td>

          </tr></form>		   
		</td></tr>
	   </table>
		<table width="100%" border="0" cellspacing="0" cellpadding="4">
      
      </table>         
   </TBODY>  
</TABLE>

<P>
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <TBODY>
  <TR vAlign=center>
    <TD align=middle width="100%"><!--#include file="footer.htm"--></TD>
  </TR></TBODY></TABLE></P></BODY></HTML>
