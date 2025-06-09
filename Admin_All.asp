<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>∷嘉美管理系统∷.</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript  src="js/selectcity.js"></script>
<script language="JavaScript" src="js/validate.js" type="text/JavaScript"></script>

<SCRIPT language=javascript src="css/init.js"></SCRIPT>
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
function Juge(myform)
{

	if (myform.Owner_name.value == "")
	{
		alert("购货单位名称不能为空！");
		myform.Owner_name.focus();
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
  case "del"
       call delCate()
  case else
       isEdit=False
       call myform(isEdit) 
  end select
 
  sub delCate()
        conn.execute("delete from ImgData where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=admin_all.asp"">"
response.end
  end sub

  %> 

<% sub myform(isEdit) %> 	  

<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >


			<form action="admin_all.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">公 司 员 工 信 息 列 表</font></b>

</td>
</tr>
</table>
	 
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">编号</td>
            <td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">员工姓名</td>
			<td width="12%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">所属部门</td>
			<td width="18%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">入厂时间</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">性别</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">学历</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">联系电话</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">详细资料</td>
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
typename1=request("typename")
keyword=request("keyword")
rs.Source="select* from ImgData where classname like '%"&typename1&"%' and  Photo_Name like '%"&keyword&"%' order by id desc"
else
rs.Source="select* from ImgData order by id"
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

<tr onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center" ><%=rs("bianhao")%></td>
<td height="18" align="center" ><a href="Admin_Edit.asp?id=<%=rs(0)%>"><%=Replace(Rs("Photo_Name"),Keyword,"<font color=red><b>"&Keyword&"</b></font>")%></a></td>
<td height="18" align="center" ><%=rs("classname")%></td>
<td height="18" align="center" ><%=rs("Photo_Time")%></td>
<td height="18" align="center" ><%=rs("Photo_author")%></td>
<td height="18" align="center" ><%=rs("Photo_Address")%></td>
<td height="18" align="center" ><%=rs("dianhua")%></td>
<td height="18" align="center" ><a href="view.Asp?id=<%=rs("id")%>" title="浏览<%=rs("photo_name")%>的详细资料及照片">查看</a></td>
<td height="18" align="center" ><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>


<%
rs.MoveNext
end if
next
end if

%>
<tr><td align="center" colspan="10" height="22" class="but1">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="admin_all.asp?classname="&typename1&"&keyword="&keyword&"&search=search&"
			else
                        url="admin_all.asp?"
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
&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input onClick="{if(confirm('此操作将删除该信息！\n\n确定要执行此项操作吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
              <input type="Hidden" name="action" value='del'></td></tr>
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
<b><font color="#ffffff">员 工 信 息 搜 索</font></b>

</td>
</tr>
</table>


        <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="admin_all.asp">
		  <tr><input name="search" type="hidden" value="search">
            <td height="40" valign="middle" align="center" width="20%">
			<select name="typename" size="1">
			<option value="">所有部门</option>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from department",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   
		   %>
		   <option value="<%=rs2("depname")%>"><%=rs2("depname")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>




            <td align="center" width="50%">员工姓名： <input name="keyword" type="text"  size="25"></td>
                      
			<td align="center" width="30%" ><input type="submit"  value="查 找"> </td>
          </tr></form>
		 
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


