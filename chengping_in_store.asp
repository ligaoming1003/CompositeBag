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
rs2.open "select top 3600 * from chengping_in_store order by id desc",conn,1,1  

'rs2.open "select * from chengping_in_store where uptime >=DATEADD(year,-2,cast(CAST(year(getdate()) as  varchar) + '-01-01' as datetime)) and uptime <=GETDATE() order by uptime desc",conn,1,1 

if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("kuang")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("kuang")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '暂无半成品出库';
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
 response.Write "<script language=javascript>alert('此页面只供修改！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_in_store.asp"">"
  response.end

end sub

 
sub SaveModify
old_fenlei=request.Form("old_fenlei")
old_num=request.Form("old_num")
fenlei=request.Form("fenlei")
Supplier=trim(request.Form("Supplier"))
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))

names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)
kuang=names(3)

num=old_num-Loan_oan

if old_fenlei<>fenlei then
response.Write "<script language=javascript>alert('只能修改数量,\n\n\其它错误，请先将数量改为0！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_in_store.asp"">"
  response.end
end if
 set rs2=server.createobject("adodb.recordset") 
           sql2="select * from chengping_store where fenlei='"&fenlei&"' and pinming='"&cname&"'and guige="&guige&" and kuang="&kuang&""
           rs2.open sql2,conn,1,3
          if not rs2.eof then
             if rs2("number")+0.0001<num then
                response.Write "<script language=javascript>alert('对不起！数量已出库,请先修改销售数量！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_in_store.asp"">"
                response.end   
             else
                rs2("number")=rs2("number")-num
              end if         
          end if
           rs2.update
           rs2.close
           set rs2=nothing 
           
          set rs1=server.createobject("adodb.recordset") 
          sql1="select * from chengping_in_store where id="&request.Form("id")
          rs1.open sql1,conn,1,3
          if not rs1.eof then
           rs1("use_num")=Loan_oan
           rs1("Supplier")=Supplier
          end if  
           rs1.update
           rs1.close
           set rs1=nothing
           
           if Loan_oan="0" then
          set rs4=server.createobject("adodb.recordset") 
           sql4="select * from dindang where pinming='"&cname&"'"
            rs4.open sql4,conn,1,2
        if not rs4.eof then
           rs4("chengping")=""
           rs4("zhuantai")="ok"
         end if
          rs4.update
          rs4.close
          set rs4=nothing
          end if
      
            
          set rs3=server.createobject("adodb.recordset") 
           sql3="select * from matur where  cname='"&cname&"'"
           rs3.open sql3,conn,1,3

           if fenlei="成品袋" then
             rs3("loan_oan")=rs3("loan_oan")-num
             rs3("daishu")=rs3("daishu")-num
           else
             rs3("loan_oan")=rs3("loan_oan")-num
             rs3("mishu")=rs3("mishu")-num
           end if

           rs3.update
             response.Write "<script language=javascript>alert('修改成功！');</script>"
             response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_in_store.asp"">"
             response.end 
           rs3.close
           set rs3=nothing 



end sub
   

sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_in_store.asp"">"
  response.end
else
        conn.execute("delete from chengping_in_store where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('删除成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=chengping_in_store.asp"">"
response.end
end if
  end sub

  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin"  or oskey="chengpin" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from chengping_in_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">修改成品入库信息</font></b>

</td>
</tr>
</table>
<%else%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">修改成品入库信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			类型</td>
			<td width="60%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'"><p align="left">&nbsp;&nbsp;&nbsp;&nbsp; 品 名 | 规&nbsp; 格 | 余数|&nbsp; 单位|&nbsp; 批号</td>
			
			<td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			数&nbsp;&nbsp; 量</td>
			
          </tr>
		  <form name="powersearch" method="post" action="chengping_in_store.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="fenlei"  type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("fenlei")
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name="cname" size=1 style='width:481; height:15'>
			<% if isedit then%>
			<option selected="selected" value="0|<%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("kuang")%>"><%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("kuang")%></option>
			<input type="hidden" value="<%=rs("use_num")%>" name="old_num">
             <input type="hidden" value="<%=rs("fenlei")%>" name="old_fenlei">

			<%else%>
			<option>--请选择--</option>
			<%end if%>
			</select>
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
 <input name="Loan_oan" type="text" style="ime-mode:disabled"  size="11" value='<% if isedit then
		                                                         response.write rs("use_num")
																  end if %>'><font color="#FFFFFF">千米(公斤)</font></td>
          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			<p>包装规格</td>
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2" rowspan="2">
			<p align="left">说明:出库数量输入有误，点击“品名”链接进行修改。
			如果其它项目输入有误，请先上数量修改为“0”，再重新输入。</td>
          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
<input name='Supplier' type="text"  size="18" value='<% if isedit then
		                                                         response.write rs("Supplier")
																  end if %>'>			</td>
          </tr>
			<tr>
			<td colspan="3" align="center"><input type="submit" value="<% if isedit then
		                                                         response.write "修改"
																 else
																 response.write "修改"
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
<b><font color="#ffffff">入 库 成 品 列 表</font></b>

</td>
</tr>
</table>
			<form action="chengping_in_store.asp" method="post" name="selform" >
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">入库日期</td>
            <td width="10%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">类 别</td>
			<td width="52%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">品 名</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">规 格</td>
			<td width="10%" height="18" align="center"  class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">数量</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">单 位</td>
            <td width="5%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">操 作</td>
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
rs.Source="select top 3600 * from chengping_In_Store where fenlei like '%"&fenlei1&"%' and pinming like '%"&keyword&"%'and use_num<>0 order by id desc"
else
rs.Source="select top 3600 * from chengping_In_Store where use_num<>0 and pihao=''or pihao is null order by id desc"
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
<td height="18" align="center"><a href="chengping_In_Store.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%><%if rs("supplier")<>"" then%><font color="#FF0000">[<%=rs("supplier")%>]</font><%end if%></a></td>
<td height="18" align="center">&nbsp;<%=rs("guige")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=Formatnumber(rs("use_num"),2,-1,0,0)%><%if rs("use_num")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="7" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="chengping_In_Store.asp?fenlei="&fenlei1&"&keyword="&keyword&"&search=search&"
			else
            url="chengping_In_Store.asp?"
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
         
		  <form name="search" method="post" action="chengping_in_store.asp">
		  <tr><input name="search" type="hidden" value="search">
            <td align="center" width="20%">
			<select name="fenlei" style='width:130'>
			<option value="">所有分类</option>
			<%
	       set rs=server.createobject("adodb.recordset")
		   rs.open "select * from fenlei1",conn,1,1
		   if not rs.eof then
		   do while not rs.eof
		   fenlei=rs("fenlei")
		   %>
		   <option value="<%=rs("fenlei")%>"><%=rs("fenlei")%></option>
			<%
			rs.movenext
			loop
			end if
			rs.close
			set rs=nothing
			%></select></td>
			<td align="center" width="10%" ><input type="submit"  value="查 找"> </td>
            <td align="center" width="60%">品名：
            <input name="keyword" type="text"  size="25"> </td>
                      
			<td align="center" width="10%" ><input type="submit"  value="查 找"> </td>
          </tr></form>		   
		</td></tr>
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