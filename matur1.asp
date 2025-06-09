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
		if (powersearch.Loan_oan.value == "")
	{
		alert("数量不能为空！");
		powersearch.Loan_oan.focus();
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
rs1.open "select * from chxp",conn,1,1
if not rs1.eof then   
do while not rs1.eof
j=0	
%>
po_ca_show[<%=i%>] = '<%=rs1("chxp")%>';
po_ca_value[<%=i%>] = '<%=rs1("chxp")%>';
po_detail_show[<%=i%>] = new Array();
po_detail_value[<%=i%>] = new Array();

<%
chxp=rs1("chxp")
set rs2=server.createobject("adodb.recordset")
rs2.open "select * from matur where chxp='"&chxp&"' order by number-daishu*2*chang*kuang/(guige-1.3)/100000-mishu/(zhesuang+lianpang+0.001) desc",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("cname")%>|<%=rs2("guige")%>|<%=Formatnumber( rs2("number")-rs2("daishu")*2*rs2("chang")*rs2("kuang")/(rs2("guige")-1.3)/100000-rs2("mishu")/(rs2("zhesuang")+rs2("lianpang")+0.001),2,-1,0,0)%>km';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("cname")%>|<%=rs2("guige")%>|<%=Formatnumber( rs2("number")-rs2("daishu")*2*rs2("chang")*rs2("kuang")/(rs2("guige")-1.3)/100000-rs2("mishu")/(rs2("zhesuang")+rs2("lianpang")+0.001),2,-1,0,0)%>km';
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
        num1= GetObjID('cname');

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

use_dep=trim(request.Form("use_dep")) 
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))
cname1=trim(request.Form("cname1"))
names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)

         set rs1=server.createobject("adodb.recordset") 
         sql1="select * from matur where  cname='"&cname&"'"
         rs1.open sql1,conn,2,3
      if not rs1.eof and not rs1.bof then
         rs1("number")=rs1("number")-Loan_oan
      end if
         rs1.update     
         rs1.close
         set rs1=nothing 

         set rs=server.createobject("adodb.recordset") 
         sql="select * from matur where  cname='"&cname1&"'"
         rs.open sql,conn,2,3
      if not rs.eof and not rs.bof then
         rs("number")=rs("number")+Loan_oan
      else
         rs.addnew
         rs("cname") = cname1
         rs("guige") = guige
         rs("number")=loan_oan
         rs("uptime")=trim(request.Form("uptime"))
      end if
         rs.update     
         rs.close
         set rs=nothing 
          
           set rs3=server.createobject("adodb.recordset") 
           sql3="select * from fuhe_in"
           rs3.open sql3,conn,1,3
           rs3.addnew
           rs3("loan_oan")=-Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") = request.Form("uptime") 
           rs3.update
           rs3.close
           set rs3=nothing 
           
            set rs2=server.createobject("adodb.recordset") 
           sql2="select * from fuhe_in"
           rs2.open sql2,conn,1,3
           rs2.addnew
           rs2("loan_oan")=Loan_oan
           rs2("guige") =guige
           rs2("cname") =cname1
           rs2("uptime") = request.Form("uptime") 
           rs2.update
                response.Write "<script language=javascript>alert('调整成功！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur1.asp"">"
                response.end
           rs2.close
           set rs2=nothing 


end sub

sub SaveModify 

end sub   
 
sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=matur1.asp"">"
  response.end
   else
        conn.execute("delete from * where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('已删除成功！');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=matur1.asp"">"
        response.end
        end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin" or oskey="chengpin"  then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from fuhe where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑成品调整信息</font></b>

</td>
</tr>
</table>
<%else%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加成品调整信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">出库日期</td>
            <td width="7%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">类 别</td>
			<td width="67%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'"><p align="left">&nbsp;&nbsp;&nbsp;&nbsp; 品 名 |宽度|尚存数|单位|&nbsp; 批号</td>
			<td width="15%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">数量</td>			
          </tr>
		  <form name="powersearch" method="post" action="matur1.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:58; height:11'>
			<% if isedit then%>
				<option selected="selected"><%=rs("chxp")%></option> 
			<%else%>
		       <option selected value='0000'>请选择类别</option>
			<%end if%>
			
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from chxp",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("chxp")%>"><%=rs1("chxp")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='cname' size=1 style='width:531; height:19'>
			<% if isedit then%>
			<option selected="selected" value="0|<%=rs("cname")%>|<%=rs("guige")%>|<%=Formatnumber((rs("m_loan_num")-rs("Loan_oan")),2,-1,0,0)%>|km|<%=rs("pinming")%>"><%=rs("cname")%>|<%=rs("guige")%>|<%=Formatnumber((rs("m_loan_num")-rs("Loan_oan")),2,-1,0,0)%>|km|<%=rs("pinming")%></option>
			<input type="hidden" value="<%=rs("Loan_oan")%>" name="old_num">
			<input type="hidden" value="<%=rs("use_dep")%>" name="old_dep">

			<%else%>
			<option>--请选择--</option>
			<%end if%>
			</select>
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="Loan_oan" type="text" style="ime-mode:disabled"  size="11" value='<% if isedit then
		                                                         response.write rs("Loan_oan")
																  end if %>'onfocus="if(this.value=='0')this.value=''" onblur="if(this.value=='')this.value='0'"><font color="#FFFFFF">千米</font></td>
          </tr>
		  <tr class="but">
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2">
			<p align="right">调整到：</td>
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2">
<p align="left"> 
			&nbsp;<select name='cname1' onchange='Do_po_Change(this);' valign=top style='width:533; height:46'>
		     <option selected value='0000'>请选择</option>
		
			
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from dindang order by -id",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("pinming")%>"><%=rs1("pinming")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td>
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
<b><font color="#ffffff">复 合 出 库 列 表</font></b>

</td>
</tr>
</table>
			<form action="matur1.asp" method="post" name="selform" >
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">入库日期</td>
			<td width="54%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">品 名</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">膜 宽</td>
			<td width="8%" height="18" align="center"  class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">半成品</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">成品</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">员工</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">操作</td>
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

rs.Source="select* from fuhe_in where cname like '%"&keyword&"%' and use_dep is null order by -id"
else
rs.Source="select* from fuhe_in where use_dep is null  order by -id"
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
<td height="18" align="center"><%=rs("cname")%></td>
<td height="18" align="center"><%=rs("guige")%></td>
<td height="18" align="center"><%if rs("one_fu")="0" then%>&nbsp;<%else%><%=rs("one_fu")%><%end if%></td>
<td height="18" align="center"><%if rs("loan_oan")="0" then%>&nbsp;<%else%><%=rs("loan_oan")%><%end if%></td>
<td height="18" align="center"><%if rs("pihao")="" then%>&nbsp;<%else%><%=rs("pihao")%><%end if%></td>
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
			url="fuhe_in.asp?keyword="&keyword&"&search=search&"
			else
            url="fuhe_in.asp?"
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

        	<div align="center">

        <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="matur1.asp">
		  <tr><input name="search" type="hidden" value="search">
             
			<td align="center" width="35%">按品名查找： 
			<input name="keyword" type="text"  size="22"></td>
                      
			<td align="center" width="17%" ><input type="submit"  value="查 找"> </td>

          </tr></form>		   
		</td></tr>
	   </table>
			</div>
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
