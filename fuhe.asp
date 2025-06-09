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
rs2.open "select * from fuhe where chxp='"&chxp&"' order by m_loan_num-Loan_oan desc",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("cname")%>|<%=rs2("guige")%>|<%=Formatnumber((rs2("m_loan_num")-rs2("Loan_oan")-rs2("m_oan_num")),2,-1,0,0)%>|km|<%=rs2("pinming")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("cname")%>|<%=rs2("guige")%>|<%=Formatnumber((rs2("m_loan_num")-rs2("Loan_oan")-rs2("m_oan_num")),2,-1,0,0)%>|km|<%=rs2("pinming")%>';
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
pinming=request.Form("pinming")
use_dep=trim(request.Form("use_dep")) 
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))
names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)


if guige="0" then
response.Write "<script language=javascript>alert('对不起,面料未进入,暂不能出库！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=fuhe.asp"">"
                response.end
end if
if use_dep<>"熟化室" then
                response.Write "<script language=javascript>alert('对不起,成型品只能进入熟化室！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=fuhe.asp"">"
                response.end
else
         set rs=server.createobject("adodb.recordset") 

         sql="select * from matur where  cname='"&cname&"'"

        rs.open sql,conn,2,3
      if not rs.eof and not rs.bof then
         rs("number")=rs("number")+Loan_oan
      else
         rs.addnew
         rs("cname") = cname
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
           rs3("loan_oan")=Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") = request.Form("uptime") 

           rs3("use_dep")=use_dep
           rs3.update
           rs3.close
           set rs3=nothing 
           

               set rs4=server.createobject("adodb.recordset") 
               sql4="select * from dindang where pinming='"&cname&"'"
                  rs4.open sql4,conn,1,2
                if not rs4.eof then
                   rs4("fuhe")="ok"
                end if
               rs4.update
               rs4.close
               set rs4=nothing  

   
           set rs1=server.createobject("adodb.recordset") 
           sql1="select * from fuhe where id="&id&""
            rs1.open sql1,conn,1,2
        if not rs1.eof then
           rs1("Loan_oan")=rs1("Loan_oan")+Loan_oan
           rs1("use_dep")=use_dep
        end if
        rs1.update
                response.Write "<script language=javascript>alert('出库成功！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=fuhe.asp"">"
                response.end
         rs1.close
         set rs1=nothing 
end if

end sub

sub SaveModify 

end sub   
 
sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=fuhe.asp"">"
  response.end
   else
        conn.execute("delete from fuhe where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('已删除成功！');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=fuhe.asp"">"
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
		   rs.open "select * from fuhe where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑复合车间出库信息</font></b>

</td>
</tr>
</table>
<%else%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加复合车间出库信息</font></b>

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
		  <form name="powersearch" method="post" action="fuhe.asp" onSubmit="return Juge(this)" >
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
																  end if %>'><font color="#FFFFFF">千米</font></td>
          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2">交付到</td>
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2" rowspan="2">
<p align="left">说明:1、复合车间如果退库,请在选择左边状态栏&quot;复合退库&quot;,不可在直接退库。<br>
&nbsp;&nbsp;&nbsp;&nbsp; 2、如果出库输入数有误，可以在左边状态栏&quot;出库明细&quot;中，点击&quot;名称&quot;的链接进行修改。</td>
          </tr>
		  <tr> 
 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="2">
	<select name="use_dep"  size="1" style='width:100; height:15'>
			<% if isedit then%><option ><%=rs("use_dep")%></option><%end if%>
			<%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from chejian order by id desc",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if trim(rs3("chejianname"))=trim(rs("use_dep")) then
			%> selected="selected"<%end if%><%end if%>><%=rs3("chejianname")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>



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
<TABLE width=1020 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6" height="164">
<tr>
<td >



			<form action="fuhe.asp" method="post" name="selform" >
  <table width="1000"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="1000" align="center" valign="top">
	        <table width="1000" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<font color="#FFFFFF"><b>&nbsp;复&nbsp; 合&nbsp; </b></font><b>
<font color="#ffffff">明&nbsp; 细&nbsp; 列&nbsp; 表</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="7%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">日期</td>
            <td width="24%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">名称</td>
			<td width="3%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">宽度</td>
			<td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" colspan="3">面膜(千米)</td>
			<td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" colspan="3">中膜(千米)</td>
			<td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" colspan="3">内膜(千米)</td>
			<td width="6%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">出库</td>
			<td width="5%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">损耗率</td>
            <td width="4%" height="36" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'" rowspan="2">操作</td>        </tr>
          <tr class="but"> 
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">领用</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">现存</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">退库</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">领用</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">现存</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">退库</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">领用</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">现存</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">退库</td>
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
rs.Source="select* from fuhe where cname like '%"&keyword&"%' order by -id"
else
rs.Source="select* from fuhe order by -id"
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
<td height="18"  align="center"><%=Formatnumber(rs("m_loan_num"),2,-1,0,0)%></td>
<td height="18" align="center" bgcolor="#FFFFCC"><%if rs("m_loan_num")="0" then%>&nbsp;<%else%><%=Formatnumber((rs("m_loan_num")-rs("Loan_oan")-rs("m_oan_num")),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%if rs("m_oan_num")="0"  then%>&nbsp;<%else%><%=Formatnumber(rs("m_oan_num"),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%=Formatnumber(rs("zh_loan_num"),2,-1,0,0)%></td>
<td height="18" align="center" bgcolor="#FFFFCC"><%if rs("zh_loan_num")="0" then%>&nbsp;<%else%><%=Formatnumber((rs("zh_loan_num")-rs("Loan_oan")-rs("zh_oan_num")),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%if rs("zh_oan_num")="0"  then%>&nbsp;<%else%><%=Formatnumber(rs("zh_oan_num"),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%=Formatnumber(rs("l_loan_num"),2,-1,0,0)%></td>
<td height="18" align="center" bgcolor="#FFFFCC"><%if rs("l_loan_num")="0" then%>&nbsp;<%else%><%=Formatnumber((rs("l_loan_num")-rs("Loan_oan")-rs("l_oan_num")),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%if rs("l_oan_num")="0"  then%>&nbsp;<%else%><%=Formatnumber(rs("l_oan_num"),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%if rs("Loan_oan")="0" then%><font color="#FF0000">在复</font><%else%><%=Formatnumber(rs("Loan_oan"),2,-1,0,0)%>km<%end if%></td>
<td height="18" align="center"><%if rs("m_Loan_num")="0" or rs("Loan_oan")="0" then%>&nbsp;<%else%><%=Formatnumber((1-rs("Loan_oan")/(rs("m_loan_num")-rs("m_oan_num")))*100,2,-1,0,0)%>%<%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="center" colspan="15" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="fuhe.asp?keyword="&keyword&"&search=search&"
			else
            url="fuhe.asp?"
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
         
		  <form name="search" method="post" action="fuhe.asp">
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
