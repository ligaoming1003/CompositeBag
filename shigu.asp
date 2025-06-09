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
  
 if (powersearch.sname.value == "")
	{
		alert("责任人不能为空！");
		powersearch.sname.focus();
		return (false);
	}
   if (powersearch.cname.value == "0")
	{
		alert("产品名称不能为空！");
		powersearch.cname.focus();
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
uptime=request.Form("uptime")  
pingming1=trim(request.Form("pingming1"))
pingming2=trim(request.Form("pingming2"))
pingming3=trim(request.Form("pingming3"))
use_dep=trim(request.Form("use_dep"))
chenben=request.Form("chenben")
loan_oan=request.Form("loan_oan")
sname=trim(request.Form("sname"))
cname=trim(request.Form("cname"))
miaoshu=trim(request.Form("miaoshu"))
    
if use_dep="彩印车间" then
         set rs3=server.createobject("adodb.recordset")         
          sql3="select * from caiyin where cname='"&cname&"'"
        rs3.open sql3,conn,2,3
         if not rs3.eof and not rs3.bof then
         rs3("qianmi")=rs3("qianmi")-Loan_oan
         else 
           response.Write "<script language=javascript>alert('无此产品！');</script>"
           response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
           response.end
           
         end if
         rs3.update
         rs3.close
         set rs3=nothing
         
elseif use_dep="复合车间" then
    set rs1=server.createobject("adodb.recordset") 
    sql1="select * from fuhe where cname='"&cname&"'"
    rs1.open sql1,conn,1,3
     if not rs1.eof and not rs1.bof then
         rs1("m_loan_num")=rs1("m_loan_num")-Loan_oan
     else 
     response.Write "<script language=javascript>alert('无此产品！');</script>"
     response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
     response.end           
     end if
    rs1.update
    rs1.close
    set rs1=nothing
    
    elseif use_dep="成品仓库" then
    set rs1=server.createobject("adodb.recordset") 
    sql1="select * from chengping_Store where pinming='"&cname&"'"
    rs1.open sql1,conn,1,3
     if not rs1.eof and not rs1.bof then
         rs1("number")=rs1("number")-Loan_oan
     else 
     response.Write "<script language=javascript>alert('无此产品！');</script>"
     response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
     response.end           
     end if
    rs1.update
    rs1.close
    set rs1=nothing

    
elseif use_dep="分切车间" or use_dep="制袋车间" then
    set rs2=server.createobject("adodb.recordset") 
    sql2="select * from matur where cname='"&cname&"'"
    rs2.open sql2,conn,1,3
     if not rs2.eof and not rs2.bof then
         rs2("number")=rs2("number")-Loan_oan
     else 
     response.Write "<script language=javascript>alert('无此产品！');</script>"
     response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
     response.end           
     end if
    rs2.update
    rs2.close
    set rs2=nothing

end if
set rs=server.createobject("adodb.recordset") 
sql="select * from shigu"
rs.open sql,conn,1,3
rs.addnew
rs("cname") = cname
rs("pingming1") = pingming1
rs("pingming2") = pingming2
rs("pingming3") = pingming3
rs("uptime") = request.Form("uptime")
rs("sname") =sname
rs("loan_oan") =loan_oan
rs("use_dep") =use_dep
rs("miaoshu") = miaoshu
rs("chenben") =chenben

rs.update
response.Write "<script language=javascript>alert('责任事故添加成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
response.end
rs.close
set rs=nothing
end sub

 
sub SaveModify
uptime=request.Form("uptime")  
pingming1=trim(request.Form("pingming1"))
pingming2=trim(request.Form("pingming2"))
pingming3=trim(request.Form("pingming3"))
use_dep=trim(request.Form("use_dep"))
chenben=request.Form("chenben")
loan_oan=request.Form("loan_oan")
sname=trim(request.Form("sname"))
cname=trim(request.Form("cname"))
miaoshu=trim(request.Form("miaoshu"))
old_cname=trim(request.Form("old_cname"))
old_num=request.Form("old_num")
num=Loan_oan-old_num
    
    if cname<>old_cname then
       response.Write "<script language=javascript>alert('品名不能修改！');</script>"
       response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
       response.end
    end if

if use_dep="彩印车间" then
         set rs3=server.createobject("adodb.recordset")

          sql3="select * from caiyin where cname='"&cname&"'"
        rs3.open sql3,conn,2,3
    if not rs3.eof and not rs3.bof then
         rs3("qianmi")=rs3("qianmi")-num

     end if
         rs3.update
         rs3.close
         set rs3=nothing
         
elseif use_dep="复合车间" then
    set rs1=server.createobject("adodb.recordset") 
    sql1="select * from fuhe where cname='"&cname&"'"
    rs1.open sql1,conn,1,3
     if not rs1.eof and not rs1.bof then
         rs1("m_loan_num")=rs1("m_loan_num")-num
          
     end if
    rs1.update
    rs1.close
    set rs1=nothing
    
elseif use_dep="成品仓库" then
    set rs1=server.createobject("adodb.recordset") 
    sql1="select * from chengping_Store where pinming='"&cname&"'"
    rs1.open sql1,conn,1,3
     if not rs1.eof and not rs1.bof then
         rs1("number")=rs1("number")-num         
     end if
    rs1.update
    rs1.close
    set rs1=nothing

    
elseif use_dep="分切车间" or use_dep="制袋车间" then
    set rs2=server.createobject("adodb.recordset") 
    sql2="select * from matur where cname='"&cname&"'"
    rs2.open sql2,conn,1,3
     if not rs2.eof and not rs2.bof then
         rs2("number")=rs2("number")-num
        
     end if
    rs2.update
    rs2.close
    set rs2=nothing

end if




set rs=server.createobject("adodb.recordset") 
sql="select * from shigu where id="&request.Form("id")
rs.open sql,conn,1,3
rs("pingming1") = pingming1
rs("pingming2") = pingming2
rs("pingming3") = pingming3
rs("uptime") = request.Form("uptime")
rs("sname") =sname
rs("loan_oan") =loan_oan
rs("use_dep") =use_dep
rs("miaoshu") = miaoshu
rs("chenben") =chenben
rs.update
rs.close
set rs=nothing

response.Write "<script language=javascript>alert('修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
response.end

end sub   
 
  sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
  response.end
else
        conn.execute("delete from shigu where id in ("&selBigClass&")")
		response.Write "<script language=javascript>alert('已删除成功！');</script>"
		response.write "<meta http-equiv=""refresh"" content=""0;url=shigu.asp"">"
		end if
response.end
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from shigu where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑事故信息</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加事故信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">日期</td>
            <td width="10%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			车&nbsp; 间</td>
			<td width="50%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">品 名</td>
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			面 膜</td>			
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">中膜</td>			
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			内 膜</td>
			
          </tr>
		  <form name="powersearch" method="post" action="shigu.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
		    <select name="use_dep" size=1 style='width:80; height:17'>
		    <% if isedit then%>
		    <option selected value='彩印车间'>彩印车间</option>
		    <option selected value='复合车间'>复合车间</option>
		    <option selected value='分切车间'>分切车间</option>
		    <option selected value='制袋车间'>制袋车间</option>
		    <option selected value='成品仓库'>成品仓库</option>		    
		    <option selected="selected"><%=rs("use_dep")%></option>
		    <input type="hidden" value="<%=rs("Loan_oan")%>" name="old_num">
		    <input type="hidden" value="<%=rs("cname")%>" name="old_cname">
		    <%else%>
		    <option selected value='彩印车间'>彩印车间</option>
		    <option selected value='复合车间'>复合车间</option>
		    <option selected value='分切车间'>分切车间</option>
		    <option selected value='制袋车间'>制袋车间</option>
		    <option selected value='成品仓库'>成品仓库</option>	
			<option selected value='0000'>请选择车间</option>
			<%end if%>
			</select>
						


</td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='cname'  size="1" style='width:396; height:27'>
			<% if isedit then%>
			<option ><%=rs("cname")%></option>
            <%else%>
			<option selected value='0'>--请选择--</option><%end if%>
 <%
               set rs1=server.createobject("adodb.recordset")
               rs1.open"select * from dindang order by -id",conn,1,2
               if not rs1.eof then
              do while not rs1.eof%>
              <option <% if isedit then
			if trim(rs1("pinming"))=trim(rs("cname")) then
			%> selected="selected"<%end if%><%end if%>><%=rs1("pinming")%></option>

                 		<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%>
			</select>
			</td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="pingming1" type="text"  size="9" value='<% if isedit then
		                                                         response.write rs("pingming1")
																  end if %>'></td>


			<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
&nbsp;<input name="pingming2" type="text"  size="9" value='<% if isedit then
		                                                         response.write rs("pingming2")
																  end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="pingming3" type="text"  size="9" value='<% if isedit then
		                                                         response.write rs("pingming3")
																  end if %>'></td>

          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			数量(千米)</td>
			<td height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			成本损失(元)</td>
			<td height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			事故描述</td>			
			<td  colspan="3" height="18" align="center" class="but" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but'">
			责任人</td>			
          </tr>
		  <tr> 
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="loan_oan" type="text"  size="8" value='<% if isedit then
		                                                         response.write rs("loan_oan")
																  end if %>'></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="chenben" type="text"  size="11" value='<% if isedit then
		                                                         response.write rs("chenben")
																  end if %>'></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<textarea name="miaoshu" cols="40" rows="2"><% if isedit then
		                                                         response.write rs("miaoshu")
																  end if %></textarea></td>

<td height="18" colspan="3" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="sname" type="text"  size="12" value='<% if isedit then
		                                                         response.write rs("sname")
																  end if %>'></td></tr>
			<tr>
			<td colspan="6" align="center"><input type="submit"  value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        <%end if%>

</td>
</tr>
</table><br><br>
<TABLE width=1020 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >

  <table width="1000"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="1000" align="center" valign="top">
	        <table width="1000" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp; 质 量 事 故 列 表</font></b>

</td>
</tr>
</table>
			<form action="#.asp" method="post" name="selform">
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">日期</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">车间</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">责任人</td>
			<td width="23%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">品 名</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">面膜</td>
			<td width="6%" height="18" align="center"  class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">中膜</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">内膜</td>
			<td width="6%" height="18" align="center"  class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">长度[km]</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">损失[元]</td>			
			<td width="26%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">事故描述</td>
            <td width="4%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">操作</td>
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
sname1=request("sname")
keyword=request("keyword")
rs.Source="select* from shigu where sname like '%"&sname1&"%' and use_dep like '%"&keyword&"%'order by -id"
else
rs.Source="select* from shigu order by -id"
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
<td height="18" align="center"><%=rs("use_dep")%><%if rs("use_dep")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("sname")%><%if rs("sname")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><a href="shigu.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("cname")%></a></td>
<td height="18" align="center"><%if rs("pingming1")="" then%>&nbsp;<%else%><%=rs("pingming1")%><%end if%></td>
<td height="18" align="center"><%if rs("pingming2")="" then%>&nbsp;<%else%><%=rs("pingming2")%><%end if%></td>
<td height="18" align="center"><%if rs("pingming3")="" then%>&nbsp;<%else%><%=rs("pingming3")%><%end if%></td>
<td height="18" align="center"><%if rs("loan_oan")="" then%>&nbsp;<%else%><%=Formatnumber(rs("loan_oan"),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%if rs("chenben")="" then%>&nbsp;<%else%><%=Formatnumber(rs("chenben"),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><%if rs("miaoshu")="" then%>&nbsp;<%else%><%=rs("miaoshu")%><%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
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
			url="shigu.asp?sname="&sname1&"&keyword="&keyword&"&search=search&"
			else
            url="shigu.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="admin" then%>
     <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input onClick="{if(confirm('此操作将删除该信息！\n\n为了数据的一致，请先把出库数量修改为0再删除\n\n确定要执行此项操作吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
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
<b><font color="#ffffff">责任搜索</font></b>

</td>
</tr>
</table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		           
<form name="search" method="post" action="shigu.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">按车间查找： 
				
				<select name='keyword'  style='width:130'>
			<option selected value="">请选择车间</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(use_dep) from shigu ",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("use_dep")%>"><%=rs2("use_dep")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
                      
			<td align="center" width="13%" ><input type="submit"  value="查 找"> </td>
			<td align="center" width="35%">按责任人查找： 
			<select name="sname"  style='width:130'>
			<option selected value="">请选择责任人</option>
			<%			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select distinct(sname) from shigu order by sname",conn,2,2
		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option value="<%=rs4("sname")%>"><%=rs4("sname")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
			%></select></td>
                      
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
