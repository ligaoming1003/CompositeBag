<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˹���ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>
<style type="text/css">
<!--
td {  font-family: "����"; font-size: 9pt}
body {  font-family: "����"; font-size: 9pt}
select {  font-family: "����"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "����"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "����"; font-size: 9pt} 
-->
</style>
<SCRIPT LANGUAGE=javascript>
<!--
function Juge(powersearch)
{
		if (powersearch.cname.value == "0000")
	{
		alert("��ѡ�����ƣ�");
		powersearch.cname.focus();
		return (false);
	}

		if (powersearch.Loan_num.value == "")
	{
		alert("��������Ϊ�գ�");
		powersearch.Loan_num.focus();
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
rs2.open "select * from cailiao_store where fenlei='"&fenleiname&"'and number<>0 order by pinming",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=Formatnumber(rs2("number"),2,-1,0,0)%>|<%=rs2("unit")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=Formatnumber(rs2("number"),2,-1,0,0)%>|<%=rs2("unit")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '���޲��ϳ���';
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

fenlei= trim(request.Form("d_position1"))
Loan_num=request.Form("Loan_num")
pinming=trim(request.Form("pinming"))
cname=trim(request.Form("cname"))
chuku=trim(request.Form("chuku"))
shueming=trim(request.Form("shueming"))
checkbox=trim(request.Form("checkbox"))
names=Split(pinming,"|")
id=names(0)
pinming=names(1)
guige=names(2)
houdou=names(3)
num=names(4)
unit=names(5)
       set rs1=server.createobject("adodb.recordset") 
       sql1="select * from cailiao_store where id="&id&""
       rs1.open sql1,conn,1,3
          if rs1("number")-Loan_num<-0.001 then
          response.Write "<script language=javascript>alert('�Բ��𣬿�治�㣬��������ĳ�������');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=dindang1.asp"">"
          response.end
          end if
          rs1.update
          rs1.close
          set rs1=nothing

        set rs=server.createobject("adodb.recordset")

        sql="select * from dindang where pinming='"&cname&"'" 
        
        rs.open sql,conn,1,3
       if not rs.eof then
        if pinming=rs("name1") then
        rs("anpai1") =rs("anpai1")+Loan_num
          if shueming="" then
          rs("shueming1")=checkbox
          else
          rs("shueming1")=shueming
          end if
        elseif pinming=rs("name2") then
        rs("anpai2") =  rs("anpai2") +Loan_num
          if shueming="" then
          rs("shueming2")=checkbox
          else
          rs("shueming2")=shueming
          end if


        elseif pinming=rs("name3") then
        rs("anpai3") = rs("anpai3")+Loan_num
          if shueming="" then
          rs("shueming3")=checkbox
          else
          rs("shueming3")=shueming
          end if


        elseif pinming=rs("name4") then
        rs("anpai4") =rs("anpai4")+ Loan_num
          if shueming="" then
          rs("shueming4")=checkbox
          else
          rs("shueming4")=shueming
          end if


        else
        response.Write "<script language=javascript>alert('�ò�Ʒ�������ֲ��ϣ�');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=dindang1.asp"">"
         response.end
        end if
        end if
        rs.update       
        rs.close
        set rs=nothing  
        
 if chuku="��ӡ" and fenlei="��Ĥ" then
 set rs=server.createobject("adodb.recordset") 
        sql="select * from cailiao_out_store"
        rs.open sql,conn,1,3
        rs.addnew          
        rs("fenlei") = fenlei
        rs("guige") = guige
        rs("houdou") = houdou
        rs("pinming") = pinming
        rs("uptime") = date()
        rs("Unit") = unit
        rs("content") = trim(request.Form("content"))
        rs("use_dep") = "��ӡ����"
        rs("Loan_num") = Loan_num
        rs("cname") = trim(request.Form("cname"))

        rs.update
        rs.close
        set rs=nothing
        
                     set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")

        
          set rs2=server.createobject("adodb.recordset")
          sql2="select * from caiyin where cname='"&cname&"'"
          rs2.open sql2,conn,2,3
    if not rs2.eof and not rs2.bof then
       rs2("guige") = guige
       rs2("houdou")=houdou
       rs2("pinming") = pinming
       rs2("Loan_num")=rs2("Loan_num")+Loan_num
       rs2("qianmi")=Formatnumber((rs2("qianmi")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
       rs2("zhemi")=Formatnumber((rs2("zhemi")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
    else
       rs2.addnew
       rs2("fenlei") = fenlei
       rs2("guige") = guige
       rs2("houdou")=houdou
       rs2("pinming") = pinming
       rs2("uptime") = date()
       rs2("Unit") = unit
       rs2("Loan_num") = Loan_num
       rs2("qianmi")=Formatnumber((rs2("loan_num")/guige/houdou/bizhong),2,-1,0,0)
       rs2("zhemi")=Formatnumber((rs2("loan_num")/guige/houdou/bizhong),2,-1,0,0)
       rs2("cname") = cname

    end if
       rs2.update
       rs2.close
       set rs2=nothing
            rs.update
            rs.close
            set rs=nothing
            
 elseif chuku="����" and fenlei="��Ĥ" then
        set rs=server.createobject("adodb.recordset") 
        sql="select * from cailiao_out_store"
        rs.open sql,conn,1,3
        rs.addnew          
        rs("fenlei") = fenlei
        rs("guige") = guige
        rs("houdou") = houdou
        rs("pinming") = pinming
        rs("uptime") = date()
        rs("Unit") = unit
        rs("content") = trim(request.Form("content"))
        rs("use_dep") = "���ϳ���"
        rs("Loan_num") = Loan_num
        rs("cname") = trim(request.Form("cname"))

        rs.update
        rs.close
        set rs=nothing
        
             set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")

        
       set rs2=server.createobject("adodb.recordset")
       sql2="select * from fuhe where cname='"&cname&"'"
       rs2.open sql2,conn,2,3
    if not rs2.eof and not rs2.bof then
       rs2("m_Loan_num")=Formatnumber((rs2("m_Loan_num")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs2("guige") = guige
         rs2("houdou")= houdou 
         rs2("pinming") = pinming
    else
       rs2.addnew
       rs2("m_Loan_num")=Formatnumber((Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs2("fenlei") = fenlei
         rs2("guige") = guige
         rs2("houdou")= houdou
         rs2("pinming") = pinming
         rs2("uptime") = date()
         rs2("cname")= cname

    end if
       rs2.update
       rs2.close
       set rs2=nothing
            rs.update
            rs.close
            set rs=nothing
            
      set rs4=server.createobject("adodb.recordset") 
        sql4="select * from mxfuhe"
        rs4.open sql4,conn,1,3
        rs4.addnew
        rs4("mid")=id
        rs4("fenlei") = fenlei
        rs4("guige") = guige
        rs4("houdou") = houdou
        rs4("pinming") = pinming
        rs4("uptime") = date()
        rs4("Unit") = unit
        rs4("use_dep") = use_dep
        rs4("Loan_num") = Loan_num
        rs4("cname") = cname

        rs4.update
        rs4.close
        set rs4=nothing

elseif fenlei<>"��Ĥ" then 
        set rs=server.createobject("adodb.recordset") 
        sql="select * from cailiao_out_store"
        rs.open sql,conn,1,3
        rs.addnew          
        rs("fenlei") = fenlei
        rs("guige") = guige
        rs("houdou") = houdou
        rs("pinming") = pinming
        rs("uptime") = date()
        rs("Unit") = unit
        rs("content") = trim(request.Form("content"))
        rs("use_dep") = "���ϳ���"
        rs("Loan_num") = Loan_num
        rs("cname") = trim(request.Form("cname"))

        rs.update
        rs.close
        set rs=nothing
        
             set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")


       set rs3=server.createobject("adodb.recordset")
       sql3="select * from fuhe where cname='"&cname&"'"
        rs3.open sql3,conn,2,3
    if not rs3.eof and not rs3.bof then
       if fenlei="��Ĥ" then 
         rs3("zh_Loan_num")=Formatnumber((rs3("zh_Loan_num")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
         elseif fenlei="��Ĥ" then 
         rs3("l_Loan_num")=Formatnumber((rs3("l_Loan_num")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
         end if
         
      else
         rs3.addnew
        
         if fenlei="��Ĥ" then 
         rs3("zh_Loan_num")=Formatnumber((Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs3("uptime") = date()
         rs3("cname")= cname


                  
         elseif fenlei="��Ĥ" then 
         rs3("l_Loan_num")=Formatnumber((Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs3("uptime") = date()
         rs3("cname")= cname

         end if
      end if
         rs3.update
         rs3.close
         set rs3=nothing
            rs.update
            rs.close
            set rs=nothing
                     
        set rs4=server.createobject("adodb.recordset") 
        sql4="select * from mxfuhe"
        rs4.open sql4,conn,1,3
        rs4.addnew
        rs4("mid")=id
        rs4("fenlei") = fenlei
        rs4("guige") = guige
        rs4("houdou") = houdou
        rs4("pinming") = pinming
        rs4("uptime") = date()
        rs4("Unit") = unit
        rs4("use_dep") = use_dep
        rs4("Loan_num") = Loan_num
        rs4("cname") = cname

        rs4.update
        rs4.close
        set rs4=nothing

end if
    
        
        set rs1=server.createobject("adodb.recordset") 
        sql1="select * from cailiao_store where id="&id&""
        rs1.open sql1,conn,1,3
     if not rs1.eof then
        
            
            rs1("number")=rs1("number")-Loan_num

     end if
        rs1.update
         response.Write "<script language=javascript>alert('�����ɹ���');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=dindang1.asp"">"
        response.end
        rs1.close
        set rs1=nothing 
end sub 
sub SaveModify 



end sub   
 
sub delCate()
  selBigClass=Request.Form("selBigClass")
if selBigClass="" then 
  response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=dindang1.asp"">"
  response.end
else
        conn.execute("delete from dindang where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('��ɾ���ɹ���');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=dindang1.asp"">"
        response.end
end if
  end sub
  %><% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="cailiao" or oskey="admin" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from dindang where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�༭������Ϣ</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��ӵ�����Ϣ</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="15%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">�� ��</td>
			<td width="60%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp; Ʒ �� | ��&nbsp; �� | ����|&nbsp; ��λ</td>
			
			
			<td width="15%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">����</td>
			
			
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			��&nbsp; ��</td>
			
			
          </tr>
		  <form name="powersearch" method="post" action="dindang1.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:64; height:17'>
			<% if isedit then%>
				<option selected="selected"><%=rs("fenlei")%></option>
			<%else%>
		        <option selected value='0000'>��ѡ��</option>
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
			<select name='pinming' size=1 style='width:332; height:16'>
			<% if isedit then%>
			<option selected="selected" value="<%=rs("id")%>|<%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("Loan_num")%>|<%=rs("unit")%>"><%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("unit")%></option>
			<input type="hidden" value="<%=rs("Loan_num")%>" name="old_num">
			<input type="hidden" value="<%=rs("cname")%>" name="cname1">

			<%else%>
			<option>--��ѡ��--</option>
			<%end if%>
			</select>
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
 <input name="Loan_num" type="text" style="ime-mode:disabled"  size="16" value='<% if isedit then
		                                                         response.write rs("Loan_num")
																  end if %>'></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
	<select name='chuku'  size="1" style='width:60; height:20'>
			<% if isedit then%>
			
			<option selected value='����'>����</option>
            <option selected value='��ӡ'>��ӡ</option>           
                 			<%else%>
			<option selected value='����'>����</option>
            <option selected value='��ӡ'>��ӡ</option>
			<%end if%>			
			</select></td>

          </tr>
		  <tr class="but">
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" rowspan="2">
			ע�⣺��ҳ���޷��޸ģ�������ֻ���ø�����ת��</td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			��Ʒ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			˵��</td>			
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			����</td>			
          </tr>
		  <tr> 

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
	      
	      <%
               set rs1=server.createobject("adodb.recordset")
               exec="select * from dindang order by -id"
               rs1.open exec,conn,1,1
               %>
           <select name="cname" id="cname" style='width:429; height:17'>
			<% if isedit then%><option ><%=rs("cname")%></option><%else%>
		        <option selected value='0000'>��ѡ������</option>
			<%end if%>

			<%do while not rs1.eof%>
                 <option  value="<%=rs1("pinming")%>"><%=rs1("pinming")%></option>
                 <%
rs1.movenext
loop
rs1.close
set rs1=nothing
%>
                            </select>			</td>																		  
																  





			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			<input name="shueming" type="text"  size="16" value='<% if isedit then
		                                                         response.write rs("shueming")
																  end if %>'></td>			






			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
<input type="checkbox" name="checkbox" value='ok'>			</td>			



			</tr>
			<tr>
			<td colspan="4" align="center"><input type="submit" value="<% if isedit then
		                                                         response.write "�༭"
																 else
																 response.write "���"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        <%end if%>


</td>
</tr>
</table>
<br>
<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >
<form action="dindang2.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp; �� �� �� �� �� ��</font></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;<% if oskey="chejian" then%><input type="button" name="close" value="�ر�" onclick="window.close();" /><%end if%>

</td>
</tr>
</table>
			
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">����</td>
            <td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">�ͻ�</td>
			<td width="26%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">Ʒ ��</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">�� ��</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">����</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">���</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">��Ҫ��</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">	������</td>
            <td width="13%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">˵��</td>
            <td width="4%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">����</td>
          </tr>
		 <%
PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 200          'ÿҳ��ʾ������
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
rs.Source="select* from dindang where kehu like '%"&akeyword&"%' and pinming like '%"&keyword&"%'and zhuantai='ok' order by -id"
else
rs.Source="select* from dindang where zhuantai='ok' order by -id"
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
<td height="72" align="center" rowspan="4"><a href="zhidai.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="72" align="center" rowspan="4"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=rs("name1")%></td>
<td height="18" align="center"><%=rs("nianbang")%>*<%=rs("cunshu")%></td>
<td height="18" align="center"><%=rs("nianti")%><%=rs("unit1")%></td>
<td height="18" align="center"><%if rs("anpai1")="0" then%>&nbsp;<%else%><%=rs("anpai1")%><%end if%></td>
<td height="18" align="center"><%=rs("shueming1")%><%if rs("shueming1")="" then%>&nbsp;<%end if%></td>
<td height="72" align="center" rowspan="4"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name2")%></td>
<td height="18" align="center"><%=rs("nianbang2")%>*<%=rs("cunshu2")%></td>
<td height="18" align="center"><%=rs("nianti2")%><%=rs("unit2")%></td>
<td height="18" align="center"><%if rs("anpai2")="0" then%>&nbsp;<%else%><%=rs("anpai2")%><%end if%></td>
<td height="18" align="center"><%=rs("shueming2")%><%if rs("shueming2")="" then%>&nbsp;<%end if%> </td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name3")%></td>
<td height="18" align="center"><%=rs("nianbang3")%>*<%=rs("cunshu3")%></td>
<td height="18" align="center"><%=rs("nianti3")%><%=rs("unit3")%></td>
<td height="18" align="center"><%if rs("anpai3")="0" then%>&nbsp;<%else%><%=rs("anpai3")%><%end if%></td>
<td height="18" align="center"> <%=rs("shueming3")%><%if rs("shueming3")="" then%>&nbsp;<%end if%></td>
</tr>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name4")%></td>
<td height="18" align="center"><%=rs("nianbang4")%>*<%=rs("cunshu4")%></td>
<td height="18" align="center"><%=rs("nianti4")%><%=rs("unit4")%></td>
<td height="18" align="center"><%if rs("anpai4")="0" then%>&nbsp;<%else%><%=rs("anpai4")%><%end if%></td>
<td height="18" align="center"> <%=rs("shueming4")%><%if rs("shueming4")="" then%>&nbsp;<%end if%></td>
			</tr>
</tr>
<%elseif rs("name1")<>"0"  and rs("name2")<>"0" and rs("name3")="0"  and rs("name4")<>"0" then%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="54" align="center" rowspan="3"><%=rs("uptime")%></td>
<td height="54" align="center" rowspan="3"><%=rs("kehu")%><%if rs("kehu")="" then%>&nbsp;<%end if%></td>
<td height="54" align="center" rowspan="3"><a href="zhidai.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="54" align="center" rowspan="3"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=rs("name1")%></td>
<td height="18" align="center"><%=rs("nianbang")%>*<%=rs("cunshu")%></td>
<td height="18" align="center"><%=rs("nianti")%><%=rs("unit1")%></td>
<td height="18" align="center"><%if rs("anpai1")="0" then%>&nbsp;<%else%><%=rs("anpai1")%><%end if%></td>
<td height="18" align="center"><%=rs("shueming1")%><%if rs("shueming1")="" then%>&nbsp;<%end if%></td>
<td height="54" align="center" rowspan="3"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name2")%></td>
<td height="18" align="center"><%=rs("nianbang2")%>*<%=rs("cunshu2")%></td>
<td height="18" align="center"><%=rs("nianti2")%><%=rs("unit2")%></td>
<td height="18" align="center"><%if rs("anpai2")="0" then%>&nbsp;<%else%><%=rs("anpai2")%><%end if%></td>
<td height="18" align="center"><%=rs("shueming2")%><%if rs("shueming2")="" then%>&nbsp;<%end if%></td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name4")%></td>
<td height="18" align="center"><%=rs("nianbang4")%>*<%=rs("cunshu4")%></td>
<td height="18" align="center"><%=rs("nianti4")%><%=rs("unit4")%></td>
<td height="18" align="center"><%if rs("anpai4")="0" then%>&nbsp;<%else%><%=rs("anpai4")%><%end if%></td>
<td height="18" align="center"><%=rs("shueming4")%><%if rs("shueming4")="" then%>&nbsp;<%end if%></td>
			</tr>
</tr>
<%elseif rs("name1")<>"0"  and rs("name2")="0" and rs("name3")="0"  and rs("name4")<>"0" then%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="36" align="center" rowspan="2"><%=rs("uptime")%></td>
<td height="36" align="center" rowspan="2"><%=rs("kehu")%><%if rs("kehu")="" then%>&nbsp;<%end if%></td>
<td height="36" align="center" rowspan="2"><a href="zhidai.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="36" align="center" rowspan="2"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=rs("name1")%></td>
<td height="18" align="center"><%=rs("nianbang")%>*<%=rs("cunshu")%></td>
<td height="18" align="center"><%=rs("nianti")%><%=rs("unit1")%></td>
<td height="18" align="center"><%if rs("anpai1")="0" then%>&nbsp;<%else%><%=rs("anpai1")%><%end if%></td>
<td height="18" align="center"><%=rs("shueming1")%><%if rs("shueming1")="" then%>&nbsp;<%end if%></td>
<td height="36" align="center" rowspan="2"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("name4")%></td>
<td height="18" align="center"><%=rs("nianbang4")%>*<%=rs("cunshu4")%></td>
<td height="18" align="center"><%=rs("nianti4")%><%=rs("unit4")%></td>
<td height="18" align="center"><%if rs("anpai4")="0" then%>&nbsp;<%else%><%=rs("anpai4")%><%end if%></td>
<td height="18" align="center"><%=rs("shueming4")%><%if rs("shueming4")="" then%>&nbsp;<%end if%></td>
			</tr>
</tr>
<%elseif rs("name1")<>"0"  and rs("name2")="0" and rs("name3")="0"  and rs("name4")="0" then%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("uptime")%></td>
<td height="18" align="center"><%=rs("kehu")%><%if rs("kehu")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><a href="zhidai.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18" align="center"><%=rs("chang")%>*<%=rs("kuang")%></td>
<td height="18" align="center"><%=rs("name1")%></td>
<td height="18" align="center"><%=rs("nianbang")%>*<%=rs("cunshu")%></td>
<td height="18" align="center"><%=rs("nianti")%><%=rs("unit1")%></td>
<td height="18" align="center"><%if rs("anpai1")="0" then%>&nbsp;<%else%><%=rs("anpai1")%><%end if%></td>
<td height="18" align="center"><%=rs("shueming1")%><%if rs("shueming1")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td></tr>
<%end if%>


<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="10" height="22">
�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="dindang1.asp?keyword="&keyword&"&akeyword="&akeyword&"&search=search&"
			else
            url="dindang1.asp?"
			end if				
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
Response.write "<a class=black href='" & Url & "page=1' title='��1ҳ'>ҳ��</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
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
Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
End if
rs.Close
set rs=nothing
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="admin" then%>
     <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> ѡ��/��ѡ
              <input type=submit value=���� name=action> 
<%end if%></td></tr>
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
<b><font color="#ffffff">�������</font></b>

</td>
</tr>
</table>

        <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="dindang1.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">��Ʒ�����ң� 
				<input name="keyword" type="text"  size="23"></td>
                      
			<td align="center" width="13%" ><input type="submit"  value="�� ��"> </td>
			<td align="center" width="35%">���ͻ����ң� 
			<input name="akeyword" type="text"  size="22"></td>
                      
			<td align="center" width="17%" ><input type="submit"  value="�� ��"> </td>

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
