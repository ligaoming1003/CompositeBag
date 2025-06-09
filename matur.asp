<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
    <title>∷管理系统∷</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <link href="css/main.css" rel="stylesheet" type="text/css">
    <script language="JavaScript" src="css/User_Info_Modify.js"></script>
    <style type="text/css">
        <!--
        td {
            font-family: "宋体";
            font-size: 9pt
        }

        body {
            font-family: "宋体";
            font-size: 9pt
        }

        select {
            font-family: "宋体";
            font-size: 9pt
        }

        A {
            text-decoration: none;
            color: #336699;
            font-family: "宋体";
            font-size: 9pt
        }

            A:hover {
                text-decoration: underline;
                color: #FF0000;
                font-family: "宋体";
                font-size: 9pt
            }
        -->
    </style>
    <script language="javascript">
<!--
    function Juge(powersearch) {
        if (powersearch.Loan_oan.value == "") {
            alert("数量不能为空！");
            powersearch.Loan_oan.focus();
            return (false);
        }
        if (powersearch.chang.value == "") {
            alert("长度不能为空！");
            powersearch.chang.focus();
            return (false);
        }
        if (powersearch.kuang.value == "") {
            alert("宽度不能为空！");
            powersearch.kuang.focus();
            return (false);
        }

    }


    function SelectAll() {
        for (var i = 0; i < document.selform.selBigClass.length; i++) {
            var e = document.selform.selBigClass[i];
            e.checked = !e.checked;
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
<!--rs2.open "select top 100 * from matur where chxp='"&chxp&"' order by number-daishu*2*chang*kuang/(guige-1.3)/100000-mishu/(zhesuang+lianpang+0.001) desc",conn,1,1-->

rs2.open "select top 500 * from matur where chxp='"&chxp&"' order by uptime desc",conn,1,1

if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("cname")%>|<%=rs2("guige")%>|<%if rs2("chang")="0" or rs2("kuang")="0" or rs2("number")="0" then%>未入库<%else%><%=Formatnumber( rs2("number")-rs2("daishu")*2*rs2("chang")*rs2("kuang")/(rs2("guige")-1.3)/100000-rs2("mishu")/(rs2("zhesuang")+rs2("lianpang")+0.01),2,-1,0,0)%><%end if%>km|<%if rs2("mishu")=0 then%><%=Formatnumber(rs2("number")*(rs2("guige")-1.3)/rs2("chang")/rs2("kuang")/2*100000,0)-rs2("daishu")%>条<%end if%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("cname")%>|<%=rs2("guige")%>|<%if rs2("chang")="0" or rs2("kuang")="0" or rs2("number")="0" then%>未入库<%else%><%=Formatnumber( rs2("number")-rs2("daishu")*2*rs2("chang")*rs2("kuang")/(rs2("guige")-1.3)/100000-rs2("mishu")/(rs2("zhesuang")+rs2("lianpang")+0.01),2,-1,0,0)%><%end if%>km|<%if rs2("mishu")=0 then%><%=Formatnumber(rs2("number")*(rs2("guige")-1.3)/rs2("chang")/rs2("kuang")/2*100000,0)-rs2("daishu")%>条<%end if%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '暂无半成品出库';
po_detail_value[<%=i%>][<%=j%>] = ''
<% end if
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
    <script language="JavaScript">



        var psid = "";

        function DoLoad(form, funtypev) {
            var n;
            var i, j, k;
            var num;
            num = GetObjID('funtype[]');
            num1 = GetObjID('cname');

            if (!funtypev)
                return;
            k = 0;
            for (i = 0; i < po_ca_show.length; i++) {
                for (j = 0; j < po_detail_value[i].length; j++) {
                    if (funtypev.indexOf(po_detail_value[i][j]) != -1) {
                        NewOptionName = new Option(po_detail_show[i][j], po_detail_value[i][j]);
                        form.elements[num].options[k] = NewOptionName;
                        k++;
                    }
                }
            }
        }

        function Do_po_Change(form) {
            var num, n, i, m;
            num = GetObjID('d_position1');
            m = document.powersearch.elements[num].selectedIndex - 1;
            n = document.powersearch.elements[num + 1].length;
            for (i = n - 1; i >= 0; i--)
                document.powersearch.elements[num + 1].options[i] = null;

            if (m >= 0) {
                for (i = 0; i < po_detail_show[m].length; i++) {
                    NewOptionName = new Option(po_detail_show[m][i], po_detail_value[m][i]);
                    document.powersearch.elements[num + 1].options[i] = NewOptionName;
                }
                document.powersearch.elements[num + 1].options[0].selected = true;
            }
        }

        function InsertItem(ObjID, Location) {
            len = document.powersearch.elements[ObjID].length;
            for (counter = len; counter > Location; counter--) {
                Value = document.powersearch.elements[ObjID].options[counter - 1].value;
                Text2Insert = document.powersearch.elements[ObjID].options[counter - 1].text;
                document.powersearch.elements[ObjID].options[counter] = new Option(trimPrefixIndent(Text2Insert), Value);
            }
        }
        function GetLocation(ObjID, Value) {
            total = document.powersearch.elements[ObjID].length;
            for (pp = 0; pp < total; pp++)
                if (document.powersearch.elements[ObjID].options[pp].text == "---" + Value + "---") {
                    return (pp);
                    break;
                }
            return (-1);
        }

        function GetObjID(ObjName) {
            for (var ObjID = 0; ObjID < window.powersearch.elements.length; ObjID++)
                if (window.powersearch.elements[ObjID].name == ObjName) {
                    return (ObjID);
                    break;
                }
            return (-1);
        }




    </script>
</head>
<body topmargin="0" rightmargin="0" leftmargin="0">
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
unit=trim(request.Form("unit")) 
use_dep=trim(request.Form("use_dep")) 
fenlei=trim(request.Form("fenlei")) 
chang=request.Form("chang")
zhesuang=request.Form("zhesuang")
lianpang=request.Form("lianpang")
kuang=request.Form("kuang")
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))
names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)


if use_dep<>"制袋车间" and use_dep<>"分切车间" and use_dep<>"彩印车间" then 
               response.Write "<script language=javascript>alert('成品只能从制袋、分切、彩印三车间入库！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
                response.end
elseif use_dep="制袋车间" and fenlei="卷材" then 
               response.Write "<script language=javascript>alert('制袋车间不能出卷材！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
                response.end
elseif use_dep="彩印车间" and fenlei="成品袋" then 
               response.Write "<script language=javascript>alert('彩印车间不能出卷材！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
                response.end

elseif use_dep="分切车间" and fenlei="成品袋" then 
               response.Write "<script language=javascript>alert('分切车间不能出成品袋！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
                response.end
elseif zhesuang=0 and unit="公斤" then
                response.Write "<script language=javascript>alert('搞错了！无千米重，应该成品袋，单位为只！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
                response.end
               
elseif zhesuang<>0 and unit<>"公斤" then
                response.Write "<script language=javascript>alert('搞错了！有千米重，应该是卷材，单位为公斤！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
                response.end
elseif unit="只" and fenlei="卷材" then 
               response.Write "<script language=javascript>alert('卷材的单位不能为只！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
                response.end
elseif unit="公斤" and fenlei="成品袋" then 
               response.Write "<script language=javascript>alert('成品袋的单位不能为公斤！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
                response.end
end if


           set rs=server.createobject("adodb.recordset") 
           sql="select * from matur where id="&id&""
            rs.open sql,conn,1,2
               if fenlei="成品袋" then
               rs("daishu")=rs("daishu")+Loan_oan 
               rs("Loan_oan")=Loan_oan       
               rs("chang")=chang
               rs("kuang")=kuang
               rs("fenlei")=fenlei
               rs("use_dep")=use_dep
               rs("uptime")=uptime
               rs("unit")=unit               
               elseif fenlei="卷材" then
               rs("mishu")=rs("mishu")+Loan_oan
               rs("Loan_oan")=Loan_oan 
               rs("zhesuang")=zhesuang
               rs("lianpang")=lianpang
               rs("chang")=chang
               rs("kuang")=kuang
               rs("fenlei")=fenlei
               rs("use_dep")=use_dep
               rs("uptime")=uptime
               rs("unit")=unit 
              end if
               rs.update
               rs.close
               set rs=nothing 
         
           set rs2=server.createobject("adodb.recordset") 
           sql2="select top 100 * from chengping_in_store order by uptime desc"
           rs2.open sql2,conn,1,3
           rs2.addnew
           rs2("fenlei") =fenlei
           rs2("unit")=unit
           rs2("use_num")=Loan_oan
           rs2("guige") =chang
           rs2("pinming") =cname
           rs2("uptime") =uptime
           rs2("kuang")=kuang

           rs2("use_dep")=use_dep
           rs2("Supplier")=trim(request.Form("Supplier"))
           rs2.update
           rs2.close
           set rs2=nothing           

         set rs1=server.createobject("adodb.recordset") 

         sql1="select * from chengping_store where pinming='"&cname&"' and fenlei='"&fenlei&"'and guige="&chang&" and kuang="&kuang&" and unit='"&unit&"'"

            rs1.open sql1,conn,1,3
         if not rs1.eof and not rs1.bof then

             rs1("number")=rs1("number")+Loan_oan

         else
            rs1.addnew
            rs1("fenlei") =fenlei            
            rs1("number")=Loan_oan
            rs1("unit")=unit
            rs1("guige") = chang
            rs1("pinming") = cname
            rs1("kuang") = kuang

            rs1("uptime")=uptime
          end if
           rs1.update
           rs1.close
           set rs1=nothing
           
          set rs4=server.createobject("adodb.recordset") 
           sql4="select * from dindang where pinming='"&cname&"'"
            rs4.open sql4,conn,1,2
        if not rs4.eof then
           rs4("chengping")="ok"
           rs4("zhuantai")=""
         end if
          rs4.update
          rs4.close
          set rs4=nothing 
          
                      
         response.Write "<script language=javascript>alert('入库成功！');</script>"
         response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
         response.end
                
         
end sub


sub SaveModify 
response.Write "<script language=javascript>alert('不可修改,如输入有误,\n\n\请输入负数充零后现重新输入！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
response.end

end sub   
    sub delCate()
   selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
  response.end
else
        conn.execute("delete from matur where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('已删除成功！');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=matur.asp"">"
        response.end
        end if
  end sub

    %>
    <% sub myform(isEdit) %>
    <table width="800" align="center" border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
        <tbody>
            <tr>
                <td align="center" valign="top">
                    <%if oskey="supper" or oskey="admin" or oskey="chengpin" then%>
                    <%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from matur where id=" & request("id"),conn,1,1
                    %>
                    <table width="800" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td height="26" background="images/background.gif">
                                <b><font color="#ffffff">编辑成品入库信息</font></b>
                            </td>
                        </tr>
                    </table>
                    <%else%>
                    <table width="800" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td height="26" background="images/background.gif">
                                <b><font color="#ffffff">添加成品入库信息</font></b>
                            </td>
                        </tr>
                    </table>
                    <%end if %>
                    <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
                        <tr class="but">
                            <td width="83" height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                onmouseup="this.className='but'" onmouseout="this.className='but'">入库日期
                            </td>
                            <td width="71" height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                onmouseup="this.className='but'" onmouseout="this.className='but'">类 别
                            </td>
                            <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                onmouseup="this.className='but'" onmouseout="this.className='but'" colspan="8">
                                <p align="left">
                                    &nbsp;&nbsp;&nbsp;&nbsp; 品 名 |宽度|&nbsp; 批号|尚存
                            </td>
                        </tr>
                        <form name="powersearch" method="post" action="matur.asp" onsubmit="return Juge(this)">
                            <tr>
                                <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'>
                                <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>">
                                <% End If %>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="uptime" onfocus="show_cele_date(uptime,'','',uptime)" type="text" size="9"
                                        value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'>
                                </td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="d_position1" onchange='Do_po_Change(this);' valign="top" style="width: 56; height: 13">
                                        <% if isedit then%>
                                        <option selected="selected">
                                            <%=rs("chxp")%></option>
                                        <%else%>
                                        <option selected value='0000'>请选择</option>
                                        <%end if%>
                                        <%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from chxp",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
                                        %>
                                        <option value="<%=rs1("chxp")%>">
                                            <%=rs1("chxp")%></option>
                                        <%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
                                        %>
                                    </select>
                                </td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="8">
                                    <select name='cname' size="1" style='width: 645; height: 30'>
                                        <% if isedit then%>
                                        <option selected="selected" value="0|<%=rs("cname")%>|<%=rs("guige")%> |<%=Formatnumber( rs("number")-rs("daishu")*2*rs("chang")*rs("kuang")/(rs("guige")-1.3)/100000-rs("mishu")/(rs("zhesuang")+rs("lianpang")+0.01),2,-1,0,0)%>|km">
                                            <%=Formatnumber( rs("number")-rs("daishu")*2*rs("chang")*rs("kuang")/(rs("guige")-1.3)/100000-rs("mishu")/(rs("zhesuang")+rs("lianpang")+0.01),2,-1,0,0)%>|km</option>
                                        <input type="hidden" value="<%=rs("Loan_oan")%>" name="old_num">
                                        <input type="hidden" value="<%=rs("zhesuang")%>" name="old_suang">
                                        <input type="hidden" value="<%=rs("use_dep")%>" name="old_dep">
                                        <input type="hidden" value="<%=rs("unit")%>" name="old_unit">
                                        <input type="hidden" value="<%=rs("fenlei")%>" name="old_fenlei">
                                        <input type="hidden" value="<%=rs("chang")%>" name="old_chang">
                                        <input type="hidden" value="<%=rs("kuang")%>" name="old_kuang">
                                        <%else%>
                                        <option>--请选择--</option>
                                        <%end if%>
                                    </select>
                                </td>
                            </tr>
                            <tr class="but">
                                <td height="36" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" colspan="2"
                                    rowspan="2">数量
                                </td>
                                <td height="36" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="90"
                                    rowspan="2">单位
                                </td>
                                <td height="36" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="136"
                                    rowspan="2">交付车间
                                </td>
                                <td height="36" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="139"
                                    rowspan="2">成品类型
                                </td>
                                <td height="36" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="134"
                                    rowspan="2">联版
                                </td>
                                <td height="36" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="137"
                                    rowspan="2">千米重
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="138"
                                    colspan="2">袋规格
                                </td>
                                <td height="14" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" rowspan="2"
                                    width="252">
                                    <p>
                                        包装规格
                                </td>
                            </tr>
                            <tr class="but">
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="83">长
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="53">宽
                                </td>
                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <input name="Loan_oan" type="text" style="ime-mode: disabled" size="10" value='<% if isedit then
		                                                         response.write rs("Loan_oan")
																  end if %>'>
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="90">
                                    <select name="unit" size="1" style='width: 47; height: 12'>
                                        <% if isedit then%><option>
                                            <%=rs("unit")%></option>
                                        <%end if%>
                                        <%
			
			set rs5=server.createobject("adodb.recordset")
		   rs5.open "select * from unit",conn,1,1
		   if not rs5.eof then
		   do while not rs5.eof
                                        %>
                                        <option <% if isedit then
			if trim(rs5("unit"))=trim(rs("unit")) then
			%>
                                            selected="selected" <%end if%><%end if%>>
                                            <%=rs5("unit")%></option>
                                        <%
			rs5.movenext
			loop
			end if
			rs5.close
			set rs5=nothing
                                        %>
                                    </select>
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="136">
                                    <select name="use_dep" size="1" style='width: 97; height: 22'>
                                        <% if isedit then%><option>
                                            <%=rs("use_dep")%></option>
                                        <%end if%>
                                        <%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from chejian order by id desc",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
                                        %>
                                        <option <% if isedit then
			if trim(rs3("chejianname"))=trim(rs("use_dep")) then
			%>
                                            selected="selected" <%end if%><%end if%>>
                                            <%=rs3("chejianname")%></option>
                                        <%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
                                        %>
                                    </select>
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="139">
                                    <select name="fenlei" size="1" style='width: 55; height: 14'>
                                        <% if isedit then%><option>
                                            <%=rs("fenlei")%></option>
                                        <%end if%>
                                        <%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from fenlei1 order by id desc",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
                                        %>
                                        <option <% if isedit then
			if trim(rs4("fenlei"))=trim(rs("fenlei")) then
			%>
                                            selected="selected" <%end if%><%end if%>>
                                            <%=rs4("fenlei")%></option>
                                        <%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
                                        %>
                                    </select>
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="134">
                                    <input name="lianpang" type="text" style="ime-mode: disabled" size="5" value='<% if isedit then
		                                                         response.write rs("lianpang")
		                                                         else
																 response.write 1

																  end if %>'>
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="137">
                                    <input name="zhesuang" type="text" style="ime-mode: disabled" size="6" value='<% if isedit then
		                                                         response.write rs("zhesuang")
		                                                          else
																 response.write 0

																  end if %>'>
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'">
                                    <input name="chang" type="text" style="ime-mode: disabled" size="4" value='<% if isedit then
		                                                         response.write rs("chang")
																  end if %>'>
                                </td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="53">
                                    <input name="kuang" type="text" style="ime-mode: disabled" size="4" value='<% if isedit then
		                                                         response.write rs("kuang")
																  end if %>'>
                                </td>
                                <td height="13" align="center" class="but" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but'" width="252">
                                    <input name='Supplier' type="text" size="18" value='<% if isedit then
		                                                         response.write rs("Supplier")
																  end if %>'
                                        maxlength="50">
                                </td>
                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'"
                                    onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="10">
                                    <p align="left">
                                        <font color="#FFFFFF">说明:<b>1、</b>类型为卷材，单位为公斤，输入千米重，联版保持默认值；<b>2、</b>类型为卷材，单位为千米，输入联版，千米重保持默认值
                                        ；<br>
                                            &nbsp;&nbsp;&nbsp;&nbsp; <b>3、</b>类型为成品袋，联版和千米重都保持默认值；<b>4、</b>输入有误，请在入库明细中修改。</font>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="10" align="center">
                                    <input type="submit" value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>">
                                </td>
                            </tr>
                        </form>
                    </table>
                    <%end if%>
                </td>
            </tr>
        </tbody>
    </table>
    <br>
    <table width="800" align="center" border="1" cellspacing="0" cellpadding="1" bordercolor="#0055E6"
        height="51">
        <tr class="but">
            <td width="100%" height="25" align="center" class="but01" onmousedown="this.className='tddown'"
                onmouseup="this.className='but'" onmouseout="this.className='but01'">
                <p align="left">
                    &nbsp;&nbsp; 客户名 『产品名』 规 格 【面膜】 批号&nbsp;&nbsp; 千米重
            </td>
        </tr>
        <tr class="but">
            <td width="100%" height="25" align="center" class="but01" onmousedown="this.className='tddown'"
                onmouseup="this.className='but'" onmouseout="this.className='but01'">
                <select size="1" style='width: 750; height: 21'>
                    <%
			
			set rs=server.createobject("adodb.recordset")
		   rs.open "select top 500 * from dindang order by uptime desc",conn,1,1
                    %><!--#include file="husuan.asp"--><%
		   if not rs.eof then%>
                    <option>请选择名称</option>
                    <%
		   do while not rs.eof
                    %>
                    <option>
                        <%=rs("kehu")%>
                        『<%=rs("pinming")%>』
                        <%if isnull(rs("bian")) or rs("bian")="" then%>
                        <%=rs("chang")%><%else%><%=rs("chang")+rs("bian")%><%end if%>*<%if isnull(rs("di")) or rs("di")="" then%><%=rs("kuang")%><%else%><%=rs("kuang")+rs("di")%><%end if%>
                        【<%=rs("nianbang")%>】&nbsp;&nbsp;&nbsp;&nbsp;<%=Formatnumber((rs("nianbang")*rs("cunshu")*0.11+rs("nianbang")*rs("cunshu2")*0.12+rs("nianbang")*rs("cunshu3")*0.12+rs("nianbang")*rs("cunshu4")*0.105),1,0)%>
                    </option>
                    <%
			rs.movenext
			loop
			end if
			rs.close
			set rs=nothing
                    %>
                </select>
            </td>
        </tr>
    </table>
    <br>
    <table width="1000" align="center" border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6"
        height="164">
        <tr>
            <td>
                <form action="matur.asp" method="post" name="selform">
                    <table width="1000" border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6"
                        align="center">
                        <tr>
                            <td width="1000" align="center" valign="top">
                                <table width="1000" border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td height="26" background="images/background.gif">
                                            <b><font color="#ffffff">产 品 库 存 列 表</font></b>
                                        </td>
                                    </tr>
                                </table>
                                <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
                                    <tr class="but">
                                        <td width="7%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">出库日期
                                        </td>
                                        <td width="5%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">类 别
                                        </td>
                                        <td width="30%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">品 名
                                        </td>
                                        <td width="5%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">规格
                                        </td>
                                        <td width="7%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">数 量
                                        </td>
                                        <td width="7%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">应做<br>
                                            袋数
                                        </td>
                                        <td width="12%" height="18" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" colspan="2">卷材（ 公斤）
                                        </td>
                                        <td width="12%" height="18" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" colspan="2">成品袋（只）
                                        </td>
                                        <td width="6%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">差额<br>
                                            (只)
                                        </td>
                                        <td width="5%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">损耗率
                                        </td>
                                        <td width="4%" height="36" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">操作
                                        </td>
                                    </tr>
                                    <tr class="but">
                                        <td width="6%" height="18" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'">千米重
                                        </td>
                                        <td width="6%" height="18" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'">实做
                                        </td>
                                        <td width="6%" height="18" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'">尚做
                                        </td>
                                        <td width="6%" height="18" align="center" class="but01" onmousedown="this.className='tddown'"
                                            onmouseup="this.className='but'" onmouseout="this.className='but01'">实做
                                        </td>
                                    </tr>
                                    <%

                                    %>
                                    <%
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 50          '每页显示多少条
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet") 
search=request("search")
if search<>"" then 

keyword=request("keyword")
rs.Source="select * from matur where cname like '%"&keyword&"%' order by id desc"
else
rs.Source="select top 500 * from matur order by uptime desc "
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
                                    <!--<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">-->
                                    <tr>
                                        <td height="18" align="center">
                                            <%=rs("uptime")%>
                                        </td>
                                        <td height="18" align="center">
                                            <%=rs("chxp")%>
                                        </td>
                                        <td height="18" align="center">
                                            <%=rs("cname")%>
                                        </td>
                                        <td height="18" align="center">
                                            <%=rs("chang")%>*<%=rs("kuang")%>
                                        </td>
                                        <td height="18" align="center">
                                            <%=Formatnumber(rs("number"),2,-1,0,0)%>
                                        </td>
                                        <td height="18" align="center">
                                            <%if rs("chang")="0" or rs("kuang")="0" then%>&nbsp;<%else%><%=Formatnumber(rs("number")*(rs("guige")-1.3)/rs("chang")/rs("kuang")/2*100000,0)%><%end if%>
                                        </td>
                                        <td height="18" align="center">
                                            <%if rs("zhesuang")="0" then%>&nbsp;<%else%><%=rs("zhesuang")%><%end if%>
                                        </td>
                                        <td height="18" align="center">
                                            <%if rs("mishu")="0" then%>&nbsp;<%else%><%=rs("mishu")%><%end if%>
                                        </td>
                                        <td height="18" align="center">
                                            <%if rs("chang")="0" or rs("kuang")="0" then%>&nbsp;
                                        <%elseif rs("mishu")="0" then%><%=Formatnumber(rs("number")*(rs("guige")-1.3)/rs("chang")/rs("kuang")/2*100000,0)%>
                                            <%elseif rs("zhesuang")<>"0" then%><%=Formatnumber((rs("number")-rs("mishu")/rs("zhesuang"))*(rs("guige")-1.3)/rs("chang")/rs("kuang")/2*100000,0)%>
                                            <%end if%>
                                        </td>
                                        <td height="18" align="center">
                                            <%if rs("daishu")="0" then%>&nbsp;<%else%><%=rs("daishu")%><%end if%>
                                        </td>
                                        <td height="18" align="center" bgcolor="#EFE6DE">
                                            <%if rs("chang")="0" or rs("kuang")="0" or rs("number")="0" then%>&nbsp;
                                        <%elseif rs("mishu")=rs("zhesuang")*rs("number")="0" then %>0
                                        <%elseif rs("mishu")="0" then%><%=Formatnumber(rs("number")*(rs("guige")-1.3)/rs("chang")/rs("kuang")/2*100000,0)-rs("daishu")%>
                                            <%elseif rs("zhesuang")<>"0" then%><%=Formatnumber((rs("number")-rs("mishu")/rs("zhesuang"))*(rs("guige")-1.3)/rs("chang")/rs("kuang")/2*100000,0)-rs("daishu")%>
                                            <%elseif rs("zhesuang")="0" and rs("unit")="千米" then %><%=Formatnumber((rs("number")*rs("lianpang")-rs("mishu")),2,-1,0,0)%>km
                                        <%end if%>
                                        </td>
                                        <td height="18" align="center" bgcolor="#EFE6DE">
                                            <%if rs("chang")="0" or rs("kuang")="0" or rs("number")="0" then%>&nbsp;<%elseif  rs("mishu")="0" then%><%=Formatnumber((1-(rs("daishu")/(rs("number")*(rs("guige")-1.3)/rs("chang")/rs("kuang")/2*100000)))*100,2,-1,0,0)%>%<%elseif rs("zhesuang")<>"0" then%><%=Formatnumber(((1-(rs("mishu")/rs("zhesuang")*(rs("guige")-1.3)/rs("chang")/rs("kuang")/2*100000+rs("daishu"))/(rs("number")*(rs("guige")-1.3)/rs("chang")/rs("kuang")/2*100000))*100),2,-1,0,0)%>%<%elseif rs("zhesuang")="0" and rs("unit")="千米" then %><%=Formatnumber(((rs("number")*rs("lianpang")-rs("mishu"))/(rs("number")*rs("lianpang")))*100,2,-1,0,0)%>%<%end if%>
                                        </td>
                                        <td height="18" align="center">
                                            <input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>">
                                        </td>
                                    </tr>
                                    <%
rs.MoveNext
end if
next
end if
                                    %>
                                    <tr>
                                        <td align="center" colspan="13" height="22">共
                                        <%=total%>
                                        条，当前第
                                        <%=Mypage%>/<%=Maxpages%>
                                        页，每页
                                        <%=MyPageSize%>
                                        条
                                        <%
			if search<>"" then
			url="matur.asp?keyword="&keyword&"&search=search&"
			else
            url="matur.asp?"
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
                                            <input type="checkbox" name="checkbox" value="checkbox" onclick="javascript:SelectAll()">
                                            选择/反选
                                        <input onclick="{if(confirm('确定要执行删除吗？')){this.document.selform.submit();return true;}return false;}"
                                            type="submit" value="删除" name="action2">
                                            <input type="Hidden" name="action" value='del'><%end if%>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </form>
            </td>
        </tr>
    </table>
    <br>
    <br>
    <table width="820" align="center" border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6">
        <tr>
            <td>
                <table width="800" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td height="26" background="images/background.gif">
                            <b><font color="#ffffff">出库搜索</font></b>
                        </td>
                    </tr>
                </table>
                <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
                    <form name="search" method="post" action="matur.asp">
                        <tr>
                            <input name="search" type="hidden" value="search">
                            <td height="40" valign="middle" align="center" width="20%"></td>
                            <td align="center" width="65%">品名关键字：
                            <input name="keyword" type="text" size="25">
                                关键字为空则搜索所有
                            </td>
                            <td align="center" width="15%">
                                <input type="submit" value="查 找">
                            </td>
                        </tr>
                    </form>
                </table>
            </td>
        </tr>
    </table>
    <%end sub%>
    <p>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tbody>
                <tr valign="center">
                    <td align="middle" width="100%">
                        <!--#include file="footer.htm"-->
                    </td>
                </tr>
            </tbody>
        </table>
    </p>
</body>
</html>
