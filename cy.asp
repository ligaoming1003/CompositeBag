<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
    <title>�˼�������ϵͳ��</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <link href="css/main.css" rel="stylesheet" type="text/css">
    <script language="JavaScript" src="css/User_Info_Modify.js"></script>
    <style type="text/css">
        <!--
        td {
            font-family: "����";
            font-size: 9pt
        }

        body {
            font-family: "����";
            font-size: 9pt
        }

        select {
            font-family: "����";
            font-size: 9pt
        }

        A {
            text-decoration: none;
            color: #336699;
            font-family: "����";
            font-size: 9pt
        }

            A:hover {
                text-decoration: underline;
                color: #FF0000;
                font-family: "����";
                font-size: 9pt
            }
        -->
    </style>
    <script language="javascript">
<!--
    function Juge(powersearch) {
        if (powersearch.Loan_oan.value == "") {
            alert("��������Ϊ�գ�");
            powersearch.Loan_oan.focus();
            return (false);
        }
        if (powersearch.bang.value == "") {
            alert("�ϰ治��Ϊ�գ�");
            powersearch.bang.focus();
            return (false);
        }
        if (powersearch.cname.value == "") {
            alert("��ѡ���Ʒ���ƣ�");
            powersearch.cname.focus();
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
rs2.open "select * from caiyin where fenlei='"&fenleiname&"' order by qianmi desc",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("cname")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=Formatnumber(rs2("qianmi"),2,-1,0,0)%>|km|<%=rs2("pinming")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("cname")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=Formatnumber(rs2("qianmi"),2,-1,0,0)%>|km|<%=rs2("pinming")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '���ް��Ʒ����';
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


fenlei= trim(request.Form("d_position1"))
use_dep=trim(request.Form("use_dep")) 
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))
unit=trim(request.Form("unit"))
names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)
houdou=names(3)
qianmi=names(4)
pinming=names(5)


if use_dep="���ϳ���" then
          set rs=server.createobject("adodb.recordset")
          sql="select * from fuhe where cname='"&cname&"'"
          rs.open sql,conn,2,3
    if not rs.eof then
         rs("m_Loan_num")=rs("m_Loan_num")+Loan_oan
         rs("guige")=guige
         rs("fenlei")=fenlei
         rs("houdou")= houdou
         rs("pinming") = pinming
    else 
        rs.addnew

         rs("m_Loan_num")=Loan_oan
         rs("fenlei") = fenlei
         rs("guige") = guige
         rs("houdou")= houdou
         rs("pinming") = pinming
         rs("uptime") = request.Form("uptime")         
         rs("cname")= cname

    
     end if
         rs.update
         rs.close
         set rs=nothing
       

       set rs3=server.createobject("adodb.recordset") 
           sql3="select * from caiyin_in"
           rs3.open sql3,conn,1,3
           rs3.addnew
           rs3("unit")=unit
           rs3("loan_oan")=Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") = request.Form("uptime") 
           rs3("houdou")=houdou

           rs3("use_dep")=use_dep
           rs3("pihao")=request.Form("pihao")
           rs3("bang")=request.Form("bang")
           rs3.update
           rs3.close
           set rs3=nothing     
        
          set rs4=server.createobject("adodb.recordset") 
           sql4="select * from dindang where pinming='"&cname&"'"
            rs4.open sql4,conn,1,2
        if not rs4.eof then
           rs4("caiyin")="ok"
         end if
          rs4.update
          rs4.close
          set rs4=nothing
        
           set rs1=server.createobject("adodb.recordset") 
           sql1="select * from caiyin where id="&id&""
            rs1.open sql1,conn,1,2

        if not rs1.eof then
           rs1("Loan_oan")=rs1("Loan_oan")+Loan_oan
           rs1("qianmi")=rs1("qianmi")-Loan_oan 
           rs1("use_dep") = use_dep
        end if
        rs1.update
                response.Write "<script language=javascript>alert('����ɹ���');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=cy.asp"">"
                response.end
         rs1.close
         set rs1=nothing
         
else
             set rs5=server.createobject("adodb.recordset") 
             sql5="select * from matur where cname='"&cname&"'"
        rs5.open sql5,conn,2,3
      if not rs5.eof then
         rs5("number")=rs5("number")+Loan_oan
      else
         rs5.addnew
         rs5("cname") = cname
         rs5("guige") = guige
         rs5("number")=loan_oan
         rs5("pihao") =pihao
         rs5("uptime")=trim(request.Form("uptime"))
      end if
         rs5.update     
         rs5.close
         set rs5=nothing 
         
          set rs3=server.createobject("adodb.recordset") 
           sql3="select * from caiyin_in"
           rs3.open sql3,conn,1,3
           rs3.addnew
           rs3("unit")=unit
           rs3("loan_oan")=Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") = request.Form("uptime") 
           rs3("houdou")=houdou
           rs3("use_dep")=use_dep
           rs3("pihao")=request.Form("pihao")
           rs3("bang")=request.Form("bang")
           rs3.update
           rs3.close
           set rs3=nothing
                
          set rs4=server.createobject("adodb.recordset") 
           sql4="select * from dindang where pinming='"&cname&"'"
            rs4.open sql4,conn,1,2
        if not rs4.eof then
           rs4("caiyin")="ok"
         end if
          rs4.update
          rs4.close
          set rs4=nothing
         
          set rs1=server.createobject("adodb.recordset") 
           sql1="select * from caiyin where id="&id&""
            rs1.open sql1,conn,1,2

        if not rs1.eof then
           rs1("Loan_oan")=rs1("Loan_oan")+Loan_oan
           rs1("qianmi")=rs1("qianmi")-Loan_oan 
           rs1("use_dep") = use_dep
        end if
        rs1.update
                response.Write "<script language=javascript>alert('����ɹ���');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=cy.asp"">"
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
  response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=cy.asp"">"
  response.end
else
        conn.execute("delete from caiyin_in where id in ("&Request.Form("selBigClass")&") and loan_oan=0 and bang=0")
		response.Write "<script language=javascript>alert('����Ϊ0�ģ�ɾ���ɹ���\n\n������Ϊ0��û�б�ɾ����');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=cy.asp"">"
        response.end
end if
  end sub
    %> <% sub myform(isEdit) %>
    <table width="800" align="center" border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
        <tbody>
            <tr>
                <td align="center" valign="top"><%if oskey="supper" or oskey="admin" or oskey="chejian" then%>

                    <%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from caiyin_in where id=" & request("id"),conn,1,1
		   
                    %>

                    <table width="800" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td height="26" background="images/background.gif">
                                <b><font color="#ffffff">�༭��ӡ���������Ϣ</font></b>

                            </td>
                        </tr>
                    </table>
                    <%else%>
                    <table width="800" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td height="26" background="images/background.gif">
                                <b><font color="#ffffff">��Ӳ�ӡ���������Ϣ</font></b>

                            </td>
                        </tr>
                    </table>
                    <%end if %>
                    <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
                        <tr class="but">
                            <td width="40" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">����</td>
                            <td width="50" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">�� ��</td>
                            <td width="550" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">
                                <p align="left">
                                    &nbsp;&nbsp;&nbsp;&nbsp; 
			Ʒ �� | ��&nbsp; �� | ����|&nbsp; ��λ</td>
                            <td width="80" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">���ó���</td>

                            <td width="80" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">��λ</td>

                        </tr>
                        <form name="powersearch" method="post" action="cy.asp" onsubmit="return Juge(this)">
                            <tr>
                                <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'>
                                <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>">
                                <% End If %>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="uptime" onfocus="show_cele_date(uptime,'','',uptime)" type="text" size="10" value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name='d_position1' onchange='Do_po_Change(this);' valign="top" style='width: 61; height: 25'>
                                        <% if isedit then%>
                                        <option selected="selected"><%=rs("fenlei")%></option>
                                        <%else%>
                                        <option selected value='0000'>��ѡ�����</option>
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
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name='cname' size="1" style='width: 515; height: 17'>
                                        <% if isedit then%>
                                        <option selected="selected" value="0|<%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=Formatnumber(rs("qianmi"),2,-1,0,0)%>|km|<%=rs("pinming")%>"><%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=Formatnumber(rs("qianmi"),2,-1,0,0)%>|km|<%=rs("pinming")%></option>

                                        <%end if%>
                                    </select>
                                </td>

                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="use_dep" size="1" style='width: 68; height: 16'>
                                        <option selected value='�컯��'>�컯��</option>
                                        <option selected value='���ϳ���'>���ϳ���</option>

                                    </select></td>

                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="unit" size="1" style='width: 50; height: 15'>
                                        <option selected value='ǧ��'>ǧ��</option>
                                    </select></td>
                            </tr>
                            <tr class="but">
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'" colspan="2">��&nbsp;&nbsp; ��</td>
                                <td height="36" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'" rowspan="2" width="315">
                                    <p align="justify">
                                        ˵��:1����������������ǧ���������ó���ѡ�񸴺ϳ��䣻<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp; 2�����踴�ϵģ���ѡ��&quot;�컯��&quot;����&nbsp;�� 
			<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp; 3�������������������б��е��Ʒ���������޸ġ�</td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'" width="315">����֧��</td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'" width="80">Ա��</td>
                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <input name="Loan_oan" type="text" style="ime-mode: disabled" size="9" value='0' onfocus="if(this.value=='0')this.value=''" onblur="if(this.value=='')this.value='0'"><font color="#FFFFFF">ǧ��</font></td>



                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'" width="315">
                                    <input name="bang" style="ime-mode: disabled" type="text" size="5" value='0' onfocus="if(this.value=='0')this.value=''" onblur="if(this.value=='')this.value='0'"></td>
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'" width="80">
                                    <input name="pihao" type="text" readonly="readonly " size="6" value="<%=name%>"></td>



                            </tr>
                            <tr>
                                <td colspan="5" align="center">
                                    <input type="submit" value="<% if isedit then
		                                                         response.write "�༭"
																 else
																 response.write "���"
																  end if %>">
                                </td>
                            </tr>
                        </form>

                    </table>
                    <%end if%>


                </td>
            </tr>
    </table>

    <br>
    <br>
    <table width="800" align="center" border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
        <tr>
            <td>

                <table width="800" border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

                    <tr>

                        <td width="800" align="center" valign="top">
                            <table width="800" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td height="26" background="images/background.gif">
                                        <b><font color="#ffffff">�� ӡ �� �� �� ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
                                            <input type="button" name="button" id="button" onclick="window.open('dindang1.asp')" value="�鿴����" /></b></td>
                                </tr>
                            </table>
                            <form action="cy.asp" method="post" name="selform">
                                <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
                                    <tr>

                                        <td width="100%" align="center" valign="top">

                                            <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
                                                <tr class="but">
                                                    <td width="8%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">�������</td>
                                                    <td width="48%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">Ʒ ��</td>
                                                    <td width="7%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">�� ��</td>
                                                    <td width="8%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">����</td>
                                                    <td width="4%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">�� λ</td>
                                                    <td width="4%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">����</td>
                                                    <td width="8%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">����</td>
                                                    <td width="6%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">Ա��</td>
                                                    <td width="7%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">����</td>
                                                </tr>
                                                <%
PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 20          'ÿҳ��ʾ������
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet") 
search=request("search")

if search<>"" then 
keyword=request("keyword")

rs.Source="select* from caiyin_in where pihao like '%"&keyword&"%'  order by -id "
else
rs.Source="select* from caiyin_In order by -id"
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
                                                <tr onmouseover="this.style.backgroundColor='#73A475';" onmouseout="this.style.backgroundColor='';">
                                                    <td height="18" align="center"><%=rs("uptime")%></td>
                                                    <td height="18" align="center"><a href="caiyin_shc.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("cname")%></a></td>
                                                    <td height="18" align="center">&nbsp;<%=rs("guige")%>*<%=rs("houdou")%></td>
                                                    <td height="18" align="center"><%=Formatnumber(rs("loan_oan"),2,-1,0,0)%><%if rs("loan_oan")="" then%>&nbsp;<%end if%></td>
                                                    <td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
                                                    <td height="18" align="center"><%=rs("bang")%></td>
                                                    <td height="18" align="center"><%=rs("use_dep")%></td>
                                                    <td height="18" align="center"><%=rs("pihao")%></td>
                                                    <td height="18" align="center">
                                                        <input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
                                                </tr>
                                                <%
rs.MoveNext
end if
next
end if
                                                %>
                                                <tr>
                                                    <td align="right" colspan="9" height="22">�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="cy.asp?keyword="&keyword&"&search=search&"
			else
            url="cy.asp?"
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
            %>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="chejian" then %>
                                                        <input type="checkbox" name="checkbox" value="checkbox" onclick="javascript:SelectAll()">
                                                        ѡ��/��ѡ
              <input onclick="{if(confirm('ȷ��Ҫɾ���ü�¼��')){this.document.selform.submit();return true;}return false;}" type="submit" value="ɾ��" name="action2">
                                                        <input type="Hidden" name="action" value='del'><%end if%></td>
                                                </tr>
                                            </table>


                                        </td>
                                    </tr>
                                </table>
                            </form>


                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <br>
    <table width="800" align="center" border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
        <tr>
            <td>


                <table width="800" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td height="26" background="images/background.gif">
                            <b><font color="#ffffff">�������</font></b>

                        </td>
                    </tr>
                </table>

                <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">

                    <form name="search" method="post" action="cy.asp">
                        <tr>
                            <input name="search" type="hidden" value="search">
                            <td align="center" width="84%">���������ң� 
			<select name='keyword' style='width: 130'>
                <option selected value="">��ѡ������</option>
                <%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pihao) from caiyin_in order by pihao",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
                %>
                <option value="<%=rs2("pihao")%>"><%=rs2("pihao")%></option>
                <%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
                %>
                                </select></td>

                            <td align="center" width="17%">
                                <input type="submit" value="�� ��">
                            </td>

                        </tr>
                    </form>
            </td>
        </tr>
    </table>
    <table width="100%" border="0" cellspacing="0" cellpadding="4">
    </table>
    </TBODY>  
</TABLE>
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
