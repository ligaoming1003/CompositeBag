<html>
<head>

<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript src="css/init.js"></SCRIPT>

<style type="text/css">
<!--
td {  font-family: "����"; font-size: 9pt}
body {  font-family: "����"; font-size: 9pt}
select {  font-family: "����"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "����"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "����"; font-size: 9pt} 
-->
</style>
<script type="text/javascript">
function js1()
{
 
 cd1=document.getElementById("cd").value;
 kd1=document.getElementById("kd").value;
 
 mdj1=document.getElementById("mdj").value;
 mhd1=document.getElementById("mhd").value;
 mcn1=document.getElementById("mcn").value;
 
 zcn1=document.getElementById("zcn").value;
 zhd1=document.getElementById("zhd").value;
 zdj1=document.getElementById("zdj").value;
 
 scn1=document.getElementById("scn").value;
 shd1=document.getElementById("shd").value;
 sdj1=document.getElementById("sdj").value;
 
 ncn1=document.getElementById("ncn").value;
 nhd1=document.getElementById("nhd").value;
 ndj1=document.getElementById("ndj").value;

 zk1=document.getElementById("zk").value;
 zj1=document.getElementById("zj").value;
 
 ysh1=document.getElementById("ysh").value;
 ym1=document.getElementById("ym").value;
 
 fh1=document.getElementById("fh").value;
 jsh1=document.getElementById("jsh").value;
 fhy1=document.getElementById("fhy").value;
 
 dx1=document.getElementById("dx").value;
 cdi1=document.getElementById("cdi").value;
 lal1=document.getElementById("lal").value;

 zei1=document.getElementById("zei").value;
 tihuang1=document.getElementById("tihuang").value;
 kai1=document.getElementById("kai").value;

 guige=cd1*2/100*kd1/100*1.03
 ggg=guige*mhd1/10*mcn1+guige*zhd1/10*zcn1+guige*shd1/10*scn1+guige*nhd1/10*ncn1+guige*zk1/1000
 cail=guige*mhd1/10*mdj1*mcn1+guige*zhd1/10*zdj1*zcn1+guige*shd1/10*sdj1*scn1+guige*nhd1/10*ndj1*ncn1+guige*zk1/1000*zj1
 caiyin=guige*0.1+guige*0.2*ysh1*0.02+guige*ym1
 fuhe=guige*0.2*fh1+guige*0.1*fh1*jsh1+guige*0.1*fhy1*fh1
 daxiao=guige*0.12*dx1+guige*cdi1+cd1/100*0.2*lal1
 zhidai=daxiao+0.07/guige*guige*0.12*dx1
 wang=1*zei1+0.00001
 teshu=0.1*zei1*zei1/2*zei1/2+zei1/wang*0.08+1*tihuang1+kai1*cd1/100*0.14*0.8*15
 bd=cail+fuhe+zhidai+teshu


 "�״�"!=document.getElementById("bzh").value?hj=bd+caiyin:hj=bd;
 "�ƴ�"!=document.getElementById("zhcai").value?hj=hj/ggg+guige*0.2/ggg:hj=hj;

 document.form1.je.value=hj;

}
</script>

</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0>       
<p>��</p>
<TABLE width=600 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top">

	

    <table width="600" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<p align="center">
<b><font color="#ffffff">�Ƽ�����</font></b>

</td>
</tr>
</table>
       <form id="form1" name="form1" method="post" action="">
       <table width="600" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">

          <tr class="but"> 

            <td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" >
			�� ��</td>


			
            <td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" >
			&nbsp;
			���ȣ�<input name="cd" type="text" style="ime-mode:disabled"  id="cd" onBlur="js1()" value="" size="7" />����</td>
<td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" >
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;
			��ȣ�<input name="kd" type="text" style="ime-mode:disabled"  id="kd" onBlur="js1()" value="" size="7" />����</td>
          </tr>
		  
		  <tr class="but"> 
<td  height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			&nbsp;&nbsp;
			��Ĥ&nbsp;
			����
    <select name="mcn" id="mcn" onChange="js1()">
      <option value="0.1">OPP</option>
      <option value="0.15">PET</option>
      <option value="0.13">PA</option>
      <option value="0.09">PT</option>
      <option value="0.1">��˿Ĥ</option>


    </select> 
			</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			��ȣ�<input name="mhd" type="text" style="ime-mode:disabled"  id="mhd" onBlur="js1()" value="" size="7" />˿</td>


<td  height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			���ۣ�<input name="mdj" type="text" style="ime-mode:disabled"  id="mdj" onBlur="js1()" value="" size="7" />Ԫ/����</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			&nbsp;
			
			��Ĥ&nbsp; ���� <select name="zcn" id="zcn" onChange="js1()">
      <option value="0.15">VMPET</option>
      <option value="0.15">PET</option>
      <option value="0.13">PA</option>
      <option value="0.27">��</option>


    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			��ȣ�<input name="zhd" type="text" style="ime-mode:disabled"  id="zhd" onBlur="js1()" value="" size="7" />˿</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			���ۣ�<input name="zdj" type="text" style="ime-mode:disabled"  id="zdj" onBlur="js1()" value="" size="7" />Ԫ/����</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp; ��Ĥ&nbsp; ���� <select name="scn" id="scn" onChange="js1()">
      <option value="0.15">VMPET</option>
      <option value="0.15">PET</option>
      <option value="0.13">PA</option>
      <option value="0.27">��</option>



    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			��ȣ�<input name="shd" type="text" style="ime-mode:disabled"  id="shd" onBlur="js1()" value="" size="7" />˿</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			���ۣ�<input name="sdj" type="text" style="ime-mode:disabled"  id="sdj" onBlur="js1()" value="" size="7" />Ԫ/����</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp;&nbsp; ֽ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �� <select name="zz" id="zz" onChange="js1()">
      <option value="0.095">ţƤֽ</option>
      <option value="0.093">��Ĥֽ</option>
    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			&nbsp;&nbsp;&nbsp;&nbsp;
			
			������<input name="zk" type="text" style="ime-mode:disabled"  id="zk" onBlur="js1()" value="" size="7" />��/ƽ��</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			���ۣ�<input name="zj" type="text" style="ime-mode:disabled"  id="zj" onBlur="js1()" value="" size="7" />Ԫ/����</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp; ��Ĥ&nbsp; ���� <select name="ncn" id="ncn" onChange="js1()">
      <option value="0.1">CPP</option>
      <option value="0.1">RCPP</option>
      <option value="0.1">VMCPP</option>
      <option value="0.1">PE</option>
      <option value="0.087">���</option>
      <option value="0.1">PAPP</option>



    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			��ȣ�<input name="nhd" type="text" style="ime-mode:disabled"  id="nhd" onBlur="js1()" value="" size="7" />˿</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			���ۣ�<input name="ndj" type="text" style="ime-mode:disabled"  id="ndj" onBlur="js1()" value="" size="7" />Ԫ/����</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<p align="left">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ӡˢ\�״� <select name="bzh" id="bzh" onChange="js1()">
      <option value="�״�">�״�</option>
      <option value="�ʴ�">�ʴ�</option>
    </select>&nbsp;&nbsp; </td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ��ɫ��<select name="ysh" id="ysh" onChange="js1()">
      <option value="0">��ѡ��</option>
      <option value="1">һɫ</option>
      <option value="2">��ɫ</option>
      <option value="3">��ɫ</option>
      <option value="4">��ɫ</option>
      <option value="5">��ɫ</option>
      <option value="6">��ɫ</option>
      <option value="7">��ɫ</option>
      <option value="8">��ɫ</option>
      <option value="9">��ɫ</option>



    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp; ��ī��<select name="ym" id="ym" onChange="js1()">
      <option value="0">��ѡ��</option>
      <option value="0">��ͨī</option>
      <option value="0.05">����ī</option>
      <option value="0.03">ˮ��ī</option>
      <option value="0.05">���ī</option>
      <option value="0.1">���ī</option>
      

    </select></td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp;&nbsp; ����&nbsp; ���� <select name="fh" id="fh" onChange="js1()">
      <option value="0">��ѡ��</option>
      <option value="1.3">һ��</option>
      <option value="2.2">����</option>
      <option value="3.5">����</option>

    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			
			��ˮ��<select name="jsh" id="jsh" onChange="js1()">
      <option value="0">��ѡ��</option>
      <option value="0">��ͨ</option>   
      <option value="0.6">����</option>
      <option value="0.15">ˮ��</option>
      
      
    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp; ���ͣ�<select name="fhy" id="fhy" onChange="js1()">
      <option value="0">��ѡ��</option>
      <option value="0">��ͨ��</option>
      <option value="0.5">������</option>
      <option value="0.8">����</option>
      

    </select></td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp;&nbsp; �ƴ�&nbsp; ���� 
			<select name="dx" id="dx" onChange="js1()">
      <option value="0">��ѡ��</option>
      <option value="1">���߷�</option>
      <option value="1.3">�۱ߴ�</option>
      <option value="1.1">�з��</option>

      </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			
			��ף�<select name="cdi" id="cdi" onChange="js1()">
      <option value="0">��</option>
      <option value="0.1">��</option>
            
    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;
			������<select name="lal" id="lal" onChange="js1()">
      <option value="0">��</option>
      <option value="1">��</option>
            
    </select></td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="left" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
		<p align="left">&nbsp;&nbsp;����(�ܿ��ھ�)��<input type="text" name="zei" style="ime-mode:disabled"  id="zei" onBlur="js1()" value="" size="6" />����</td>



<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �ỷ��<select name="tihuang" id="tihuang" onChange="js1()">
      <option value="0">��</option>
      <option value="0.235">��ѹ��</option>
      <option value="0.2">�ں���</option>    
    </select>
			</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<p align="left">&nbsp;ֽ��������<select name="kai" id="kai" onChange="js1()">
      <option value="0">��</option>
      <option value="0.1">10cm</option>
      <option value="0.08">8cm</option>
      <option value="0.06">6cm</option>  
      <option value="0.04">4cm</option>   
    </select>		
</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="3">
&nbsp;			
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br>
&nbsp;�ƴ������ 
			<select name="zhcai" id="zhcai" onChange="js1()">
      <option value="�ƴ�">�ƴ�</option>
      <option value="���">���</option>
    </select>&nbsp; <br>
��</td>


          </tr>
		  
		  <tr> 
            
			<td align="center" colspan="3" height="25">���
    <input name="je" type="text" value="" id="je" />
 </td>
          </tr>
		  
           </table>
           </form>
        </td>
</tr>
</table>  </TBODY>
<br><br>
<br><br>

          
 

</TABLE>

<P>
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <TBODY>
  <TR vAlign=center>
    <TD align=middle width="100%"></TD>
  </TR></TBODY></TABLE></P></BODY></HTML>