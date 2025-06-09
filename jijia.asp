<html>
<head>

<title>∷嘉美管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=javascript src="css/init.js"></SCRIPT>

<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 9pt}
body {  font-family: "宋体"; font-size: 9pt}
select {  font-family: "宋体"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "宋体"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
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


 "白袋"!=document.getElementById("bzh").value?hj=bd+caiyin:hj=bd;
 "制袋"!=document.getElementById("zhcai").value?hj=hj/ggg+guige*0.2/ggg:hj=hj;

 document.form1.je.value=hj;

}
</script>

</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0>       
<p>　</p>
<TABLE width=600 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top">

	

    <table width="600" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<p align="center">
<b><font color="#ffffff">计价运算</font></b>

</td>
</tr>
</table>
       <form id="form1" name="form1" method="post" action="">
       <table width="600" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">

          <tr class="but"> 

            <td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" >
			规 格</td>


			
            <td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" >
			&nbsp;
			长度：<input name="cd" type="text" style="ime-mode:disabled"  id="cd" onBlur="js1()" value="" size="7" />厘米</td>
<td width="200" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" >
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;
			宽度：<input name="kd" type="text" style="ime-mode:disabled"  id="kd" onBlur="js1()" value="" size="7" />厘米</td>
          </tr>
		  
		  <tr class="but"> 
<td  height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			&nbsp;&nbsp;
			面膜&nbsp;
			材料
    <select name="mcn" id="mcn" onChange="js1()">
      <option value="0.1">OPP</option>
      <option value="0.15">PET</option>
      <option value="0.13">PA</option>
      <option value="0.09">PT</option>
      <option value="0.1">拉丝膜</option>


    </select> 
			</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			厚度：<input name="mhd" type="text" style="ime-mode:disabled"  id="mhd" onBlur="js1()" value="" size="7" />丝</td>


<td  height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			单价：<input name="mdj" type="text" style="ime-mode:disabled"  id="mdj" onBlur="js1()" value="" size="7" />元/公斤</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			&nbsp;
			
			中膜&nbsp; 材料 <select name="zcn" id="zcn" onChange="js1()">
      <option value="0.15">VMPET</option>
      <option value="0.15">PET</option>
      <option value="0.13">PA</option>
      <option value="0.27">铝</option>


    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			厚度：<input name="zhd" type="text" style="ime-mode:disabled"  id="zhd" onBlur="js1()" value="" size="7" />丝</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			单价：<input name="zdj" type="text" style="ime-mode:disabled"  id="zdj" onBlur="js1()" value="" size="7" />元/公斤</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp; 三膜&nbsp; 材料 <select name="scn" id="scn" onChange="js1()">
      <option value="0.15">VMPET</option>
      <option value="0.15">PET</option>
      <option value="0.13">PA</option>
      <option value="0.27">铝</option>



    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			厚度：<input name="shd" type="text" style="ime-mode:disabled"  id="shd" onBlur="js1()" value="" size="7" />丝</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			单价：<input name="sdj" type="text" style="ime-mode:disabled"  id="sdj" onBlur="js1()" value="" size="7" />元/公斤</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp;&nbsp; 纸&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 张 <select name="zz" id="zz" onChange="js1()">
      <option value="0.095">牛皮纸</option>
      <option value="0.093">凝膜纸</option>
    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			&nbsp;&nbsp;&nbsp;&nbsp;
			
			克数：<input name="zk" type="text" style="ime-mode:disabled"  id="zk" onBlur="js1()" value="" size="7" />克/平米</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			单价：<input name="zj" type="text" style="ime-mode:disabled"  id="zj" onBlur="js1()" value="" size="7" />元/公斤</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp; 内膜&nbsp; 材料 <select name="ncn" id="ncn" onChange="js1()">
      <option value="0.1">CPP</option>
      <option value="0.1">RCPP</option>
      <option value="0.1">VMCPP</option>
      <option value="0.1">PE</option>
      <option value="0.087">珠光</option>
      <option value="0.1">PAPP</option>



    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			厚度：<input name="nhd" type="text" style="ime-mode:disabled"  id="nhd" onBlur="js1()" value="" size="7" />丝</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			单价：<input name="ndj" type="text" style="ime-mode:disabled"  id="ndj" onBlur="js1()" value="" size="7" />元/公斤</td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<p align="left">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 印刷\白袋 <select name="bzh" id="bzh" onChange="js1()">
      <option value="白袋">白袋</option>
      <option value="彩袋">彩袋</option>
    </select>&nbsp;&nbsp; </td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 颜色：<select name="ysh" id="ysh" onChange="js1()">
      <option value="0">请选择</option>
      <option value="1">一色</option>
      <option value="2">二色</option>
      <option value="3">三色</option>
      <option value="4">四色</option>
      <option value="5">五色</option>
      <option value="6">六色</option>
      <option value="7">七色</option>
      <option value="8">八色</option>
      <option value="9">九色</option>



    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp; 油墨：<select name="ym" id="ym" onChange="js1()">
      <option value="0">请选择</option>
      <option value="0">普通墨</option>
      <option value="0.05">蒸煮墨</option>
      <option value="0.03">水煮墨</option>
      <option value="0.05">半金墨</option>
      <option value="0.1">多金墨</option>
      

    </select></td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp;&nbsp; 复合&nbsp; 层数 <select name="fh" id="fh" onChange="js1()">
      <option value="0">请选择</option>
      <option value="1.3">一层</option>
      <option value="2.2">二层</option>
      <option value="3.5">三层</option>

    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			
			胶水：<select name="jsh" id="jsh" onChange="js1()">
      <option value="0">请选择</option>
      <option value="0">普通</option>   
      <option value="0.6">蒸煮</option>
      <option value="0.15">水煮</option>
      
      
    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp; 袋型：<select name="fhy" id="fhy" onChange="js1()">
      <option value="0">请选择</option>
      <option value="0">普通袋</option>
      <option value="0.5">铝箔袋</option>
      <option value="0.8">特殊</option>
      

    </select></td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			&nbsp;&nbsp; 制袋&nbsp; 袋别 
			<select name="dx" id="dx" onChange="js1()">
      <option value="0">请选择</option>
      <option value="1">三边封</option>
      <option value="1.3">折边袋</option>
      <option value="1.1">中封袋</option>

      </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			
			插底：<select name="cdi" id="cdi" onChange="js1()">
      <option value="0">否</option>
      <option value="0.1">是</option>
            
    </select></td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;
			拉链：<select name="lal" id="lal" onChange="js1()">
      <option value="0">否</option>
      <option value="1">是</option>
            
    </select></td>


          </tr>
		  
		  <tr class="but"> 
<td height="18" align="left" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
		<p align="left">&nbsp;&nbsp;吸嘴(管口内径)：<input type="text" name="zei" style="ime-mode:disabled"  id="zei" onBlur="js1()" value="" size="6" />厘米</td>



<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" >
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 提环：<select name="tihuang" id="tihuang" onChange="js1()">
      <option value="0">无</option>
      <option value="0.235">外压扣</option>
      <option value="0.2">内焊扣</option>    
    </select>
			</td>


<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<p align="left">&nbsp;纸袋开窗：<select name="kai" id="kai" onChange="js1()">
      <option value="0">无</option>
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
&nbsp;制袋＆卷材 
			<select name="zhcai" id="zhcai" onChange="js1()">
      <option value="制袋">制袋</option>
      <option value="卷材">卷材</option>
    </select>&nbsp; <br>
　</td>


          </tr>
		  
		  <tr> 
            
			<td align="center" colspan="3" height="25">金额
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