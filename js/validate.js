function isEmail(vEMail)
{
	var regInvalid=/(@.*@)|(\.\.)|(@\.)|(\.@)|(^\.)/;
	var regValid=/^.+\@(\[?)[a-zA-Z0-9\-\.]+\.([a-zA-Z]{2,3}|[0-9]{1,3})(\]?)$/;
	return (!regInvalid.test(vEMail)&&regValid.test(vEMail));
}
function CheckForm(theform)
{
if(theform.username.value==""){
alert("����������ǳ�");
return false;
}
if(theform.old.value==""){
alert("�����Ǳ������Ϣ");
return false;
}
if(theform.sex.value==""){
alert("�Ա��Ǳ������Ϣ");
return false;
}
if((theform.pwd.value=="")||(theform.pwd1.value==""))
{
alert("���������룡");
return false;
}
if(theform.pwd.value.length<4)
{
alert("���벻��С����λ��");
return false;
}
if(theform.pwd.value!=theform.pwd1.value)
{
alert("������������벻һ��");
return false;
}
if(theform.turenaem.value==""){
alert("�����������ʵ����");
return false;
}
if(theform.mail.value!=="")
 {
		if(!isEmail(theform.mail.value))
       {
			alert("��������ȷ���ʼ���ʽ!");
			return false;
			theform.email.focus();
		}
 }
}