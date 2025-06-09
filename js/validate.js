function isEmail(vEMail)
{
	var regInvalid=/(@.*@)|(\.\.)|(@\.)|(\.@)|(^\.)/;
	var regValid=/^.+\@(\[?)[a-zA-Z0-9\-\.]+\.([a-zA-Z]{2,3}|[0-9]{1,3})(\]?)$/;
	return (!regInvalid.test(vEMail)&&regValid.test(vEMail));
}
function CheckForm(theform)
{
if(theform.username.value==""){
alert("请输入你的昵称");
return false;
}
if(theform.old.value==""){
alert("年龄是必填的信息");
return false;
}
if(theform.sex.value==""){
alert("性别是必填的信息");
return false;
}
if((theform.pwd.value=="")||(theform.pwd1.value==""))
{
alert("请输入密码！");
return false;
}
if(theform.pwd.value.length<4)
{
alert("密码不能小于四位！");
return false;
}
if(theform.pwd.value!=theform.pwd1.value)
{
alert("两次输入的密码不一样");
return false;
}
if(theform.turenaem.value==""){
alert("请输入你的真实性名");
return false;
}
if(theform.mail.value!=="")
 {
		if(!isEmail(theform.mail.value))
       {
			alert("请输入正确的邮件格式!");
			return false;
			theform.email.focus();
		}
 }
}