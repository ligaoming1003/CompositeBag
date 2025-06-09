<!--#include file="Conn.Asp"-->
<%
id=request("id")
FormSize = Request.TotalBytes '得到数据
FormData = Request.BinaryRead( FormSize )
function ImageUp(formsize,formdata) '这个函数的功能是截取其中的图像部分。
bncrlf=chrb(13) & chrb(10) '做成函数后。以后你可以自己随意使用了。
divider=leftb(formdata,instrb(formdata,bncrlf)-1)
datastart=instrb(formdata,bncrlf&bncrlf)+4
dataend=instrb(datastart+1,formdata,divider)-datastart
imageup=midb(formdata,datastart,dataend)
end function
Image=ImageUp (FormSize,Formdata) '这里就是图像部分了。
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from ImgData where id="&id
rs.open sql,conn,3,3
rs("img").appendchunk Image '所以这里是添加。
rs.Update
rs.Close
response.contenttype="image/gif" 
response.binarywrite imageup(formsize,formdata) '这里是显示图像。表示成功！

rs.close
set rs=nothing
response.Write "<script language=javascript>alert('添加成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=guolu.asp"">"
response.end
rs.close
set rs=nothing
%>
