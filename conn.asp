<%
starttime=timer()
ConnStr = "driver={SQL Server};server=192.168.5.77;uid=sa;pwd=Lgm731003;database=JiaMeiDB_20250518"
On Error Resume Next
Set conn = Server.CreateObject("ADODB.Connection")
conn.open ConnStr
If Err Then
	err.Clear
	Set Conn = Nothing
	Response.Write "数据库连接出错。"
	Response.End
End If
%>
