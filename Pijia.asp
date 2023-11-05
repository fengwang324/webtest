<!-- #include file="conn.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<%
if session("pijia")="1" then
else
	response.write "请返回，刷新，重新填写。"
	response.end
end if

dim sql,rs,yourname,tel,memo,productid

productid=trim(request.form("productid")) 	
yourname=trim(request.form("yourname")) 
tel=trim(request.form("tel"))
memo=trim(request.form("memo")) 

if yourname="" then yourname="(匿名)"
if tel="" then tel="-"

s = trim(request.Form("memo"))
Set oReg    = New RegExp
With oReg
	.IgnoreCase = False
	.Global     = True
	.Pattern    = "[\u4E00-\u9FFF]+"
	If .Test(s) Then
		chinese="1"
	End If 
End With

if memo<>"" then
memo=replace(memo,"iframe","")
memo=replace(memo,"script","")
end if

if chinese="1" then    '有begin
sql="select * from z_pijia "
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
rs("productid")=productid 
rs("yourname")=yourname 
rs("tel")=tel 
rs("memo")=memo 
rs.update
rs.close
set rs=nothing
end if    '有end

conn.close
set conn=nothing

response.write "<script language=JavaScript>" & "alert('成功保存您的评价。谢谢！');" & "window.location.href='show.asp?pkid="&productid&"'" & "</script>"

%>



