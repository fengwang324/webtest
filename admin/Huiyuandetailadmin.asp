<!-- #include file="conn.asp" -->
<!--#include file="check4.asp"-->
<%
dim username,password,password_prompt,password_Answer,vipno,truename,province,city,telphone1,telphone2,fax,mobile,address,postno
dim email,babysex,babybirthday
dim sql,rs
	
id=trim(request.form("id")) 
password=trim(request.form("password"))
password_prompt=trim(request.form("password_prompt")) 
password_Answer=trim(request.form("password_Answer")) 
vipno=trim(request.form("vipno")) 
truename=trim(request.form("truename")) 
province=trim(request.form("province")) 
city=trim(request.form("city")) 
telphone1=trim(request.form("telphone1")) 
telphone2=trim(request.form("telphone2")) 
fax=trim(request.form("fax")) 
mobile=trim(request.form("mobile")) 
address=trim(request.form("address")) 
postno=trim(request.form("postno")) 
email=trim(request.form("email")) 
babysex=trim(request.form("babysex")) 
babybirthday=trim(request.form("babybirthday"))
s_stat=trim(request.form("s_stat"))
customkind=trim(request.form("customkind"))

if InStr(password,"'")>0 or InStr(password,"--")>0 or InStr(password,"(")>0 or InStr(password,";")>0 then
	response.write "密码要为4-20个字符（0-9，a-z，A-Z，下划线）"
	response.end
END IF


sql="select * from x_huiyuan where id="&id&" "

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
	response.write "<script language=JavaScript>" & "alert('对不起，修改失败。请重新尝试。');" & "window.history.back()" & "</script>"
	response.end
else
	'rs.addnew
	'rs("username")=username 
	rs("password")=password 
	rs("password_prompt")=password_prompt 
	rs("password_Answer")=password_Answer 
	rs("vipno")=vipno 
	rs("truename")=truename 
	rs("province")=province 
	rs("city")=city 
	rs("telphone1")=telphone1 
	rs("telphone2")=telphone2 
	rs("fax")=fax 
	rs("mobile")=mobile 
	rs("address")=address 
	rs("postno")=postno 
	rs("email")=email 
	rs("babysex")=babysex 
	rs("babybirthday")=babybirthday
	rs("s_stat")=s_stat
	rs("customkind")=customkind
	rs.update
	response.redirect (request.cookies("hypage"))

end if

rs.close
set rs=nothing
conn.close
set conn=nothing	
%>


