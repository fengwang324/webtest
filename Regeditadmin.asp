
<!-- #include file="conn.asp" -->
<%
if session("reg")="1" then
else
	response.write "�뷵�أ�ˢ�£�������д��"
	response.end
end if
dim password,password_prompt,password_Answer,vipno,truename,province,city,telphone1,telphone2,fax,mobile,address,postno
dim email,babysex,babybirthday
dim sql,rs
	
username=trim(request.form("username")) 
password=trim(request.form("password"))
password_prompt=trim(request.form("password_prompt")) 
password_Answer=trim(request.form("password_Answer")) 
vipno=trim(request.form("vipno")) 
truename=trim(request.form("truename")) 

province=trim(request.form("province")) 
city=trim(request.form("city"))
area=trim(request.form("area"))

telphone1=trim(request.form("telphone1")) 
telphone2=trim(request.form("telphone2")) 
fax=trim(request.form("fax")) 
mobile=trim(request.form("mobile")) 
address=trim(request.form("address")) 
postno=trim(request.form("postno")) 
email=trim(request.form("email")) 
babysex=trim(request.form("babysex")) 
babybirthday=trim(request.form("babybirthday"))

if province=city then
	address=city&" "&area&" "&address
else
	address=province&" "&city&" "&area&" "&address
end if
if address<>"" then
	address=replace(address,"ʡֱϽ������λ","")
	address=replace(address,"ʡֱϽ�ؼ�������λ","")
end if


if InStr(username,"�ǻ�Ա")>0 or InStr(username,"'")>0 or InStr(username,"--")>0 or InStr(username,"(")>0 or InStr(username,";")>0 or InStr(username,"#")>0 then
	response.write "�û���ҪΪ4-20���ַ���a-z��A-Z��0-9��"
	response.end
END IF
if InStr(password,"'")>0 or InStr(password,"--")>0 or InStr(password,"(")>0 or InStr(password,";")>0 then
	response.write "����ҪΪ4-20���ַ���0-9��a-z��A-Z���»��ߣ�"
	response.end
END IF


sql="select * from x_huiyuan where username='"&username&"' "

'response.write sql
'response.end

set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof or rs.eof then
	rs.addnew
	rs("username")=username 
	rs("password")=password 
	rs("password_prompt")=password_prompt 
	rs("password_Answer")=password_Answer 
	rs("vipno")=vipno 
	rs("truename")=truename 
	rs("province")=province 
	rs("city")=city
	rs("self1")=area
	rs("telphone1")=telphone1 
	rs("telphone2")=telphone2 
	rs("fax")=fax 
	rs("mobile")=mobile 
	rs("address")=address 
	rs("postno")=postno 
	rs("email")=email 
	rs("babysex")=babysex 
	rs("babybirthday")=babybirthday
	rs.update
	
	session("username")=username
	session("s_stat")="0"
	session("customkind")="2"
	session("customkindname")="��ͨ��Ա"
	
	rs.close
	conn.close	
	
	if request.form("kuai_reg")="1" then
	response.write "<script language=JavaScript>" & "alert('��ϲ��ע��ɹ��������ȷ��������ȥ���㡣');" & "window.location.href='orderjs.asp'" & "</script>"
	else
	response.write "<script language=JavaScript>" & "alert('��ϲ��ע��ɹ������ѳɱ���վ����ͨ��Ա���������ܡ���ͨ��Ա�����Żݡ�');" & "window.location.href='index.asp'" & "</script>"
	end if
	
	RESPONSE.END

else
	response.write "<script language=JavaScript>" & "alert('�Բ���ע��ʧ��01��������д�ġ��û��������Ļ�Ա�û�����ͬ��\n����ȷ����������д�û�����');" & "window.history.back()" & "</script>"
	
end if

rs.close
set rs=nothing
conn.close
set conn=nothing
	


	
%>



