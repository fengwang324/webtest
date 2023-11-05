<% 
public mydataname
public mydatapsw
public must_login
public managerfolder
public newwindow
dim path_back

mydataname = "#ioaskl089235.mdb"		'设置数据库名，即是DATA文件夹下的数据库名。
mydatapsw = ""                    	'设置数据库密码。如果在数据库里设了密码，在这里修改即可。

must_login="1"				'1为必须为会员登陆后才可下订单，0为不需登录而以游客的身份购买

managerfolder = "admin" 	'后台管理文件夹名
'如果这里修改，则要将admin文件夹改为相应名，同时后台路径也相应改变


newwindow=""	  '点击商品名称或图片是否新开窗口, 设为newwindow="target='_blank'"时新开；设为newwindow=""时有原窗口打开

path_back=LCase(request.servervariables("QUERY_STRING"))
if instr(path_back,"insert")<>0 or instr(path_back,"update")<>0 or instr(path_back,"delete")<>0 or instr(path_back,"(")<>0 or instr(path_back,"'")<>0 or instr(path_back," or ")<>0 or instr(path_back,"replace")<>0 or instr(path_back,"eval")<>0 then
response.write "<BR><BR><center>你的网址不合法</a>"
response.end
end if
%>
