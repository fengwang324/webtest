<% 
public mydataname
public mydatapsw
public must_login
public managerfolder
public newwindow
dim path_back

mydataname = "#ioaskl089235.mdb"		'�������ݿ���������DATA�ļ����µ����ݿ�����
mydatapsw = ""                    	'�������ݿ����롣��������ݿ����������룬�������޸ļ��ɡ�

must_login="1"				'1Ϊ����Ϊ��Ա��½��ſ��¶�����0Ϊ�����¼�����ο͵���ݹ���

managerfolder = "admin" 	'��̨�����ļ�����
'��������޸ģ���Ҫ��admin�ļ��и�Ϊ��Ӧ����ͬʱ��̨·��Ҳ��Ӧ�ı�


newwindow=""	  '�����Ʒ���ƻ�ͼƬ�Ƿ��¿�����, ��Ϊnewwindow="target='_blank'"ʱ�¿�����Ϊnewwindow=""ʱ��ԭ���ڴ�

path_back=LCase(request.servervariables("QUERY_STRING"))
if instr(path_back,"insert")<>0 or instr(path_back,"update")<>0 or instr(path_back,"delete")<>0 or instr(path_back,"(")<>0 or instr(path_back,"'")<>0 or instr(path_back," or ")<>0 or instr(path_back,"replace")<>0 or instr(path_back,"eval")<>0 then
response.write "<BR><BR><center>�����ַ���Ϸ�</a>"
response.end
end if
%>
