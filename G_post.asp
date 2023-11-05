<!--#include file="conn.asp"-->
<%
if s_head="head4.asp" or s_productkind="4" then
	response.write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
else
	response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
end if
%>
<head>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<title><%=application("sitename")%></title>
<link href="i.css" type=text/css rel=stylesheet>
<style type="text/css">
<!--

.style1 {font-size: 14px}
.style11 {
	font-size: 14px;
	color: #990000;
	font-weight: bold;
}
.style12 {color: #990000}
.style13 {color: #333333}
#tabcss{
	BORDER-COLLAPSE: collapse;
	border-top-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-left-style: solid;
	border-top-color: #dddddd;
	border-left-color: #dddddd;
}#tabcss td {
	line-height: 24px;
	padding-left: 3px;
	padding-right: 2px;
	padding-top: 2px;
	padding-bottom: 2px;
	color: #333333;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-right-style: solid;
	border-bottom-style: solid;
	border-right-color: #dddddd;
	border-bottom-color: #dddddd;
}
-->
</style>
<SCRIPT language=JavaScript>
<!--

function Juge()
{
  if (theForm.g_title.value == "")
  {
    alert("请填写留言标题，以便检索!");
    theForm.g_title.focus();
    return (false);
  }
  	  if (theForm.g_content.value == "")
  {
    alert("留言内容不能为空!");
    theForm.g_content.focus();
    return (false);
  }

 if (isNaN(document.theForm.g_qq.value)) {
	  alert("oicq必须为数字,其实您可以不填");
	  theForm.g_qq.focus();
	  return false;
	}
 if(theForm.g_mail.value.length<=0)
      {
   return true;
       }
	var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_@.";
	var checkStr = theForm.g_mail.value;
	var allValid = true;
	for (i = 0;  i < checkStr.length;  i++)
	{
		ch = checkStr.charAt(i);
		for (j = 0;  j < checkOK.length;  j++)
		if (ch == checkOK.charAt(j))
			break;
		if (j == checkOK.length)
		{
			allValid = false;
			break;
		}
	}

	if (theForm.g_mail.value.length < 6)
	{
			allValid = false;
	}

	if (!allValid)
	{
		alert("您输入的 \"电子邮件地址\" 无效,其实您可以不填!");
		theForm.g_mail.focus();
		return (false);
	}

	address=theForm.g_mail.value;
    if(address.length>0)
	{
        i=address.indexOf("@");
        if(i==-1)
		{
			window.alert("对不起！电子邮件错误，其实您可以不填！")
			theForm.g_mail.focus();
			return false
        }
       ii=address.indexOf(".")
        if(ii==-1)
		{
			window.alert("对不起！电子邮件错误，其实您可以不填！")
			theForm.g_mail.focus();
			return false
        }

    }

	}
//-->
</SCRIPT>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>

<body leftmargin="0" topmargin="0">
<!--#include file="head.asp"-->
<TABLE width="1020" border=0 align="center" cellPadding=0 cellSpacing=0 bgcolor="#FFFFFF">
  <tr>
    <td bgcolor="#fafafa"><br>
      <table width="1020" border="0" align="center" cellpadding="0" cellspacing="0">
<td valign="top">
<%

	if trim(request.form("g_content"))="" and trim(request.form("submit"))<>"发 表" then
%>
	<FORM name="theForm" onSubmit="return Juge()" action="g_post.asp" method="post">
              
              <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" id="tabcss">
                <TBODY>
                  <TR> 
                    <TD height="38" colSpan=2 align=middle> <div align="center" class="style1"><strong>
					<%
					if request.querystring("re")="1" then
						response.write "回 复 留 言&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					else
						response.write "欢 迎 您 留 言&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					end if
					
					%>
					</strong></div></TD>
            </TR>
            <TR> 
              <TD bgColor=#D0E8FF width="12%"> <font color="#FF0000">*</font>留言标题：</TD>
              <TD bgcolor="#f1f9ec"> 
			  <%
			  title=request("g_title")
			  if title <>"" then
			  response.Write"<INPUT name=g_title size=58 value=Re:"& title &">"
			  else
			  response.Write"<INPUT name=g_title size=58>"
			  end if
			  %>			  
              </TD>
            </TR>
            <TR> 
              <TD bgColor=#D0E8FF width="12%">留言表情：</TD>
              <TD bgcolor="#f1f9ec"> <INPUT name=g_brow type=radio value=1 checked> <IMG height=15 
      src="pic/icon1.gif" width=15> <INPUT name=g_brow type=radio 
      value=2> <IMG height=15 src="pic/icon2.gif" width=15> <INPUT 
      name=g_brow type=radio value=3> <IMG height=15 src="pic/icon3.gif" 
      width=15> <INPUT name=g_brow type=radio value=4> <IMG height=15 
      src="pic/icon4.gif" width=15> <INPUT name=g_brow type=radio 
      value=5> <IMG height=15 src="pic/icon5.gif" width=15> <INPUT 
      name=g_brow type=radio value=6> <IMG height=15 src="pic/icon6.gif" 
      width=15> <INPUT name=g_brow type=radio value=7> <IMG height=15 
      src="pic/icon7.gif" width=15> <INPUT name=g_brow type=radio 
      value=8> <IMG height=15 src="pic/icon8.gif" width=15> <INPUT 
      name=g_brow type=radio value=9> <IMG height=15 src="pic/icon9.gif" 
      width=15><BR> <INPUT name=g_brow type=radio value=10> <IMG height=15 
      src="pic/icon10.gif" width=15> <INPUT name=g_brow type=radio 
      value=11> <IMG height=15 src="pic/icon11.gif" width=15> <INPUT 
      name=g_brow type=radio value=12> <IMG height=15 
      src="pic/icon12.gif" width=15> <INPUT name=g_brow type=radio 
      value=13> <IMG height=15 src="pic/icon13.gif" width=15> <INPUT 
      name=g_brow type=radio value=14> <IMG height=15 
      src="pic/icon14.gif" width=15> <INPUT name=g_brow type=radio 
      value=15> <IMG height=15 src="pic/icon15.gif" width=15> <INPUT 
      name=g_brow type=radio value=16> <IMG height=15 
      src="pic/icon16.gif" width=15> <INPUT name=g_brow type=radio 
      value=17> <IMG height=15 src="pic/icon17.gif" width=15> <INPUT 
      name=g_brow type=radio value=18> <IMG height=15 
      src="pic/icon18.gif" width=15></TD>
            </TR>
            <TR> 
              <TD bgColor=#D0E8FF noWrap width="12%"> <font color="#FF0000">*</font>留言内容：</TD>
              <TD bgcolor="#F4FAFF"> <TEXTAREA cols=58 name=g_content rows=10></TEXTAREA> 
              </TD>
            </TR>
			<%
			if request.querystring("re")="1" then
				man=session("m_user")
			elseif session("username")<>"" then
				man=session("username")
			else
				man="[游客]"
			end if
			%>
            <TR><td bgColor=#D0E8FF>&nbsp;</td>
              <TD noWrap bgColor=#f0E8FF> <font color="#FF0000">*</font>昵称：<INPUT name=g_man size=12 value='<%=man%>' readonly style="background-color:#dddddd">
              &nbsp;E-mail：<INPUT name=g_mail size=15> &nbsp;<strong>QQ：<INPUT name=g_qq size=10></strong></TD>
            </TR>
            <TR> 
                    <TD height="37" colSpan=2 align=middle bgcolor="#f1f9ec"> 
					<div align="center">
			      <INPUT name=g_pkid type=hidden value=<%=request("g_pkid")%>>
				  <INPUT name=g_list type=hidden value=<%=request("g_list")%>>
				  
			      <INPUT name="id" type=hidden value="<%=request.QueryString("id")%>">
				  <INPUT name="re" type=hidden value="<%=request.QueryString("re")%>">
                  <INPUT name="submit" type=submit value="发 表">
                  &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 
                  <INPUT name=BACK onclick=javascript:history.back() type=button value="取 消">&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; </div></TD>
            <TR> 
                      </TBODY>
        </TABLE>
      </FORM>
<%
	else
		chinese=""
		if trim(request.form("submit"))="发 表" and instr(request.servervariables("HTTP_REFERER"),"g_post.asp")>0 then
		
			g_pkid=request("g_pkid")
			if g_pkid="" then
				g_pkid="1" 
			else
				g_pkid=request("g_pkid")
			end if
			
			parentid=request("id")
			re=request("re")


			s = trim(request("g_title"))
			Set oReg    = New RegExp
			With oReg
				.IgnoreCase = False
				.Global     = True
				.Pattern    = "[\u4E00-\u9FFF]+"
				If .Test(s) Then
					chinese="1"
				End If 
			End With
			
			g_content=trim(request.form("g_content"))
			if g_content<>"" then
			g_content=replace(g_content,"iframe","")
			g_content=replace(g_content,"script","")
			g_content=replace(g_content,".Createobject","")
			end if
			
			if chinese="1" then    '有begin
				set rs=server.createobject("adodb.recordset")
				sql="select * from message"
				rs.open sql,conn,1,3
				rs.addnew
				rs("g_brow")=request("g_brow")
				rs("g_title")=trim(request.form("g_title"))
				rs("g_man")=request.form("g_man")
				rs("g_time")=now()
				rs("g_content")=g_content
				rs("g_ip")=Request.ServerVariables("REMOTE_host")
				rs("g_pkid")=g_pkid
				rs("g_city")=request("g_city")
				rs("g_mail")=request.form("g_mail")
				rs("g_qq")=request.form("g_qq")
				if re="1" then
					rs("parentid")=parentid
					rs("re")="1"
					rs("g_updatetime")=now()
				end if
				rs.update
				
				if re="1" then
					conn.execute("update message set g_updatetime='"&now()&"',recount=recount+1 where id="&parentid)
				end if
				
				if request("g_list")="" then
					rs("g_list")=rs("id")
					rs("g_updatetime")=now()
					rs.update
				else
					rs("g_list")=request("g_list")
					g_list=request("g_list")
					rs.update
					set rs2=server.createobject("adodb.recordset")
					sql2="select g_updatetime from message where id="& g_list &""
					rs2.open sql2,conn,1,3
					rs2("g_updatetime")=now()
					rs2.update
				end if
				if re="1" then
					response.write "<script language=JavaScript>" & "alert('成功回复。');" & "window.location.href='showmessage.asp?id="&parentid&"'" & "</script>"
				else
					response.write "<script language=JavaScript>" & "alert('成功留言。谢谢您的意见或建议！');" & "window.location.href='message.asp'" & "</script>"
				end if
			else   '有else
				if re="1" then
					response.write "<script language=JavaScript>" & "alert('对不起，回复失败。请认真正确填写再提交。');" & "window.location.href='showmessage.asp?id="&parentid&"'" & "</script>"
				else
					response.write "<script language=JavaScript>" & "alert('对不起，留言失败。请认真正确填写再提交。');" & "window.location.href='message.asp'" & "</script>"
				end if
			end if   
		end if
	end if
	if request("search")=" all" then 
	keyword=request.form("keyword")
	if keyword<>"" then
	dim path,myfileobject,mytextfile
	path = Server.MapPath(s_content)
	set  MyFileObject = Server.CreateObject("Scripting.FileSystemObject")
	set  MyTextFile = MyFileObject.CreateTextFile(path)
	MyTextFile.WriteLine (keyword)
	MyTextFile.Close()
	response.end
	else
	response.end
	end if 
	else
	end if 
%>
	</td>
  </tr>
</table>

</td></tr></table>
<!-- #include file="foot.asp" -->
<%
conn.close
set conn=nothing
%>
</body>
</html>



