<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->

<%
action=request.form("action")
id=request("id")
rememo=request.form("rememo")
if action="re" and id<>"" then
conn.execute("update z_pijia set rememo='"&rememo&"' where id="&id&" ")
response.write "<script language=JavaScript>" & "alert('保存成功！');" &"window.opener.location.reload();window.close();" & "</script>"
call connclose
response.end
end if
dim rememo
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title> 回复评论</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
td{font-size:9pt}
body {
	margin-top: 5px;
}
.style3 {color: #A84E22; font-weight: bold; }
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
.input3 {
      	BORDER-RIGHT: #666666 1px solid;BORDER-TOP: #666666 0px solid; height:23px;FONT-SIZE: 10pt; BORDER-LEFT: #666666 0px solid; BORDER-BOTTOM: #666666 1px solid; FONT-FAMILY: "宋体"; BACKGROUND-COLOR: #bbbbbb;padding-top: 4px;
}
-->
</style>
</head>

<body bgcolor="#fcfcfc">
<table width="550" align="center" border="0" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
<tr>
                      <td height="230" bgcolor="#FFFFFF"> 
					<table width="100%" border="0" cellspacing="0" cellpadding="5">
					<tr>
							<td height="23" bgcolor="#E3E3E3"><strong>商品评论：</strong></td>
					</tr>
					</table>
					<%
						sql="select * from z_pijia where id="&request.querystring("id")&" order by id desc"
						set rsp=server.CreateObject("adodb.recordset")
						rsp.open sql,conn,1,1
						if rsp.bof or rsp.eof then
							response.write "<div align=left>暂时没有此产品的评价记录！</div>"
						else
								set yourname=rsp("yourname")
								set tel=rsp("tel")
								set memo=rsp("memo")
								rememo=rsp("rememo")
								set addtime=rsp("addtime")
								if memo<>"" then
								memo = replace(memo,vbcrlf,"<br>")
								memo = replace(memo," ","&nbsp;")
								end if
					%>
				  
                        <table width="100%" border="0" cellspacing="2" cellpadding="0" class="border" >
                          <tr> 
                            <td width="17%" height="23" bgcolor="#f1f1f1"> 
								<div align="right"> 姓名：</div></td>
                            <td width="26%"><%=yourname%>&nbsp; </td>
                            <td width="17%" bgcolor="#f1f1f1"> 
                              <div align="right">联系方式：</div></td>
                            <td width="40%"><%=tel%></td>
                          </tr>
                          <tr> 
                            <td height="23" bgcolor="#f1f1f1"> 
								<div align="right">评论内容：</div></td>
                            <td colspan="3" style="WIDTH: 450; WORD-WRAP: break-word"><%=memo%>&nbsp;[<%=addtime%>]</td>
                          </tr>
						</table>
						<table width="100%" border="0" cellspacing="0">
                          <tr> 
                            <td height="5"></td>
                          </tr>
                        </table>
          <%
				end if
				rsp.close
				set rsp=nothing
				session("pijia")="1"
		  %>  

                  <table width="100%" border="0" cellspacing="0" cellpadding="5">
						<tr> 
                            <td height="23" bgcolor="#E3E3E3"><strong>回复评论：</strong></td>
                    </tr>
                  </table>
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <form name="Form2" method="post" action="pijiare.asp">
                            <tr> 
							<td><div align="right">回复内容：</div></td>
                              <td><textarea name="rememo" cols="50" rows="5"><%=rememo%></textarea></td> 
                            </tr>
							<input type="hidden" name="action" value="re">
							<input type="hidden" name="id" value="<%=id%>">
                            <tr> 
							<td height="30">&nbsp;</td>
                              <td><input type="submit" name="Submit2" value="提交回复" class='input3'></td>
                            </tr>
                          </form>
                        </table>
                </td>
				</tr>
			</table>
<br>

</body>
</html>
<%
call connclose
%>


