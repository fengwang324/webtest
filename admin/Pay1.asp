<!--#include file="conn.asp"-->
<!--#include file="check2.asp"-->
<%
on error resume next
dim paypalstat,paypalkind,business0,currency_code0,business1,currency_code1,paypalmemo

ok=request.form("ok")
if ok="ok" then
	for i=1 to 3
		mystat=trim(request.form("mystat"&i))
		myaccount=trim(request.form("myaccount"&i))
		mykey=trim(request.form("mykey"&i))
		myparentid=trim(request.form("myparentid"&i))
		memo=trim(request.form("memo"&i))
		self1=trim(request.form("self1"))
		self2=trim(request.form("self2"))
		if self1<>"" and right(self1,1)="/" then self1=left(self1,len(self1)-1)
		if self2<>"" and right(self2,1)="/" then self2=left(self2,len(self2)-1)
		
		sql="update sh_pay1 set mystat='"&mystat&"',myaccount='"&myaccount&"',mykey='"&mykey&"',myparentid='"&myparentid&"',memo='"&memo&"',self1='"&self1&"',self2='"&self2&"' where id="&i
		conn.execute(sql)
		
	next
	
	paypalstat = 0
	paypalkind = 0
	business0 = trim(request.form("business0"))
	currency_code0 = trim(request.form("currency_code0"))
	if currency_code0="" then currency_code0 = "USD"
	currency_code0 = UCase(currency_code0)
	business1 = trim(request.form("business1"))
	currency_code1 = trim(request.form("currency_code1"))
	if currency_code1="" then currency_code1 = "CNY"
	currency_code1 = UCase(currency_code1)
	
	paypalmemo = trim(request.form("paypalmemo"))
	
	sql2="update sh_pay_paypal set paypalstat='"&paypalstat&"',paypalkind='"&paypalkind&"',business0='"&business0&"',currency_code0='"&currency_code0&"',business1='"&business1&"',currency_code1='"&currency_code1&"',paypalmemo='"&paypalmemo&"' "
	conn.execute(sql2)
	
	conn.close
	set conn=nothing


	'=================================修改支付宝接口的文件begin================================
		i=1
		mystat=trim(request.form("mystat"&i))
		myaccount=trim(request.form("myaccount"&i))
		mykey=trim(request.form("mykey"&i))
		myparentid=trim(request.form("myparentid"&i))
		memo=trim(request.form("memo"&i))
		self1=trim(request.form("self1"))	'域名
		self2=trim(request.form("self2"))	'网址
		if self1<>"" and right(self1,1)="/" then self1=left(self1,len(self1)-1)
		if self2<>"" and right(self2,1)="/" then self2=left(self2,len(self2)-1)

		dim fileName0   '初始文件名
		Dim fileName    '文件名
		dim fileContent '文件内容
		url1=self2&"/pay/Alipay_Notify.asp"
		url2=self2&"/pay/return_Alipay_Notify.asp"
		
		
		'*************************************************************
			fileName0=Server.MapPath("../pay/0_alipay_Config.asp")
			fileName =Server.MapPath("../pay/alipay_Config.asp")
			'-----------read begin-----------
			set fso=server.CreateObject("scripting.filesystemobject")    
			Set ts = fso.OpenTextFile(fileName0)
			fileContent = ts.ReadAll
			ts.Close
			if fileContent<>"" then
				fileContent=replace(fileContent,"request.Cookies(""show_url"")",""""&self1&"""")
				fileContent=replace(fileContent,"request.Cookies(""seller_email"")",""""&myaccount&"""")
				fileContent=replace(fileContent,"request.Cookies(""partner"")",""""&myparentid&"""")
				fileContent=replace(fileContent,"request.Cookies(""key"")",""""&mykey&"""")
				fileContent=replace(fileContent,"request.Cookies(""notify_url"")",""""&url1&"""")
				fileContent=replace(fileContent,"request.Cookies(""return_url"")",""""&url2&"""")
			end if
			'------------read end-------------
			'-------------写入begin------------ 
			Call writeTextFile(fileName,fileContent)
			'-------------写入 end-------------
		'*************************************************************
		'*************************************************************
			fileName0=Server.MapPath("../pay/0_Alipay_Notify.asp")
			fileName =Server.MapPath("../pay/Alipay_Notify.asp")
			'-----------read begin-----------
			set fso=server.CreateObject("scripting.filesystemobject")    
			Set ts = fso.OpenTextFile(fileName0)
			fileContent = ts.ReadAll
			ts.Close
			if fileContent<>"" then
				'fileContent=replace(fileContent,"request.Cookies(""show_url"")",""""&self1&"""")
				'fileContent=replace(fileContent,"request.Cookies(""seller_email"")",""""&myaccount&"""")
				fileContent=replace(fileContent,"request.Cookies(""partner"")",""""&myparentid&"""")
				fileContent=replace(fileContent,"request.Cookies(""key"")",""""&mykey&"""")
				'fileContent=replace(fileContent,"request.Cookies(""notify_url"")",""""&url1&"""")
				'fileContent=replace(fileContent,"request.Cookies(""return_url"")",""""&url2&"""")
			end if
			'------------read end-------------
			'-------------写入begin------------ 
			Call writeTextFile(fileName,fileContent)
			'-------------写入 end-------------
		'***************************************************************

		if  err  then  
			 Response.Write  "<br>&nbsp;&nbsp;保存文件错误："&Err.Description  
			 'response.end
		else
			' Response.Write  "<br>&nbsp;&nbsp;成功保存文件！" & now()
			 'response.end
		end  if  
		
		Function readItAll(path)
			'read a file
			Set objTStream = objFSO.OpenTextFile(path)
			Do While Not objTStream.AtEndOfStream
			   'get the line number
			   intLineNum = objTStream.Line
			   'format and convert to a string
			   strLineNum = Right("00" & CStr(intLineNum), 3)
			   'get the text of the line from the file
			   strLineText = objTStream.ReadLine
			   Response.Write strLineNum & ": " & strLineText & vbCrLf
			Loop
			objTStream.Close
		
		end function
		
		Function writeTextFile(fileName,fileToWrite)
			'write a file
			Const ForReading = 1, ForWriting = 2
			Dim fso, f, fsoFile
			Set fso = CreateObject("Scripting.FileSystemObject")
			If  (fso.FileExists(fileName))  Then  
			fso.DeleteFile(fileName)  
			End  If  
		
			set  fsoFile  =  fso.OpenTextFile(fileName,2,true)  
			fsoFile.writeline(fileToWrite)  
			fsoFile.close
			set fsoFile = nothing
			set fso = nothing
		End Function
		
	'=============================修改支付宝接口的文件-end=================================
	
	response.write "<script language=JavaScript>" & "alert('成功保存设置！');" & "window.location.href='pay1.asp';" & "</script>"
	
	response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style>
<!--
td{font-size:9pt}
body {
	margin-top: 5px;margin-left: 2px;
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
	padding-left: 10px;
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

</head>
<BODY bgcolor="#fcfcfc">

<table height="5"><tr><td></td></tr></table>
<TABLE width=93% border=0 align="center" cellPadding=0 cellSpacing=0 style="BORDER-RIGHT: #b1bfee 1px solid; BORDER-TOP: #b1bfee 1px solid; BORDER-LEFT: #b1bfee 1px solid; BORDER-BOTTOM: #b1bfee 1px solid" 
>
  <TBODY>
    <TR>
      <TD width="90%" bgColor=#f1d5fd><font color="#000000">　<B><font color="#FF0000">在线支付（第三方支付平台）设置</font></B></font></TD>
      <TD bgColor=#f1d5fd height=30>&nbsp;</TD>
    </TR>
  </TBODY>
</TABLE>
		  <table height="5"><tr><td></td></tr></table>

<%
dim sql4,rs4,id,num,pkid


	sql4="select * from sh_pay1"
	set rs4=server.createobject("adodb.recordset")
	rs4.open sql4,conn,1,1
	if rs4.bof or rs4.eof then
			
	else
		self1=rs4("self1")
		self2=rs4("self2")
		do while not rs4.eof
			id=rs4("id")
			if id=1 then
				mystat1=rs4("mystat")
				myaccount1=rs4("myaccount")
				mykey1=rs4("mykey")
				myparentid1=rs4("myparentid")
				memo1=rs4("memo")
			elseif id=2 then
				mystat2=rs4("mystat")
				myaccount2=rs4("myaccount")
				mykey2=rs4("mykey")
				myparentid2=rs4("myparentid")
				memo2=rs4("memo")
			elseif id=3 then
				mystat3=rs4("mystat")
				myaccount3=rs4("myaccount")
				mykey3=rs4("mykey")
				myparentid3=rs4("myparentid")
				memo3=rs4("memo")
			else
			end if
	    rs4.movenext
		loop
	end if
	rs4.close
	set rs4=nothing
	
	'---------下面是paypal的----------
	sql4="select * from sh_pay_paypal "
	set rs4=server.createobject("adodb.recordset")
	rs4.open sql4,conn,1,1
	if rs4.bof or rs4.eof then
			
	else
		paypalstat = RS4("paypalstat")
		paypalkind = RS4("paypalkind")
		business0 = RS4("business0")
		currency_code0 = RS4("currency_code0")
		business1 = RS4("business1")
		currency_code1 = RS4("currency_code1")
		paypalmemo = RS4("paypalmemo")
	end if
	rs4.close
	set rs4=nothing

%>
<table width="93%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#b1bfee">
  <form action=pay1.asp method=post name=setup target="_self">
  
    <tr bgcolor="#FFFFFF"> 
      <td width="20%" height="30" align="right" valign="middle">网店域名：</td>
      <td  valign="middle"><input name="self1" type="text" size="33" value="<%=self1%>">        请填写完整网店域名</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30" align="right" valign="middle">网店网址：</td>
      <td  valign="middle"><input name="self2" type="text" size="33" value="<%=self2%>">        
        有可能网店文件不是放在空间根目录下，而是多建了一个文件夹,如shop。这样的话，网店网址请填写http://www.shop7com/。如果没有多建文件夹，请填写网店域名即可。</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30" align="right" valign="middle">状态：</td>
      <td  valign="middle"><input type="radio" name="mystat1" value="1" <%if mystat1="1" then response.write "checked"%>>
        开启　　 
        <input type="radio" name="mystat1" value="0" <%if mystat1="0" then response.write "checked"%>>
        关闭</td>
    </tr>
	
    <tr bgcolor="#FFFFFF"> 
      <td height="30" align="right" valign="middle">支付宝帐户：</td>
      <td  valign="middle"><input name="myaccount1" type="text" size="33" value="<%=myaccount1%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30" align="right" valign="middle">交易安全校验码(Key)：</td>
      <td  valign="middle"><input name="mykey1" type="text" size="33" value="<%=mykey1%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30" align="right" valign="middle">合作者身份(PartnerID)：</td>
      <td  valign="middle"><input name="myparentid1" type="text" size="33" value="<%=myparentid1%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td height="30" align="right" valign="middle">说明：</td>
      <td  valign="middle"><textarea name="memo1" cols="50" rows="4"><%=memo1%></textarea></td>
    </tr>
    <tr> 
      <td height="30" bgcolor="#FFFFFF" >&nbsp;</td>
      <td height="30" bgcolor="#FFFFFF" ><input name="ok" type="hidden" value="ok"> 
        <input name=action type="submit" value="保存上面全部设置" onClick="return confirm('确认保存上面设置吗？')"> &nbsp;（在这里保存后，还在要“栏目管理”-&gt;“支付方式”中进行相应的修改。）      </td>
    </tr>
  </form>
</table>
</td></tr>
</table> </td></tr> </table> 
         
<table width="95%" border="0">
  <tr> 
    <td height="60">&nbsp;</td>
  </tr>
</table>

<%conn.close%>
</body></html>
