<%
dim CheckForm,TreeStr

Rem<<====================================  语言版本信息公共函数开始 ===============================
Rem  程序作用  -- 获取语言版本信息
Function getLanID()
	getLanID = Request("LanID")
	if Trim(getLanID)<>"1" and Trim(getLanID)<>"2" and Trim(getLanID)<>"3" then
		getLanID =Session("LanID")
		if getLanID="" then getLanID=1
	end if
end Function

Function LanIDSelect(ID)
For N=0 To Ubound(Split(Setting(162),vbcrlf))
	LanIDSelect = LanIDSelect & "<option value='"&n+1&"' "	
	if chkclng(ID)=n+1 then LanIDSelect = LanIDSelect & " selected "	
	LanIDSelect = LanIDSelect & ">"&trim(Lcase(Split(Setting(162),vbcrlf)(n)))&"</option>"
	next
End Function

Rem  程序作用  -- 获取语言版本名称信息
Function getLanIDTitle(LanID)
	if Trim(LanID)="1" then
		getLanIDTitle = "中文版"
	elseif Trim(LanID)="2" then
		getLanIDTitle = "英文版"
	else
	    getLanIDTitle = "末知版本"
	end if
End Function

Rem  ===================================  语言版本信息公共函数结束 ============================== >>

Rem<<======================================  类别信息公共函数开始 ================================
Rem  程序作用  -- 获取网站栏目信息
Rem  ChannelID -- 栏目ID
Rem  CID       -- 要获取的字段名称
Rem  LanID     -- 语言版本
Function getChannelInfo(ChannelID,CID)
    if Trim(CID)="" or Trim(ChannelID)="" or isNumeric(ChannelID)=False then
	    getChannelInfo = "Param Error"
		Exit Function
	end if
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select TOP 1 ["&CID&"] From [ChannelInfo] Where [ID]="&ChannelID),conn,1,1
	if objRs.bof or objRs.eof then
	    getChannelInfo = ""
	else
	    getChannelInfo = objRs(0)
	end if
	objRs.Close
	Set objRs = Nothing	
End Function
Rem  程序作用  -- 获取网站子类别信息
Rem  CateID    -- 栏目ID
Rem  CID       -- 要获取的字段名称
Function getInfo(ParentID,CID)
    if Trim(CID)="" or isNumeric(ParentID)=False then
	    getInfo = "Param Error"
		Exit Function
	end if
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select ["&CID&"] From [ChannelInfo] Where [ParentID]="&ParentID&" "),conn,1,1
	if objRs.bof or objRs.eof then
	    getInfo = ParentID
	else
		getInfo=ParentID
		do while not objrs.eof
	    getInfo =getInfo&","&objRs(0)
		getInfo =getInfo&","&getInfo(objRs(0),cid)
		objrs.movenext
		loop
	end if
	objRs.Close
	Set objRs = Nothing	
End Function

'字符串查找		
function topmuen(Domain,str)
dim GetAutoDomain
topmuen=true
	if instr(Lcase(Domain),Lcase(str))>0 then
	topmuen=true
	else
	topmuen=False
	end if
end function
Function getCatePaths(CateTree)
    ArrCate = Split(CateTree,",")	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	For M = 0 To UBound(ArrCate)
		if ArrCate(M)<>0 then
			objRs.Open("Select TOP 1 [Title] From [CateInfo] Where [ID]="&ArrCate(M)),conn,1,1
			if not (objRs.bof or objRs.eof) then
				if getCatePaths = "" then 
					getCatePaths = " > " & "<a href=""ProductList.asp?Cate="&ArrCate(M)&""">"&objRs(0)&"</a>"
				else
					getCatePaths = getCatePaths & " > " & "<a href=""ProductList.asp?Cate="&ArrCate(M)&""">"&objRs(0)&"</a>"
				end if
				objRs.Close
			end if
		end if
	Next
	Set objRs = Nothing	
End Function


Rem  程序作用  -- 当前类别下拉选择列表
Rem  CateID -- 当前类别ID
Function getCateSelect2( CateID )	
	Set objRs2 = Server.CreateObject("Adodb.Recordset")
	objRs2.Open("Select [ID],[Title] From [ChannelInfo] Where [id]="&CateID),conn,1,1
	while not objRs2.eof
		if Trim(objRs2(0)) = Trim(CateID) then 
			Response.Write"<option value='"&objRs2(0)&"' selected>"&LevelLine&""&objRs2(1)&"</option>"
		else
			Response.Write("<option value='"&objRs2(0)&"'>"&LevelLine&""&objRs2(1)&"</option>")
		end if		
		getCateSelect objRs2(0), 1, CateID 
		objRs2.MoveNext
	wend
	objRs2.Close
	Set objRs2 = Nothing	
End Function

Function getPChannelSelect( CateID )	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select [ID],[Title] From [CateInfo] Where [ParentID]=0"),conn,1,1
	while not objRs.eof
		PChannelID = Int(objRs(0)) + 10000
		if Trim(PChannelID) = Trim(CateID) then 
			Response.Write"<option value='"&PChannelID&"' selected>"&objRs(1)&"</option>"
		else
			Response.Write("<option value='"&PChannelID&"'>"&objRs(1)&"</option>")
		end if 
		objRs.MoveNext
	wend
	objRs.Close
	Set objRs = Nothing	
End Function

Function getCateID( CateID )	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select [ID],[Title] From [CateInfo] Where [channelid]="&CateID&""),conn,1,1
	if not objRs.eof then
		getCateID=objRs(0)
	end if
	objRs.Close
	Set objRs = Nothing	
End Function

Rem  程序作用 -- 获取类别ID树
Rem  ParentID -- 父类别ID
Function getCateTree( ParentID )	
	if Trim(ParentID) = "0" then
	    getCateTree = "0"
	else
	    getCateTree = getCateInfo(ParentID,"CateTree") & "," & ParentID
	end if
End Function
Function getCateBottomList(ParentID,CateList)	 
	Set objRss = Server.CreateObject("Adodb.Recordset")
	objRss.Open(" Select [ID],(Select Count(*) From [CateInfo] Where [ParentID]=P.[ID]) AS [SON] "&_
	            " From [CateInfo] P Where [ParentID]="&ParentID ),conn,1,1
	if objRss.bof or objRss.eof then
		if Trim(CateList) = "" then
			CateList = ParentID
		else
			CateList = CateList & "," & ParentID
		end if
	else
		while not objRss.eof
			if objRss(1) = 0 then
				if Trim(CateList) = "" then
					CateList = objRss(0)
				else
					CateList = CateList & "," & objRss(0)
				end if
			else
				getCateBottomList objRss(0), CateList
			end if
			objRss.Movenext		
		wend
	end if
	objRss.Close
	Set objRss = Nothing
	getCateBottomList = CateList
End Function
  
Rem  =====================================  类别信息公共函数结束 ================================= >>

Rem<<=====================================  栏目信息公共函数开始 =================================

Rem  程序作用  -- 获取网站栏目信息
Rem  ChannelID -- 栏目ID
Rem  CID       -- 要获取的字段名称
Rem  LanID     -- 语言版本
Function getChannelInfo(ChannelID,CID)
    if Trim(CID)="" or Trim(ChannelID)="" or isNumeric(ChannelID)=False then
	    getChannelInfo = "Param Error"
		Exit Function
	end if
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select TOP 1 ["&CID&"] From [ChannelInfo] Where [ID]="&ChannelID),conn,1,1
	if objRs.bof or objRs.eof then
	    getChannelInfo = ""
	else
	    getChannelInfo = objRs(0)
	end if
	objRs.Close
	Set objRs = Nothing	
End Function


Rem  程序作用    -- 网站栏目路径信息
Rem  ChannelTree -- 栏目ID树
Function getChannelPath(ChannelTree)
if lanID="1" then
		title_name="title"
	else
		title_name="titleEN"
	end if 
    ArrChannel = Split(ChannelTree,",")	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	For M = 0 To UBound(ArrChannel)			
		if ArrChannel(M)=0 then
		    getChannelPath = ""
		else
			objRs.Open("Select TOP 1 ["&title_name&"] From [ChannelInfo] "&_
			" Where [ID]="&ArrChannel(M)),conn,1,1
			if not (objRs.bof or objRs.eof) then
				if getChannelPath = "" then 
					getChannelPath = objRs(0)
				else
					getChannelPath = getChannelPath & " > " & objRs(0)
				end if
				objRs.Close
			end if
		end if
	Next
	Set objRs = Nothing	
End Function

Rem  程序作用    -- 前台网站栏目路径信息
Rem  ChannelTree -- 栏目ID树
Function getChannelDH(ChannelTree)
    ArrChannel = Split(ChannelTree,",")	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	For M = 0 To UBound(ArrChannel)
		if ArrChannel(M)=0 then
		    getChannelDH = ""
		else
			objRs.Open("Select TOP 1 [Title] From [ChannelInfo] "&_
			" Where [ID]="&ArrChannel(M)),conn,1,1
			if not (objRs.bof or objRs.eof) then
					getChannelDH = getChannelDH & " > " &objRs(0)&""
				objRs.Close
			end if
		end if
	Next
	Set objRs = Nothing	
	getChannelDH=Replace(getChannelDH,"产品展示 >","")
End Function

Rem  程序作用  -- 网站栏目路径信息
Rem  ChannelID -- 栏目ID
Function ChannelPath(ChannelID)
    ChannelTree = getChannelInfo(ChannelID,"ChannelTree")
	ArrChannel = Split(ChannelTree,",")	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	For M = 0 To UBound(ArrChannel)			
		if ArrChannel(M)=0 then
		    ChannelPath = "<a href='index.asp'>首页</a>"
		else
			objRs.Open("Select TOP 1 [ID],[Title],[ChannelModel] From [ChannelInfo] Where [ID]="&ArrChannel(M)),conn,1,1
			if not (objRs.bof or objRs.eof) then		
					if Trim(ChannelModel) = "1" then
						ChannelPath = ChannelPath & " > " & "<a href='NewsList.asp?ChannelID="&objRs(0)&"'>"& objRs(1) &"</a>"
					elseif Trim(ChannelModel) = "3" then
						ChannelPath = ChannelPath & " > " & "<a href='VideoList.asp?ChannelID="&objRs(0)&"'>"& objRs(1) &"</a>"
					else
						ChannelPath = ChannelPath & " > " & "<a href='?ChannelID="&objRs(0)&"'>"& objRs(1) &"</a>"
					end if
				objRs.Close
			end if
		end if
	Next
	Set objRs = Nothing
	CTitle = getChannelInfo(ChannelID,"Title")
	CChannelModel = getChannelInfo(ChannelID,"ChannelModel")
	if Trim(CChannelModel) = "1" then
		ChannelPath = ChannelPath & " > " & "<a href='NewsList.asp?ChannelID="&ChannelID&"'>"& CTitle &"</a>"
	elseif Trim(CChannelModel) = "3" then
		ChannelPath = ChannelPath & " > " & "<a href='VideoList.asp?ChannelID="&ChannelID&"'>"& CTitle &"</a>"
	else
		ChannelPath = ChannelPath & " > " & "<a href='?ChannelID="&ChannelID&"'>"& CTitle &"</a>"
	end if
End Function


Rem  程序作用  -- 网站栏目下拉选择列表
Rem  ParentID  -- 所属栏目ID
Rem  LevelID   -- 栏目层级ID
Rem  ChannelID -- 当前栏目ID
Function getChannelSelect( ParentID, LevelID, ChannelID )
	if lanID="1" then
		title_name="title"
	else
		title_name="titleEN"
	end if	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select [ID],["&title_name&"] From [ChannelInfo] Where [ParentID]="&ParentID),conn,1,1
	while not objRs.eof
		LevelLine = ""
		For L = 1 To LevelID 
		    LevelLine = LevelLine & "—"
		Next 
		if Trim(objRs(0)) = Trim(ChannelID) then 
			Response.Write"<option value='"&objRs(0)&"' selected>"&LevelLine&" "&objRs(1)&"</option>"
		else
			Response.Write("<option value='"&objRs(0)&"'>"&LevelLine&" "&objRs(1)&"</option>")
		end if		
		getChannelSelect objRs(0), Int(LevelID) + 1, ChannelID 
		objRs.MoveNext
	wend
	objRs.Close
	Set objRs = Nothing	
End Function


Rem  程序作用  -- 网站栏目下拉选择列表
Rem  ParentID  -- 所属栏目ID
Rem  LevelID   -- 栏目层级ID
Rem  ChannelID -- 当前栏目ID
Function getChannel( ParentID, LevelID, ChannelID )	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select [ID],[Title] From [ChannelInfo] Where [ParentID]="&ParentID),conn,1,1
	while not objRs.eof
		LevelLine = ""
		For L = 1 To LevelID 
		    LevelLine = LevelLine & "—"
		Next 
		if Trim(objRs(0)) = Trim(ChannelID) then 
			Response.Write"<option value='"&objRs(0)&"' selected>"&LevelLine&" "&objRs(1)&"</option>"
		else
			Response.Write("<option value='"&objRs(0)&"'>"&LevelLine&" "&objRs(1)&"</option>")
		end if		
		getChannelSelect objRs(0), Int(LevelID) + 1, ChannelID 
		objRs.MoveNext
	wend
	objRs.Close
	Set objRs = Nothing	
End Function


Rem  程序作用 -- 获取栏目ID树
Rem  ParentID -- 父栏目ID
Function getChannelTree( ParentID )	
	if Trim(ParentID) = "0" then
	    getChannelTree = "0"
	else
	    getChannelTree = getChannelInfo(ParentID,"ChannelTree") & "," & ParentID
	end if
End Function

Function getModelSelect(ID)
    getModelSelect = "<input type='radio' name='ChannelModel' value='0' checked>栏目列表"
	
	getModelSelect = getModelSelect & "<input type='radio' name='ChannelModel' value='1' "
	if Trim(ID)="1" then getModelSelect = getModelSelect & " checked "
	getModelSelect = getModelSelect & ">新闻列表"
	
	getModelSelect = getModelSelect & "<input type='radio' name='ChannelModel' value='2' "
	if Trim(ID)="2" then getModelSelect = getModelSelect & " checked "
	getModelSelect = getModelSelect & ">单篇文章"
	
	getModelSelect = getModelSelect & "<input type='radio' name='ChannelModel' value='3' "
	if Trim(ID)="3" then getModelSelect = getModelSelect & " checked "
	getModelSelect = getModelSelect & ">产品新闻"
	
	getModelSelect = getModelSelect & "<input type='radio' name='ChannelModel' value='4' "
	if Trim(ID)="4" then getModelSelect = getModelSelect & " checked "
	getModelSelect = getModelSelect & ">其它连接"
End Function

Function getModelLink(ID)
    Select Case Trim(ID)
	    Case "0" 
		    getModelLink = "ChannelList.asp?ParentID="	
	    Case "1" 
		    getModelLink = "NewsList.asp?ChannelID="			
		Case "2" 
		    getModelLink = "ChannelInfo.asp?ChannelID="			
		Case "3" 
		    getModelLink = "ProductList.asp?ChannelID="			
	End Select
End Function

Function getModelName(ID)
    Select Case Trim(ID)
	    Case "0" 
		    getModelLink = "栏目列表"	
	    Case "1" 
		    getModelLink = "新闻列表"			
		Case "2" 
		    getModelLink = "单篇文章"			
		Case "3" 
		    getModelLink = "产品新闻"				
	End Select
End Function

Function RequestChannelID(ParentID)
    ChannelID = Request.QueryString("ChannelID")
	if ChannelID<>"" and isNumeric(ChannelID) then
	    RequestChannelID = ChannelID
	else
	    Set objRs = Server.CreateObject("Adodb.Recordset")
		objRs.Open("Select TOP 1 [ID] From [ChannelInfo] Where [ParentID]="&ParentID&" Order by [OIndex] ASC,[ID] ASC"),conn,1,1
		if objRs.bof or objRs.eof then
			RequestChannelID = ParentID
		else
		    RequestChannelID = objRs(0)
		end if
		objRs.Close
		Set objRs = Nothing
	end if
End Function

Rem  程序作用 -- 分页控件
Rem  Curpage  -- 当前页码
Rem  Records  -- 总记录数
Rem  Pages    -- 总页码数
Rem  FileName -- 链接字串
Sub ShowPage(Curpage,Records,Pages,FileName)
	Response.Write("<script language=""javascript"">function viewPage(ipage,url) "&_
    "{document.location.href=url+ipage;}</script><div style=""padding:5px 0px 5px 0px;"&_
	"text-align:center;"">"&_
	" 第 <b style=""color:Red"">"&Curpage&"</b> 页 / 总 <b style=""color:Red"">"&Pages&"</b> "&_
	"页 | 总 <b style=""color:Red"">"&Records&"</b> 条记录 ")
	if Trim(Curpage)="" or isNumeric(Curpage)=False then 
	    Response.Write("<script>alert('页码请输入数字！');history.go(-1);</script>")
		Response.End()
	end if
    if Pages>1 then  
        if Curpage=1 then 
            Response.Write(" <font color=""#999999"">首页</font> "&_
            " <font color=""#999999"">上页</font> "&_
            " <a href=""javascript:viewPage("&Curpage+1&",'"&FileName&"')"" style=""color:#000"">下页</a> "&_
            " <a href=""javascript:viewPage("&Pages&",'"&FileName&"')"" style=""color:#000"">尾页</a> ")
        elseif Curpage=Pages then                 
            Response.Write(" <a href=""javascript:viewPage(1,'"&FileName&"')"" style=""color:#000"">首页</a> "&_
            " <a href=""javascript:viewPage("&Curpage-1&",'"&FileName&"')"" style=""color:#000"">上页</a> "&_
            " <font color=""#999999"">下页</font> "&_
            " <font color=""#999999"">尾页</font> ")
        else 
            Response.Write(" <a href=""javascript:viewPage(1,'"&FileName&"')"" style=""color:#000"">首页</a> "&_
            " <a href=""javascript:viewPage("&Curpage-1&",'"&FileName&"')"" style=""color:#000"">上页</a> "&_
            " <a href=""javascript:viewPage("&Curpage+1&",'"&FileName&"')"" style=""color:#000"">下页</a> "&_
            " <a href=""javascript:viewPage("&Pages&",'"&FileName&"')"" style=""color:#000"">尾页</a> ")
        end if
    else            
        Response.Write(" <font color=""#999999"">首页</font> "&_
        " <font color=""#999999"">上页</font> "&_
        " <font color=""#999999"">下页</font> "&_
        " <font color=""#999999"">尾页</font> ")
    end if

    Response.Write(" 转到第 "&_ 
    "<input type=""text"" name=""GotoPage"" value="""&Curpage&""" size=""4"" maxlength=""4"" "&_
	" onBlur=""javascript:if(isNaN(this.value)){alert('请输入数字！');this.select();};"" "&_
	" style=""text-align:center;border: solid 1px RGB(200,200,200)""> 页 <input "&_
	" style=""height:20px;width:42px;"" "&_
	" type=""button"" value=""跳转"" name=""cmd_goto"" "&_
	" onClick=""javascript:viewPage(document.all.GotoPage.value,'"&FileName&"');""></div>")
      
End Sub

Sub ShowPages(Curpage,Records,Pages,LanID,FileName)
	if Trim(LanID)="1" then
	    Response.Write("<script language=""javascript"">function viewPage(ipage,url) "&_
		"{document.location.href=url+ipage;}</script><div style=""padding:5px 0px 5px 0px;"&_
		"text-align:center;"">"&_
		" 第 <b>"&Curpage&"</b> 页 / 总 <b>"&Pages&"</b> "&_
		"页 | 总 <b>"&Records&"</b> 条记录 ")
		if Trim(Curpage)="" or isNumeric(Curpage)=False then 
			Response.Write("<script>alert('页码请输入数字！');history.go(-1);</script>")
			Response.End()
		end if
		if Pages>1 then  
			if Curpage=1 then 
				Response.Write(" [ <font color=""#999999"">首页</font> ] "&_
				" <font color=""#999999"">上页</font> "&_
				" <a href=""javascript:viewPage("&Curpage+1&",'"&FileName&"')"">下页</a> "&_
				" <a href=""javascript:viewPage("&Pages&",'"&FileName&"')"">尾页</a> ")
			elseif Curpage=Pages then                 
				Response.Write(" <a href=""javascript:viewPage(1,'"&FileName&"')"">首页</a> "&_
				" <a href=""javascript:viewPage("&Curpage-1&",'"&FileName&"')"">上页</a> "&_
				" <font color=""#999999"">下页</font> "&_
				" <font color=""#999999"">尾页</font> ")
			else 
				Response.Write(" <a href=""javascript:viewPage(1,'"&FileName&"')"">首页</a> "&_
				" <a href=""javascript:viewPage("&Curpage-1&",'"&FileName&"')"">上页</a> "&_
				" <a href=""javascript:viewPage("&Curpage+1&",'"&FileName&"')"">下页</a> "&_
				" <a href=""javascript:viewPage("&Pages&",'"&FileName&"')"">尾页</a> ")
			end if
		else            
			Response.Write(" <font color=""#999999"">首页</font> "&_
			" <font color=""#999999"">上页</font> "&_
			" <font color=""#999999"">下页</font> "&_
			" <font color=""#999999"">尾页</font> ")
		end if
	
		Response.Write(" 转到第 "&_ 
		"<input type=""text"" name=""GotoPage"" value="""&Curpage&""" size=""4"" maxlength=""4"" "&_
		" onBlur=""javascript:if(isNaN(this.value)){alert('请输入数字！');this.select();};"" "&_
		" style=""text-align:center;border: solid 1px RGB(200,200,200)""> 页 <input "&_
		" style=""height:20px;width:42px;"" "&_
		" type=""button"" value=""跳转"" name=""cmd_goto"" "&_
		" onClick=""javascript:viewPage(document.all.GotoPage.value,'"&FileName&"');""></div>")
	else
		Response.Write("<script language=""javascript"">function viewPage(ipage,url) "&_
		"{document.location.href=url+ipage;}</script><div style=""padding:5px 0px 5px 0px;"&_
		"text-align:center;"">"&_
		" Current is No. <b style=""color:red"">"&Curpage&"</b> Page / Total "&_
		" <b style=""color:red"">"&Pages&"</b> "&_
		" Pages | <b style=""color:red"">"&Records&"</b> Records ")
		if Trim(Curpage)="" or isNumeric(Curpage)=False then 
			Response.Write("<script>alert('Please input a number！');history.go(-1);</script>")
			Response.End()
		end if
		if Pages>1 then  
			if Curpage=1 then 
				Response.Write(" <font color=""#999999"">First</font> "&_
				" <font color=""#999999"">Previous</font> "&_
				" <a href=""javascript:viewPage("&Curpage+1&",'"&FileName&"')"">Next</a> "&_
				" <a href=""javascript:viewPage("&Pages&",'"&FileName&"')"">Last</a> ")
			elseif Curpage=Pages then                 
				Response.Write(" <a href=""javascript:viewPage(1,'"&FileName&"')"">First</a> "&_
				" <a href=""javascript:viewPage("&Curpage-1&",'"&FileName&"')"">Previous</a> "&_
				" <font color=""#999999"">Next</font> "&_
				" <font color=""#999999"">Last</font> ")
			else 
				Response.Write(" <a href=""javascript:viewPage(1,'"&FileName&"')"">First</a> "&_
				" <a href=""javascript:viewPage("&Curpage-1&",'"&FileName&"')"">Previous</a> "&_
				" <a href=""javascript:viewPage("&Curpage+1&",'"&FileName&"')"">Next</a> "&_
				" <a href=""javascript:viewPage("&Pages&",'"&FileName&"')"">Last</a> ")
			end if
		else            
			Response.Write(" <font color=""#999999"">First</font> "&_
			" <font color=""#999999"">Previous</font> "&_
			" <font color=""#999999"">Next</font> "&_
			" <font color=""#999999"">Last</font> ")
		end if
	
		Response.Write(" Goto No. "&_ 
		"<input type=""text"" name=""GotoPage"" value="""&Curpage&""" size=""4"" maxlength=""4"" "&_
		" onBlur=""javascript:if(isNaN(this.value))"&_
		"{alert('Please input a Number！');this.select();};"" "&_
		" style=""text-align:center;border: solid 1px RGB(200,200,200)""> Page <input "&_
		" style=""height:20px;"" "&_
		" type=""button"" value=""Go"" name=""cmd_goto"" "&_
		" onClick=""javascript:viewPage(document.all.GotoPage.value,'"&FileName&"');""></div>")
    end if
End Sub
'**************************************************
	'函数名：GetIP
	'作  用：取得正确的IP
	'返回值：IP字符串
	'**************************************************
	Function GetIP() 
		Dim strIPAddr 
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
			strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
		Else 
			strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		End If 
		getIP = Checkstr(Trim(Mid(strIPAddr, 1, 30)))
	End Function
	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function
	Public Function S(Str)
	 S = Replace(Replace(Request(Str), "'", ""), """", "")
	End Function
	'取得Request.Querystring 或 Request.Form 的值
	Public Function G(Str)
	 G = Replace(Replace(Request(Str), "'", ""), """", "")
	End Function
		'读Cookies值
	Public Function C(Str)
	 C=DelSql(Request.Cookies(SiteSn)(Str))
	End Function
	Function DelSql(Str)
		Dim SplitSqlStr,SplitSqlArr,I
		SplitSqlStr="dbcc|alter|drop|*|and |exec|or |insert|select|delete|update|count |master|truncate|declare|char|mid|chr|set |where|xp_cmdshell"
		SplitSqlArr = Split(SplitSqlStr,"|")
		For I=LBound(SplitSqlArr) To Ubound(SplitSqlArr)
			If Instr(LCase(Str),SplitSqlArr(I))>0 Then
				Die "<script>alert('系统警告！\n\n1、您提交的数据有恶意字符" & SplitSqlArr(I) &";\n2、您的数据已经被记录;\n3、您的IP："&GetIP&";\n4、操作日期："&Now&";\n		Powered By Kesion.Com!');window.close();</script>"
			End if
		Next
		DelSql = Str
    End Function
	Sub ShowError(Errmsg)
	    With Response
		.Write ("<br><br><div align=""center"">")
		.Write ("  <center>")
		.Write ("  <table border=""0"" cellpadding='2' cellspacing='1' width=""75%"" style=""MARGIN-TOP: 3px"" class='table'>")
		.Write ("	 <tr class=hback>")
		.Write ("			  <td width=""100%"" height=""30"" align=""center"">")
		.Write ("				<b> " & Errmsg & "&nbsp; </b>")
		.Write ("				</b>")
		.Write ("			  </td>")
		.Write ("	 </tr>")
		.Write ("	 <tr  class=hback>")
		.Write ("			  <td width=""100%"" height=""30"" align=""center"">")
		.Write ("				<p><b><a href=""javascript:history.go(-1)"">...::: 点 此 返 回 ")
		.Write ("				:::...</a></b>")
		.Write ("			  </td>")
		.Write ("			</tr>")		
		.Write ("  </table>")
		.Write ("  </center>")
		.Write ("</div>")
		.End
		End With
    end sub
	'提示成功。并返回
	Sub AlertHintScript(SuccessStr)
	  With Response
	  .Write "<script language=JavaScript>" & vbCrLf
	  .Write "alert('" & SuccessStr & "');"
	  .Write "location.replace('" & Request.ServerVariables("HTTP_REFERER") & "')" & vbCrLf
	  .Write "</script>" & vbCrLf
	  .end
	  end with
	End Sub
	'**************************************************
	'函数名：Alert
	'作  用：弹出成功提示。
	'参  数：SuccessStr  ----成功提示信息
	'        Url   ------成功提示按下"确定"转向链接
	'返回值：无
	'**************************************************
	Public Function Alert(SuccessStr, Url)
	 If Url <> "" Then
	  Response.Write ("<script language=""Javascript""> alert('" & SuccessStr & "');location.href='" & Url & "';</script>")
	 Else
	  Response.Write ("<script language=""Javascript""> alert('" & SuccessStr & "');</script>")
	 End If
	End Function
	Function ChkClng(ByVal str)
	    'on error resume next
		If IsNumeric(str) Then
			ChkClng = CLng(str)
		Else
			ChkClng = 0
		End If
		If Err Then ChkClng=0
	End Function
	'*************************************************************************
	'函数名：gotTopic
	'作  用：截字符串，汉字一个算两个字符，英文算一个字符
	'参  数：str   ----原字符串
	'       strlen ----截取长度
	'返回值：截取后的字符串
	'*************************************************************************
    Public Function GotTopic(ByVal Str, ByVal lennum)
		If lennum=0 Then GotTopic=Str:Exit Function
		Dim oRegExp, p_num, i, x
		Set oRegExp = new RegExp
		oRegExp.IgnoreCase = True
		oRegExp.Global = True
		oRegExp.Pattern = "[\uff00-\uffff\u4e00-\u9fa5\ufe10-\ufe1f\ufe30-\ufe4f\u1100-\u11ff\u2600-\u26ff\u2700-\u27bf\u2800-\u28ff\u3300-\u33ff\u3200-\u32ff\ua490-\ua4cf\ua000-\ua48f\u3130-\u318f\uac00-\ud7af\u31f0-\u31ff\u30a0-\u30ff\u3040-\u309f\u31a0-\u31bf\u3100-\u312F\u2FF0-\u2FFF\u2F00-\u2FDF\u31c0-\u31ef\u3000-\u303f\u2e80-\u2eff\uff00-\uffef]"
		p_num = 0
		x = 0
		Do While Not p_num > lennum - 2
		x = x + 1
		p_num = int(p_num) + 1
		
		If oRegExp.Test(str) Then 
		p_num = p_num + 1
		End If
		GotTopic = left(trim(str),x)
		Loop
		Set oRegExp=Nothing

	End Function
	'返回付款方式名称,参数TypeID,0名称 1折扣率
	Function ReturnPayment(ID,TypeID)
	  If Application(SiteSn &"Payment_" & ID&TypeID)="" Then
         Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select TypeName,Discount From PaymentType Where TypeID=" & ID,conn,1,1
		 If Not RS.Eof Then
		     If TypeID=0 Then
		  	  ReturnPayment=rs(0)
			  If RS(1)<100 Then ReturnPayment=ReturnPayment & "&nbsp;&nbsp;<font color=red>折扣率:" & RS(1) & "%"
			 Else
			  ReturnPayment=rs(1)
			 End if
		End iF 
		Application(SiteSn &"Payment_" & ID&TypeID)=ReturnPayment
	  Else
	    ReturnPayment=Application(SiteSn &"Payment_" & ID&TypeID)
	  End If
	End Function
		'返回收货方式名称,参数TypeID,0名称 1费用
	Function ReturnDelivery(ID,TypeID)
	  If Application(SiteSn &"Delivery_" & ID&TypeID)="" Then
         Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select TypeName,fee From Delivery Where TypeID=" & ID,conn,1,1
		 If Not RS.Eof Then
		     If TypeID=0 Then
		  	  ReturnDelivery=rs(0)
			  If RS(1)=0 Then ReturnDelivery=ReturnDelivery & "&nbsp;<font color=blue>免费</font>" Else ReturnDelivery=ReturnDelivery & "&nbsp;<font color=red>加收 " & RS(1) & "元"
			 Else
			  ReturnDelivery=rs(1)
			 End iF
		End iF 
		Application(SiteSn &"Delivery_" & ID&TypeID)=ReturnDelivery
	  Else
	    ReturnDelivery=Application(SiteSn &"Delivery_" & ID&TypeID)
	  End If
	End Function
	'********************************************
	'函数名：IsValidEmail
	'作  用：检查Email地址合法性
	'参  数：email ----要检查的Email地址
	'返回值：True  ----Email地址合法
	'       False ----Email地址不合法
	'********************************************
	Public Function IsValidEmail(Email)
		Dim names, name, I, c
		IsValidEmail = True
		names = Split(Email, "@")
		If UBound(names) <> 1 Then IsValidEmail = False: Exit Function
		For Each name In names
			If Len(name) <= 0 Then IsValidEmail = False:Exit Function
			For I = 1 To Len(name)
				c = LCase(Mid(name, I, 1))
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then IsValidEmail = False:Exit Function
		   Next
		   If Left(name, 1) = "." Or Right(name, 1) = "." Then IsValidEmail = False:Exit Function
		Next
		If InStr(names(1), ".") <= 0 Then IsValidEmail = False:Exit Function
		I = Len(names(1)) - InStrRev(names(1), ".")
		If I <> 2 And I <> 3 Then IsValidEmail = False:Exit Function
		If InStr(Email, "..") > 0 Then IsValidEmail = False
	End Function
	'**************************************************
	'函数名：strLength
	'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
	'参  数：str  ----要求长度的字符串
	'返回值：字符串长度
	'**************************************************
	Public Function strLength(Str)
		'on error resume next
		Dim WINNT_CHINESE:WINNT_CHINESE = (Len("中国") = 2)
		If WINNT_CHINESE Then
			Dim l, T, c,I
			l = Len(Str)
			T = l
			For I = 1 To l
				c = Asc(Mid(Str, I, 1))
				If c < 0 Then c = c + 65536
				If c > 255 Then
					T = T + 1
				End If
			Next
			strLength = T
		Else
			strLength = Len(Str)
		End If
		If Err.Number <> 0 Then Err.Clear
	End Function
	'**************************************************
	'函数：FoundInArr
	'作  用：检查一个数组中所有元素是否包含指定字符串
	'参  数：strArr     ----字符串
	'        strToFind    ----要查找的字符串
	'       strSplit    ----数组的分隔符
	'返回值：True,False
	'**************************************************
	Function FoundInArr(strArr, strToFind, strSplit)
		Dim arrTemp, i
		FoundInArr = False
		If InStr(strArr, strSplit) > 0 Then
			arrTemp = Split(strArr, strSplit)
			For i = 0 To UBound(arrTemp)
			If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
				FoundInArr = True:Exit For
			End If
			Next
		Else
			If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then FoundInArr = True
		End If
	End Function
	'**************************************************
	'函数名：MakeRandomChar
	'作  用：生成指定位数的随机数字符串 如 "sJKD_!@KK"
	'参  数： Length  ----生成位数
	'返回值：成功返回随机字符串
	'**************************************************
	Function MakeRandomChar(Length)
	  Dim I, tempS, v
	  Dim c(65)
	   tempS = ""
	   c(1) = "a": c(2) = "b": c(3) = "c": c(4) = "d": c(5) = "e": c(6) = "f": c(7) = "g"
	   c(8) = "h": c(9) = "i": c(10) = "j": c(11) = "k": c(12) = "l": c(13) = "m": c(14) = "n"
	  c(15) = "o": c(16) = "p": c(17) = "q": c(18) = "r": c(19) = "s": c(20) = "t": c(21) = "u"
	  c(22) = "v": c(23) = "w": c(24) = "x": c(25) = "y": c(26) = "z": c(27) = "1": c(28) = "2"
	   c(29) = "3": c(30) = "4": c(31) = "5": c(32) = "6": c(33) = "7": c(34) = "8": c(35) = "9"
	  c(36) = "-": c(37) = "_": c(38) = "@": c(39) = "!": c(40) = "A": c(41) = "B": c(42) = "C"
	  c(43) = "D": c(44) = "E": c(45) = "F": c(46) = "G": c(47) = "H": c(48) = "I": c(49) = "J": c(50) = "K"
	  c(51) = "L": c(52) = "M": c(53) = "N": c(54) = "O": c(55) = "P": c(56) = "Q": c(57) = "R": c(58) = "S"
	  c(59) = "J": c(60) = "U": c(61) = "V": c(62) = "W": c(63) = "X": c(64) = "Y": c(65) = "Z"
	
	  If IsNumeric(Length) = False Then
		 MakeRandomChar = "":Exit Function
	  End If
	  For I = 1 To Length
		 Randomize
		 v = Int((65 * Rnd) + 1):tempS = tempS & c(v)
		 Next
		MakeRandomChar = tempS
	End Function
	'**************************************************
	'函数名：R
	'作  用：过滤非法的SQL字符
	'参  数：strChar-----要过滤的字符
	'返回值：过滤后的字符
	'**************************************************
	Function R(strChar)
		If strChar = "" Or IsNull(strChar) Then R = "":Exit Function
		Dim strBadChar, arrBadChar, tempChar, I
		'strBadChar = "$,#,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		arrBadChar = Split(strBadChar, ",")
		tempChar = strChar
		For I = 0 To UBound(arrBadChar)
			tempChar = Replace(tempChar, arrBadChar(I), "")
		Next
		tempChar = Replace(tempChar, "@@", "@")
		R = tempChar
	End Function
	'判断字符串是否为空,返回True/False
	Function IsNul(Str)
	  If Str="" Or IsNull(Str) Then IsNul=True Else IsNul=false
	End Function
'****************************************************
	'参数说明
	  'Subject     : 邮件标题
	  'MailAddress : 发件服务器的地址,如smtp.163.com
	  'LoginName     ----登录用户名(不需要请填写"")
	  'LoginPass     ----用户密码(不需要请填写"")
	  'Email       : 收件人邮件地址
	  'Sender      : 发件人姓名
	  'Content     : 邮件内容
	  'Fromer      : 发件人的邮件地址
	'****************************************************
	  Function SendMail(MailAddress, LoginName, LoginPass, Subject, Email, Sender, Content, Fromer)
	   'on error resume next
		Dim JMail
		  Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象
			jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值j
			jmail.Charset = "GB2312" '邮件的文字编码为国标
			jmail.ContentType = "text/html" '邮件的格式为HTML格式
			jmail.AddRecipient Email '邮件收件人的地址
			jmail.From = Fromer '发件人的E-MAIL地址
			jmail.FromName = Sender
			  If LoginName <> "" And LoginPass <> "" Then
				JMail.MailServerUserName = LoginName '您的邮件服务器登录名
				JMail.MailServerPassword = LoginPass '登录密码
			  End If

			jmail.Subject = Subject '邮件的标题 
			JMail.Body = Content
			JMail.Priority = 1'邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
			jmail.Send(MailAddress) '执行邮件发送（通过邮件服务器地址）
			jmail.Close() '关闭对象
		Set JMail = Nothing
		If Err Then
			SendMail = Err.Description
			Err.Clear
		Else
			SendMail = "OK"
		End If
	  End Function	
	  '分页SQL语句生成代码
		Function GetPageSQL(tblName,fldName,PageSize,PageIndex,OrderType,strWhere,fieldIds)
			Dim strTemp,strSQL,strOrder
			
			'根据排序方式生成相关代码
			if OrderType=0 then
				strTemp=">(select max([" & fldName & "])"
				strOrder=" order by [" & fldName & "] asc"
			else
				strTemp="<(select min([" & fldName & "])"
				strOrder=" order by [" & fldName & "] desc"
			end if
			
			'若是第1页则无须复杂的语句
			if PageIndex=1 then
			strTemp=""
			if strWhere<>"" then
			strTemp = " where " + strWhere
			end if
			strSQL = "select top " & PageSize & " " & fieldIds & " from [" & tblName & "]" & strTemp & strOrder
			else '若不是第1页，构造SQL语句
			strSQL="select top " & PageSize & " " & fieldIds & " from [" & tblName & "] where [" & fldName & "]" & strTemp & _
			" from (select top " & (PageIndex-1)*PageSize & " [" & fldName & "] from [" & tblName & "]" 
			if strWhere<>"" then
			strSQL=strSQL & " where " & strWhere
			end if
			strSQL=strSQL & strOrder & ") as tblTemp)"
			if strWhere<>"" then
			strSQL=strSQL & " And " & strWhere
			end if
			strSQL=strSQL & strOrder
			end if
			GetPageSQL=strSQL 
		End Function
 '======================================会员相关函数====================================
    '取得会员组选项--下拉列表  参数：Selected--默认选项
 Function GetUserGroup_Option(Selected)
	 Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
	  RSObj.Open "Select ID,GroupName From KS_UserGroup",Conn,1,1
	  	Do While Not RSObj.Eof
		   IF Selected=RSObj(0) Then
			GetUserGroup_Option=GetUserGroup_Option & "<option value=""" & RSObj(0) & """ Selected>" & RSObj(1) & "</option>"
		   Else
			GetUserGroup_Option=GetUserGroup_Option & "<option value=""" & RSObj(0) & """>" & RSObj(1) & "</option>"
		   End If
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	End Function
	 '取得会员组选项--多选列表 参数：SelectArr--默认选择项以","隔开,RowNum--每行显示选项数
 Function GetUserGroup_CheckBox(OptionName,SelectArr,RowNum)
	   Dim n:n=0
	   Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
	   IF RowNum<=0 Then RowNum=3
	   RSObj.Open "Select ID,GroupName From KS_UserGroup",Conn,1,1
	   GetUserGroup_CheckBox="<table width=""100%"" align=""center"" border=""0"">"
	   Do While Not RSObj.Eof
	        GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<TR>"
	     For N=1 To RowNum
		    GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<TD WIDTH=""" & CInt(100 / CInt(RowNum)) & "%"">"
			If FoundInArr(SelectArr,RSObj(0),",")<>0 Then
			 GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<input type=""checkbox"" checked name=""" & OptionName & """ value=""" & RSObj(0) & """>" & RSObj(1) & "&nbsp;&nbsp;"
			Else
			 GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<input type=""checkbox"" name=""" & OptionName & """ value=""" & RSObj(0) & """>" & RSObj(1) & "&nbsp;&nbsp;"
			End IF
		 GetUserGroup_CheckBox=GetUserGroup_CheckBox & "</TD>"
		 		RSObj.MoveNext
				If RSObj.Eof Then Exit For
		Next
		GetUserGroup_CheckBox=GetUserGroup_CheckBox & "</TR>"
		If RSObj.Eof Then Exit Do
	   Loop
	   GetUserGroup_CheckBox=GetUserGroup_CheckBox & "</TABLE>"
	   RSObj.Close:Set RSObj=Nothing
	End Function
	 
  	'取得用户组名称
 Function GetUserGroupName(GroupID)
	 'on error resume next
	 GetUserGroupName=Conn.Execute("Select GroupName From KS_UserGroup Where ID=" & GroupID)(0)
	 if err then GetUserGroupName=""
	End Function
 '取得用户组折扣
 Function GetUserZK(GroupID)
	 'on error resume next
	 GetUserZK=Conn.Execute("Select Discount From KS_UserGroup Where ID=" & GroupID)(0)
	 if err then GetUserZK=0
	End Function	
'**************************************************
	'函数名：ShowPage
	'作  用：显示“上一页 下一页”等信息
	'参  数：filename文件名 TotalNumber总数量 MaxPerPage每页数量 ShowTurn显示转到 PrintOut立即输出
	'**************************************************
	Function ShowPage_1(totalnumber, MaxPerPage, FileName, CurrPage,ShowTurn,PrintOut)
	             Dim n,j,startpage,pageStr,TotalPage,ParamStr
				 If totalnumber Mod MaxPerPage = 0 Then
						TotalPage = totalnumber \ MaxPerPage
				 Else
						TotalPage = totalnumber \ MaxPerPage + 1
				 End If
				 ParamStr=QueryParam("page") : If ParamStr<>"" Then ParamStr="&" & ParamStr	
				 n=0:startpage=1
				 pageStr=pageStr & "<div id='fenye' class='fenye'><table border=""0"" align=""right""><form action=""" & FileName & "?1=1" & ParamStr & """ name=""pageform"" method=""post""><tr><td>" & vbcrlf
				 if (CurrPage>1) then pageStr=PageStr & "<a href=""" & FileName & "?page=" & CurrPage-1 & ParamStr & """ class=""prev"">上一页</a>"
				 if CurrPage<>TotalPage and totalnumber>MaxPerPage then pageStr=PageStr & "<a href=""" & FileName & "?page=" & CurrPage+1 & ParamStr & """ class=""next"">下一页</a>"
				 pageStr=pageStr & "<a href=""" & FileName & "?page=1" & ParamStr & """ class=""prev"">首 页</a>"
				 if (CurrPage>=7) then startpage=CurrPage-5
				 if TotalPage-CurrPage<5 Then startpage=TotalPage-9
				 If startpage<0 Then startpage=1
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
				    Else
				     PageStr=PageStr & " <a class=""num"" href=""" & FileName & "?page=" &J& ParamStr & """>" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 pageStr=pageStr & "<a href=""" & FileName & "?page=" & TotalPage & ParamStr & """ class=""prev"">末页</a>"
				 pageStr=PageStr & " <span>分" & TotalPage & "页"
				 If ShowTurn=true Then
				 If CurrPage=TotalPage Then CurrPage=0
				 pageStr=PageStr & " 转到:<input type='text' value='" & (CurrPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>"
				 End If
				 PageStr=PageStr & "</span></td></tr></form></table></div>"
				If PrintOut=true Then Response.Write PageStr Else ShowPage=PageStr
	End Function	
'取得QueryString,或Form参数集合,参数NoCollect表示不收集的字段,多个用英文逗号隔开
	Function QueryParam(NoCollect)
		 Dim Param,R
		 For Each r In Request.QueryString
		  If FoundInArr(NoCollect,R,",")=false Then
			  If Request.QueryString(r)<>"" Then
				If Param="" Then
				 Param=r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				Else
				 Param=Param & "&" & r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				End If
			  End If
		 End If
		 Next
		' If Param<>"" Then QueryParam=Param:Exit Function
		 For Each r In Request.Form
		  If FoundInArr(NoCollect,R,",")=false Then
			  If Request.Form(r)<>"" Then
				If Param="" Then
				 Param=r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				Else
				 Param=Param & "&" & r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				End If
			  End If
		 End If
		 Next
		 QueryParam=Param
	End Function
	Function FilterIDs(byval strIDs)
	Dim arrIDs,i,strReturn
	strIDs=Trim(strIDs)
	If Len(strIDs)=0  Then Exit Function
	arrIDs=Split(strIDs,",")
	For i=0 To Ubound(arrIds)
		If ChkClng(Trim(arrIDs(i)))<>0 Then
			strReturn=strReturn & "," & Int(arrIDs(i))
		End If
	Next
	If Left(strReturn,1)="," Then strReturn=Right(strReturn,Len(strReturn)-1)
	FilterIDs=strReturn
	End Function	
	 Function U_G(GroupID,FieldName)
	   Dim Node:Set Node=Application(SiteSN&"_UserGroup").DocumentElement.selectSingleNode("row[@id=" & GroupID & "]/@" & Lcase(FieldName))
	   If Not Node Is Nothing Then U_G=Node.text
	   Set Node=Nothing
	 End Function
	 
'*********************************************************************************************************
		'函数名：GetSingleFieldValue
		'作  用：取单字段值
		'参  数：SQLStr SQL语句
		'*********************************************************************************************************
		Function GetSingleFieldValue(SQLStr)
			'on error resume next
			GetSingleFieldValue=Conn.Execute(SQLStr)(0)
			If Err Then GetSingleFieldValue=""
		End Function
		
		'**************************************************
	'函数名：AlertHistory
	'作  用：弹出警告消息后,停止所在页面的执行,返回n级。
	'参  数：SuccessStr  ----成功提示信息
	'        n   ------返回级数
	'返回值：无
	'**************************************************
	Public Function AlertHistory(SuccessStr, N)
		c2 ("<script language=""Javascript""> alert('" & SuccessStr & "');history.back(" & N & ");</script>")
	End Function
	'会员注册表单
sub UserReg()
	Dim FieldsList:FieldsList=GetSingleFieldValue("Select FormField From UserForm Where ID=68")
	  Dim FileContent:FileContent=GetSingleFieldValue("Select Template From UserForm Where ID=68")
		   Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType,MustFillTF,title from UserField Where ChannelID=101 Order By OrderID",conn,1,1
		   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
		   For K=0 TO Ubound(SQL,2)
		     If FoundInArr(FieldsList,SQL(0,k),",") Then
			  InputStr="":xing="":kong="empty:true,"
			  if SQL(8,K)=1 then xing="*":kong=""
			  If lcase(SQL(2,K))="province&city" Then
				 InputStr="<script language=""javascript"" src=""cenbel/area.asp""></script></td><td width=""630""><div id=""ProvinceTip"">"&xing&"</div>"
				 CheckForm=CheckForm&"$(""#Province"").formValidator({"&kong&"tipid:""ProvinceTip"",onshow:""请选择省份"",onfocus:""省份必须选择"",oncorrect:""请选择城市!""}).inputValidator({min:1,onerror: ""省份必须选择!""});"
				 CheckForm=CheckForm&"$(""#City"").formValidator({"&kong&"tipid:""ProvinceTip"",onshow:""请选择城市"",onfocus:""城市必须选择"",oncorrect:""谢谢你的配合""}).inputValidator({min:1,onerror: ""城市必须选择!""});"
			  Else
			  Select Case SQL(1,K)
			    Case 2:InputStr="<textarea style=""width:" & SQL(4,K) & "px;height:" & SQL(5,K) & "px"" rows=""5"" class=""textbox"" name=""" & SQL(2,K) & """>" & SQL(3,K) & "</textarea></td><td><div id=""" & SQL(2,K) & "Tip"">"&xing&"</div>"
				Case 3
				  InputStr="<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """>"
				  O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
					 F_V=Split(O_Arr(N),"|")
					 If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					 Else
						O_Value=F_V(0):O_Text=F_V(0)
					 End If						   
					 If SQL(3,K)=O_Value Then
						InputStr=InputStr & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
					 Else
						InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
					 End If
				  Next
					InputStr=InputStr & "</select></td><td width=""630""><div id=""" & SQL(2,K) & "Tip"">"&xing&"</div>"
				Case 6
					 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
					 If O_Len>1 And Len(SQL(6,K))>50 Then BrStr="<br>" Else BrStr=""
					 For N=0 To O_Len
						F_V=Split(O_Arr(N),"|")
						If Ubound(F_V)=1 Then
						 O_Value=F_V(0):O_Text=F_V(1)
						Else
						 O_Value=F_V(0):O_Text=F_V(0)
						End If						   
					    If SQL(3,K)=O_Value Then
							InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr&"</td><td width=""630""><div id=""" & SQL(2,K) & "Tip"">"&xing&"</div>"
						Else
							InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text & BRStr&"</td><td width=""630""><div id=""" & SQL(2,K) & "Tip"">"&xing&"</div>"
						 End If
			         Next
			  Case 7
					O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
					 For N=0 To O_Len
						  F_V=Split(O_Arr(N),"|")
						  If Ubound(F_V)=1 Then
							O_Value=F_V(0):O_Text=F_V(1)
						  Else
							O_Value=F_V(0):O_Text=F_V(0)
						  End If						   
						  If KS.FoundInArr(SQL(3,K),O_Value,",")=true Then
								 InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text&"</td><td width=""630""><div id=""" & SQL(2,K) & "Tip""></div>"
						 Else
						  InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text&"</td><td width=""630""><div id=""" & SQL(2,K) & "Tip"">"&xing&"</div>"
						 End If
				   Next
			  Case 10
					InputStr=InputStr & "<input type=""hidden"" id=""" & SQL(2,K) &""" name=""" & SQL(2,K) &""" value="""& Server.HTMLEncode(SQL(3,K)) &""" style=""display:none"" /><iframe id=""" & SQL(2,K) &"___Frame"" src=""../../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=" & SQL(2,K) &"&amp;Toolbar=" & SQL(7,K) & """ width=""" &SQL(4,K) &""" height=""" & SQL(5,K) & """ frameborder=""0"" scrolling=""no""></iframe></td><td width=""630""><div id=""" & SQL(2,K) & "Tip"">"&xing&"</div>"				
			  Case Else
			    If lcase(SQL(2,K))="mobile" Then
			  InputStr="<input type=""text"" class=""textbox"" style=""width:" & SQL(4,K) & "px"" name=""" & SQL(2,K) & """ id=""" & SQL(2,K) & """ value=""" & S("Mobile") & """></td><td width=""630""><div id=""" & SQL(2,K) & "Tip""></div>"
			   CheckForm=CheckForm&"$(""#" & SQL(2,K) & """).formValidator({"&kong&"onshow:""请输入" & SQL(9,K) & """,onfocus:""请认真填写" & SQL(9,K) & """,oncorrect:""输入正确!""}).regexValidator({regexp:""mobile"",datatype:""enum"",onerror:""" & SQL(9,K) & "格式不正确""});"
				Else
			  InputStr="<input type=""text"" class=""textbox"" style=""width:" & SQL(4,K) & "px"" name=""" & SQL(2,K) & """ id=""" & SQL(2,K) & """ value=""" & SQL(3,K) & """></td><td><div id=""" & SQL(2,K) & "Tip""></div>"
			  CheckForm=CheckForm&"$(""#" & SQL(2,K) & """).formValidator({"&kong&"onshow:""请输入" & SQL(9,K) & """,onfocus:""请认真填写" & SQL(9,K) & """,oncorrect:""输入正确!""}).inputValidator({min:1,onerror:""" & SQL(9,K) & "不能为空,请确认""});"
			    End If
			  End Select
			  End If
			  'if SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?UPType=Field&FieldID=" & SQL(2,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
			  If Instr(FileContent,"{@NoDisplay(" & SQL(2,K) & ")}")<>0 Then
			   FileContent=Replace(FileContent,"{@NoDisplay(" & SQL(2,K) & ")}"," style='display:none'")
			  End If
			   FileContent=Replace(FileContent,"[@" & SQL(2,K) & "]",InputStr)
			  End If
		   Next
		    FileContent=Replace(FileContent,"{@NoDisplay}","")
			Response.Write(FileContent)
	end sub 
'**************************************************
			'函数名：UserLoginChecked
			'作  用：判断用户是否登录
			'返回值：true或false
			'**************************************************
Function UserLoginChecked()
                 'on error resume next
				UserName = R(Trim(C("UserName")))
				PassWord= R(Trim(C("Password")))
				RndPassword=R(Trim(C("RndPassword")))
				IF UserName="" Then
				   UserLoginChecked=false
				   Exit Function
				Else
					Dim UserRs:Set UserRS=Server.CreateOBject("ADODB.RECORDSET")
					UserRS.Open "Select top 1 a.*,b.SpaceSize From [User] a inner join KS_UserGroup b on a.groupid=b.id Where UserName='" & UserName & "' And PassWord='" & PassWord & "'",Conn,2,3
					IF UserRS.Eof And UserRS.Bof Then
					  UserLoginChecked=false
					Else
						  UserLoginChecked=true
						  for I=0 to UserRS.fields.count-1
							if lcase(UserRS.Fields(i).Name)="sign" then
							  Sign=UserRS.Fields(i).Value
							else
							 Execute(UserRS.Fields(i).Name&"=ForValue("""&trim(UserRS.Fields(i).Value)&""")")
							 
							end if
						  Next
					End if
					UserRS.Close:Set UserRS=Nothing
			   End IF
			End Function

Sub ShowPageParamter(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		Response.Write(ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr))
	End Sub
'**************************************************
	'函数名：ShowPagePara
	'作  用：显示“上一页 下一页”等信息
	'参  数：filename  ----链接地址
	'       TotalNumber ----总数量
	'       MaxPerPage  ----每页数量
	'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。
	'       strUnit     ----计数单位,CurrentPage--当前页,ParamterStr参数
	'返回值：无返回值
	'**************************************************
	Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "<span style='font-family:webdings;font-size:14px' title='第一页'>9</span>" '定义第一页按钮显示样式
				Const Btn_Prev = "<span style='font-family:webdings;font-size:14px' title='上一页'>3</span>" '定义前一页按钮显示样式
				Const Btn_Next = "<span style='font-family:webdings;font-size:14px' title='下一页'>4</span>" '定义下一页按钮显示样式
				Const Btn_Last = "<span style='font-family:webdings;font-size:14px' title='最后一页'>:</span>" '定义最后一页按钮显示样式
				  PageStr = ""
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
				If N > 1 Then
					PageStr = PageStr & ("<div class='showpage' style='height:20px'><form action=""" & FileName & "?" & ParamterStr & """ name=""myform"" method=""post"">页次：<font color=red>" & CurrentPage & "</font>/" & N & "页 共有:" & totalnumber & strUnit & " 每页:" & MaxPerPage & strUnit & " ")
					If CurrentPage < 2 Then
						PageStr = PageStr & Btn_First & " " & Btn_Prev & " "
					Else
						PageStr = PageStr & ("<a href=" & FileName & "?page=1" & "&" & ParamterStr & ">" & Btn_First & "</a> <a href=" & FileName & "?page=" & CurrentPage - 1 & "&" & ParamterStr & ">" & Btn_Prev & "</a> ")
					End If
					
					If N - CurrentPage < 1 Then
						PageStr = PageStr & " " & Btn_Next & " " & Btn_Last & " "
					Else
						PageStr = PageStr & (" <a href=" & FileName & "?page=" & (CurrentPage + 1) & "&" & ParamterStr & ">" & Btn_Next & "</a> <a href=" & FileName & "?page=" & N & "&" & ParamterStr & ">" & Btn_Last & "</a> ")
					End If
					If ShowAllPages = True Then
						PageStr = PageStr & ("转到:<input type='text' value='" & (CurrentPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>")
				  End If
				  PageStr = PageStr & "</form></div>"
			 End If
			 ShowPagePara = PageStr
	End Function
'付款方式
  Function GetPaymentTypeStr()
   Dim DiscountStr,SQL,I,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select TypeID,TypeName,IsDefault,Discount from PaymentType order by orderid",conn,1,1
   If Not RS.Eof Then
     SQL=RS.GetRows(-1)
   End IF
   RS.Close:Set RS=Nothing
   GetPaymentTypeStr="<table width='100%' border='0' cellpadding='2' cellspacing='1' class='tdbg'>"
   For I=0 To UBound(SQL,2)
     If SQL(3,I)<>100 Then
	  DiscountStr="折扣率 " & SQL(3,I) & "%"
	 Else
	  DiscountStr=""
	 End iF
     If SQL(2,I)=1 Then
    GetPaymentTypeStr=GetPaymentTypeStr& "<tr class='tdbg3'><td width='20'><input type='radio' name='PaymentType' value='" & SQL(0,I) & "' checked></td><td width='60'>"  &SQL(1,I) & "</td><td>" & DiscountStr & "</td> </tr>"
	 Else
    GetPaymentTypeStr=GetPaymentTypeStr& "<tr class='tdbg3'><td width='20'><input type='radio' name='PaymentType' value='" & SQL(0,I) & "'></td><td width='60'>"  &SQL(1,I) & "</td><td>" & DiscountStr & "</td> </tr>"
	End If
   Next
   GetPaymentTypeStr=GetPaymentTypeStr & "</table>"
  End Function
  '发货方式
  Function GetDeliveryTypeStr()
   Dim DiscountStr,SQL,I,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
   RS.Open "select TypeID,TypeName,IsDefault,fee from Delivery order by orderid",conn,1,1
   If Not RS.Eof Then
     SQL=RS.GetRows(-1)
   End IF
   RS.Close:Set RS=Nothing
   GetDeliveryTypeStr="<table width='100%' border='0' cellpadding='2' cellspacing='1' class='tdbg'>"
   For I=0 To UBound(SQL,2)
     If SQL(3,I)=0 Then
	  DiscountStr="<font color=blue>免费</font>"
	 Else
	  DiscountStr="<font color=red>加收 ＄" & SQL(3,I) & " 元</font>"
	 End iF
     If SQL(2,I)=1 Then
    GetDeliveryTypeStr=GetDeliveryTypeStr& "<tr class='tdbg3'><td width='20'><input type='radio' name='DeliverType' value='" & SQL(0,I) & "' checked></td><td width='60'>"  &SQL(1,I) & "</td><td>" & DiscountStr & "</td> </tr>"
	 Else
    GetDeliveryTypeStr=GetDeliveryTypeStr& "<tr class='tdbg3'><td width='20'><input type='radio' name='DeliverType' value='" & SQL(0,I) & "'></td><td width='60'>"  &SQL(1,I) & "</td><td>" & DiscountStr & "</td> </tr>"
	End If
   Next
   GetDeliveryTypeStr=GetDeliveryTypeStr & "</table>"
  End Function
 '返回付款方式名称,参数TypeID,0名称 1折扣率
	Function ReturnPayment(ID,TypeID)
	  If Application(SiteSn &"Payment_" & ID&TypeID)="" Then
         Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select TypeName,Discount From PaymentType Where TypeID=" & ID,conn,1,1
		 If Not RS.Eof Then
		     If TypeID=0 Then
		  	  ReturnPayment=rs(0)
			  If RS(1)<100 Then ReturnPayment=ReturnPayment & "&nbsp;&nbsp;<font color=red>折扣率:" & RS(1) & "%"
			 Else
			  ReturnPayment=rs(1)
			 End if
		End iF 
		Application(SiteSn &"Payment_" & ID&TypeID)=ReturnPayment
	  Else
	    ReturnPayment=Application(SiteSn &"Payment_" & ID&TypeID)
	  End If
	End Function
		'返回收货方式名称,参数TypeID,0名称 1费用
	Function ReturnDelivery(ID,TypeID)
	  If Application(SiteSn &"Delivery_" & ID&TypeID)="" Then
         Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select TypeName,fee From Delivery Where TypeID=" & ID,conn,1,1
		 If Not RS.Eof Then
		     If TypeID=0 Then
		  	  ReturnDelivery=rs(0)
			  If RS(1)=0 Then ReturnDelivery=ReturnDelivery & "&nbsp;<font color=blue>免费</font>" Else ReturnDelivery=ReturnDelivery & "&nbsp;<font color=red>加收 " & RS(1) & "元"
			 Else
			  ReturnDelivery=rs(1)
			 End iF
		End iF 
		Application(SiteSn &"Delivery_" & ID&TypeID)=ReturnDelivery
	  Else
	    ReturnDelivery=Application(SiteSn &"Delivery_" & ID&TypeID)
	  End If
	End Function
	
	'**************************************************
	'函数名：MakeRandom
	'作  用：生成指定位数的随机数
	'参  数： maxLen  ----生成位数
	'返回值：成功:返回随机数
	'**************************************************
	Public Function MakeRandom(ByVal maxLen)
	  Dim strNewPass,whatsNext, upper, lower, intCounter
	  Randomize
	 For intCounter = 1 To maxLen
	   upper = 57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
	End Function 
	
	Function C_S(sChannelID,FieldID)
	  'on error resume next
	  If not IsObject(Application(SiteSN&"_ChannelConfig")) Then LoadChannelConfig()
	   C_S=Application(SiteSN&"_ChannelConfig").documentElement.selectSingleNode("channel[@ks0=" & sChannelID & "]/@ks" & FieldID & "").text
	   if err then C_S=0:err.Clear
	 End Function
'*************************************************************************************************
		'函数名：ReplacePrevNext
		'作  用：上一篇、下一篇
		'参  数：
		'		ChannelID 目录ID
		'		NowID 现在ID
		'		TypeStr(<:上一篇 >:下一篇)
		'		fromtype 数据类型  1为产品数据库  0为文章数据库
		'		title 显示内容如:上一篇(当为空时显示标题)
		'*************************************************************************************************
Function ReplacePrevNext(ChannelID,NowID,TypeStr,fromtype,title)
		     Dim SqlStr
			 if fromtype=1 then
				 SqlStr=" SELECT Top 1 ID,Title,Cate"
				 SqlStr=SqlStr & " From ProductInfo Where Cate=" & ChannelID & " And ID" & TypeStr & NowID & " And Locked=0 Order By OIndex"
			 else
			 	SqlStr=" SELECT Top 1 ID,Title,ChannelID"
			 SqlStr=SqlStr & " From NewsInfo Where ChannelID=" & ChannelID & " And ID" & TypeStr & NowID & " And Locked='0' Order By OIndex"
			 end if
			 If TypeStr=">" Then SqlStr=SqlStr & " asc" else SqlStr=SqlStr & " desc"
			 Dim RS:Set RS=Conn.Execute(SqlStr)
			 If RS.EOF And RS.BOF Then
			 	if title="" then
					ReplacePrevNext="没有了"
				else
				 	ReplacePrevNext = title
				end if
			 
			 Else
			 if title="" then title=rs("title")
			  ReplacePrevNext = "<a href=""?channelId="&ChannelID&"&id="&rs(0)&""" title=""" & RS(1) & """>" & title & "</a>"
			 End If
			 RS.Close:Set RS = Nothing
		End Function
'函数名:QQList
'作用:显示QQ列表横排
'参数:0=不显示名称 1=反之
function QQList(types)
set rs=server.CreateObject("adodb.recordset")
sql="select * from QQlist order by q_paixu"
rs.open sql,conn,1,1
if not (rs.eof or rs.bof) then
QQList="<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
QQList=QQList&"  <tr>"
do while not rs.eof
	if chkclng(rs("q_s_id"))=0 then
		QQList=QQList&"<td><a target=blank href=http://wpa.qq.com/msgrd?V=1&amp;Uin="&rs("q_num")&"&amp;Site="&Setting(0)&"&amp;Menu=yes><img src=cenbel/qqsever/images/qqface/"&rs("q_face")&"_f.gif border=0>&nbsp;"
			if chkclng(types)=1 then
			QQList=QQList&"<font style=font-size:12px;TEXT-DECORATION:none;color:"&rs("q_color")&";>"&rs("q_name")&"</font>"
			end if
		QQList=QQList&"</a></td>"
	elseif chkclng(rs("q_s_id"))=1 then
		QQList=QQList&"<td><a target=blank href=msnim:chat?contact="&rs("q_num")&"&amp;Site="&Setting(0)&"&amp;Menu=yes><img src=cenbel/qqsever/images/msn/"&rs("q_face")&"_f.gif border=0>&nbsp;"
		if chkclng(types)=1 then
			QQList=QQList&"<font style=font-size:12px;TEXT-DECORATION:none;color:"&rs("q_color")&";>"&rs("q_name")&"</font>"
			end if
		QQList=QQList&"</a></td>"
	else
		QQList=QQList&"<td><a target=blank href=Skype:"&rs("q_num")&"?call><img src=cenbel/qqsever/images/skype/"&rs("q_face")&"_f.gif border=0>&nbsp;"
		if chkclng(types)=1 then
			QQList=QQList&"<font style=font-size:12px;TEXT-DECORATION:none;color:"&rs("q_color")&";>"&rs("q_name")&"</font>"
			end if
		QQList=QQList&"</a></td>"
	end if
rs.movenext
loop
QQList=QQList&"</tr>"
QQList=QQList&"</table>"
end if
rs.close
set rs=nothing
end function	
	%>
