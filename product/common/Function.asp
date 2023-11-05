<%
Rem<<====================================  语言版本信息公共函数开始 ===============================

Rem  程序作用  -- 获取语言版本信息
Sub getLanID()
	LanID = Request.QueryString("LanID")
	if Trim(LanID)="" then
		if Trim(Session("LanID"))="" then Session("LanID")="1"
	else
		if Trim(LanID)="2" then
			Session("LanID")="2"
		else
			Session("LanID")="1"
		end if
	end if
End Sub

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

Rem  程序作用  -- 获取网站类别信息
Rem  CateID    -- 栏目ID
Rem  CID       -- 要获取的字段名称
Function getCateInfo(ParentID,CID)
    if Trim(CID)="" or isNumeric(ParentID)=False then
	    getCateInfo = "Param Error"
		Exit Function
	end if
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select ["&CID&"] From [CateInfo] Where [ID]="&ParentID&" "),conn,1,1
	if objRs.bof or objRs.eof then
	    getCateInfo = "0"
	else
	    getCateInfo = objRs(0)
	end if
	objRs.Close
	Set objRs = Nothing	
End Function

Rem  程序作用  -- 获取网站子类别信息
Rem  CateID    -- 栏目ID
Rem  CID       -- 要获取的字段名称
Function getInfo(ParentID,CID)
    if Trim(CID)="" or isNumeric(ParentID)=False then
	    getCateInfo = "Param Error"
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
		objrs.movenext
		loop
	end if
	objRs.Close
	Set objRs = Nothing	
End Function

Rem  程序作用  -- 网站类别路径信息
Rem  CateTree  -- 类别ID树
Function getCatePath(CateTree)
    ArrCate = Split(CateTree,",")	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	For M = 0 To UBound(ArrCate)
		if ArrCate(M)=0 then
		    getCatePath = "TOP"
		else
			objRs.Open("Select TOP 1 [Title] From [CateInfo] "&_
			" Where [ID]="&ArrCate(M)),conn,1,1
			if not (objRs.bof or objRs.eof) then
				if getCatePath = "" then 
					getCatePath = "<a href=""ChannelList.asp?ParentID="&ArrCate(M)&""">"&_
					""&objRs(0)&"</a>"
				else
					getCatePath = getCatePath & " > " & "<a "&_
					" href=""CateList.asp?ParentID="&ArrCate(M)&""">"&objRs(0)&"</a>"
				end if
				objRs.Close
			end if
		end if
	Next
	Set objRs = Nothing	
End Function

Function getCatePaths(CateTree)
    ArrCate = Split(CateTree,",")
		Set objRs = Server.CreateObject("Adodb.Recordset")
		For M = 0 To UBound(ArrCate)
			if ArrCate(M)<>0 then
				objRs.Open("Select TOP 1 [Title] From [CateInfo] Where [ID]="&ArrCate(M)),conn,1,1
				if not (objRs.bof or objRs.eof) then
					if getCatePaths = "" then 
						getCatePaths = objRs(0)
					else
						getCatePaths = getCatePaths & " > "&objRs(0)
					end if
					objRs.Close
				end if
			end if
		Next
		if getCatePaths="" then getCatePaths="全部列表"
		Set objRs = Nothing	
End Function


Rem  程序作用  -- 网站类别下拉选择列表
Rem  ParentID  -- 所属类别ID
Rem  LevelID   -- 类别层级ID
Rem  CateID -- 当前类别ID
Function getCateSelect( ParentID, LevelID, CateID )	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	objRs.Open("Select [ID],[Title] From [CateInfo] Where [ParentID]="&ParentID),conn,1,1
	while not objRs.eof
		LevelLine = ""
		For L = 1 To LevelID 
		    LevelLine = LevelLine & "&nbsp;&nbsp;&nbsp;&nbsp;"
		Next 
		if Trim(objRs(0)) = Trim(CateID) then 
			Response.Write"<option value='"&objRs(0)&"' selected>"&LevelLine&""&objRs(1)&"</option>"
		else
			Response.Write("<option value='"&objRs(0)&"'>"&LevelLine&""&objRs(1)&"</option>")
		end if		
		getCateSelect objRs(0), Int(LevelID) + 1, CateID 
		objRs.MoveNext
	wend
	objRs.Close
	Set objRs = Nothing	
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


Rem  程序作用 -- 获取类别ID树
Rem  ParentID -- 父类别ID
Function getCateTree( ParentID )	
	if Trim(ParentID) = "0" then
	    getCateTree = "0"
	else
	    getCateTree = getCateInfo(ParentID,"CateTree") & "," & ParentID
	end if
End Function

Function getCateList(ParentID,CateList)	 
	Set objRss = Server.CreateObject("Adodb.Recordset")
	objRss.Open(" Select [ID],(Select Count(*) From [CateInfo] Where [ParentID]=P.[ID]) AS [SON] "&_
	            " From [CateInfo] P Where [ParentID]="&ParentID ),conn,1,1
	while not objRss.eof
	    if Trim(CateList) = "" then
		    CateList = objRss(0)
		else
			CateList = CateList & "," & objRss(0)
		end if
		if objRss(1) > 0 then getCateList objRss(0), CateList 
		objRss.Movenext		
	wend
	objRss.Close
	Set objRss = Nothing
	getCateList = CateList 
	if Trim(getCateList) = "" then
		getCateList = ParentID
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
    ArrChannel = Split(ChannelTree,",")	
	Set objRs = Server.CreateObject("Adodb.Recordset")
	For M = 0 To UBound(ArrChannel)			
		if ArrChannel(M)=0 then
		    getChannelPath = ""
		else
			objRs.Open("Select TOP 1 [Title] From [ChannelInfo] "&_
			" Where [ID]="&ArrChannel(M)),conn,1,1
			if not (objRs.bof or objRs.eof) then
				if getChannelPath = "" then 
					getChannelPath =" > "
				else
					getChannelPath = getChannelPath & " > " &objRs(0)&""
				end if
				objRs.Close
			end if
		end if
	Next
	Set objRs = Nothing	
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
    getModelSelect = "<input type='radio' name='ChannelModel' value='0' checked onClick=""ChangeType(0)"">栏目列表"
	
	getModelSelect = getModelSelect & "<input type='radio' name='ChannelModel' value='1' onClick=""ChangeType(1)"" "
	if Trim(ID)="1" then getModelSelect = getModelSelect & " checked "
	getModelSelect = getModelSelect & ">新闻列表"
	
	getModelSelect = getModelSelect & "<input type='radio' name='ChannelModel' value='2' onClick=""ChangeType(2)"""
	if Trim(ID)="2" then getModelSelect = getModelSelect & " checked "
	getModelSelect = getModelSelect & ">单篇文章"
	
	getModelSelect = getModelSelect & "<input type='radio' name='ChannelModel' value='3' onClick=""ChangeType(3)"""
	if Trim(ID)="3" then getModelSelect = getModelSelect & " checked "
	getModelSelect = getModelSelect & ">产品新闻"
	
	getModelSelect = getModelSelect & "<input type='radio' name='ChannelModel' value='4' onClick=""ChangeType(0)"""
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
		    getModelLink = "PNewsList.asp?CateID="			
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
	
	Sub ShowError(Errmsg)
	    With Response
		.Write ("<br><br><div align=""center"">")
		.Write ("  <center>")
		.Write ("  <table border=""0"" cellpadding='2' cellspacing='1' class='border' width=""75%"" style=""MARGIN-TOP: 3px"" class='border'>")
		.Write ("	 <tr class=tdbg>")
		.Write ("			  <td width=""100%"" height=""30"" align=""center"">")
		.Write ("				<b> " & Errmsg & "&nbsp; </b>")
		.Write ("				</b>")
		.Write ("			  </td>")
		.Write ("	 </tr>")
		.Write ("	 <tr  class=tdbg>")
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
	

function isHttp(str)
	tempStr = ""
	if str<>"" And not isNull(str) then
		if Left(LCase(Trim(str)),len("http://"))<>"http://" then
			if Left(LCase(Trim(str)),len("uploadimages/"))<>"uploadimages" then
				tempStr = "uploadimages/image/Pro/small/" & str
			end if
		else
			tempStr = str
		end if
	else
		tempStr = str
	end if
	isHttp = tempStr
end function
function isHttp2(str)
	tempStr = ""
	if str<>"" And not isNull(str) then
		if Left(LCase(Trim(str)),len("http://"))<>"http://" then
			if Left(LCase(Trim(str)),len("../uploadimages/"))<>"../uploadimages" then
				tempStr = "uploadimages/image/Pro/big/" & str
			end if
		else
			tempStr = str
		end if
	else
		tempStr = str
	end if
	isHttp2 = tempStr
end function
function topmuen(str)
dim GetAutoDomain
topmuen=true
GetAutoDomain =LCase(Request.ServerVariables("Path_Info"))
	if instr(GetAutoDomain,Lcase(str))>0 then
	topmuen=true
	else
	topmuen=False
	end if
end function
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
	
		'检查是否是数字 ，并转换为长整型
	Public Function ChkClng(ByVal str)
	    'on error resume next
		If IsNumeric(str) Then
			ChkClng = CLng(str)
		Else
			ChkClng = 0
		End If
		If Err Then ChkClng=0
	End Function 	  %>
