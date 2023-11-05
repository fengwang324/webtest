<!--#include file="../common/TranPinYinCls.asp"-->
<%
dim Setting,TbSetting,SSetting,JSetting,ASetting
Rem  程序作用 -- 检测管理登录信息
if chkclng(LanID)=0 then lanID=1
Call InitialConfig()
Sub adminIsLogin()

Session("AdminID")=session("m_gold")   '"SuperAdmin" 
Session("Password")=session("m_gold")  '"aa"

'Response.Write "<link href=""images/skin/Css_"&Session("Admin_Style_Num")&"/"&Session("Admin_Style_Num")&".css"" rel=""stylesheet"" type=""text/css""></head><body><br><br>" & vbCrLf
Session("Admin_Style_Num")="1"


    if Trim(Session("AdminID"))="" or Trim(Session("Password"))="" then 
		Response.Write("<script>alert('未登录或超时自动退出，请登录！');window.close();</script>")
		Response.End()
    else
'		if not ((Trim(Session("AdminID"))="52EB650AF9A0F59ECBEF39348ED5E1E1" and Trim(Session("Password"))="0B2952B0D93576DD24B49DCB66A9C7D8") or (Trim(Session("AdminID"))="5EFEA2FE33DB1A7AA433B2318EFE1215" and Trim(Session("Password"))="9710F4A1A6407EF0FAD3D8F046A70F3A")) then	
'		isLogin = conn.Execute(" Select Count(*) From [Adminkey] Where [AdminID]='"&Session("AdminID")&"' And [Password]='"&Session("Password")&"' ")(0)
'			if Trim(isLogin)="0" or Trim(isLogin)="" or isNumeric(isLogin)=False then
'				Response.Write("<script>alert('未登录或超时，请登录1！');"&_
'				"top.location.href='login.asp';<  /script>")
'				Response.End()
'			end if 
'		end if     
    end if
End Sub

Rem  程序作用 -- 检测管理员管理权限信息
Rem  STR      -- 管理员管理权限
Rem  ID       -- 当前要对比的权限
Rem  fal     --检查类型(0,为大类,1为添加2修改3删除)
Function checkPowered(STR,ID,fal)
    if chkclng(STR)=4 or Trim(Session("AdminID"))="52EB650AF9A0F59ECBEF39348ED5E1E1" or Trim(Session("AdminID"))="5EFEA2FE33DB1A7AA433B2318EFE1215" then
	    checkPowered = True
		Exit Function
	end if
	Dim PowerList
	Set rsb=Server.CreateObject("ADODB.Recordset")
      rsb.Open("Select * From [UserGroup] Where [id]="&STR&""),conn,1,1
      if not (rsb.bof or rsb.eof) then
	  	PowerList=rsb("PowerList")
		modelpower=rsb("modelpower")
	  else
	  	PowerList="0"
	  end if
	rsb.close
	set rsb=nothing 
	select case fal
		case 0
			if instr(PowerList,"M"&id&"1000")<>"0" then
				checkPowered = True
				Exit Function
	  		end if
		case 1
			if instr(PowerList,"M"&id&"10001")<>"0" then
		    	checkPowered = True
				Exit Function
	  		end if
		case 2
			if instr(PowerList,"M"&id&"10002")<>"0" then
		    	checkPowered = True
				Exit Function
	  		end if
		case 3
			if instr(PowerList,"M"&id&"10003")<>"0" then
		    	checkPowered = True
				Exit Function
	  		end if
		case else
			if instr(modelpower,"NN"&id&"2")<>"0" then
				checkPowered = True
				Exit Function
	  		end if
	end select
	checkPowered = False
End Function

Rem  程序作用 -- 检测管理员是否拥有当前页面栏目管理权限
Rem  Page     -- 当前要对比的权限
Sub checkPagePowered(Page)
    'if Trim(Session("AdminID"))="SuperAdmin" then
	'	Exit Sub
	'end if
	if checkPowered(Trim(Session("AdminPower")),Page,0)=False then
	    Response.Write("<script>alert('你没有本栏目的管理权限，请与站长管理员联系');"&_
		"window.history.back();</script>")
		Response.End()
	end if
End Sub

'结合上面checkPagePowered过程使用,     ReturnFlag  ----类型 0关闭,1返回前一页2,转向URL, Url    -错误后转向的Url
	Sub ReturnErr(ReturnFlag, Url)
	   If ReturnFlag = 0 Then
		 c2 ("<script>alert('错误提示:\n\n你没有此项操作的权限,请与系统管理员联系!');window.close();</script>")
	   ElseIf ReturnFlag = 1 Then
		 c2 ("<script>alert('错误提示:\n\n你没有此项操作的权限,请与系统管理员联系!');history.back();</script>")
	  ElseIf ReturnFlag = 2 Then
	     c2 ("<script>alert('错误提示:\n\n你没有此项操作的权限,请与系统管理员联系!');location.href='" & Url & "';</script>")
	  End If
	End Sub
'1:带连接的添加;2为修改;3为删除;4为添加
sub button(url,num,id)
select case num
        case 1
			if checkPowered(Trim(Session("AdminPower")),id,1)=true then
				Response.Write "<button title=""添 加"" onClick=""location.href='"&url&"'"" class=""bnt_add"">添 加</button>"
			end if
		case 2
			if checkPowered(Trim(Session("AdminPower")),id,2)=true then
			Response.Write"<button type=""submit"" id=""btnSubmit"" title="" 修改 "">修 改</button>"
			end if
		case 3
			if url="" then url="此操作将无法恢复,您确定要删除该信息吗?"
			if checkPowered(Trim(Session("AdminPower")),id,3)=true then
			Response.Write "<button name=""btnDelete"" type=""submit"" title=""删除"" onClick='return confirm("""&url&""")'>删 除</button>"
			end if
        case else
			if checkPowered(Trim(Session("AdminPower")),id,1)=true then
			Response.Write"<button type=""submit"" name=""btnSubmit"" id=""btnSubmit"" title="" 添加 "">添 加</button>"
			end if
    end select

end sub
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

function fileBait(id)
	Set rs = Server.CreateObject("Adodb.Recordset")
	  rs.Open("Select * From [ChannelInfo] Where [ID]="&ID),conn,1,1
	  if not (rs.bof or rs.eof) then
	  	fileBait=RS("FieldBit")
	  end if
	  rs.close
	  set rs=nothing
end function

'返回相应模型的自定义字段名称数组
		   Function Get_KS_D_F_P_Arr(ChannelID,Param)
		      Dim KS_RS_Obj:Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
			   KS_RS_Obj.Open "Select FieldName,Title,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,Height,FieldID,ShowOnUserForm,EditorType From UserField Where ChannelID=" & ChannelID &" And ShowOnForm=1 " & Param & " Order By OrderID Asc",Conn,1,1
			 If Not KS_RS_Obj.Eof Then
			  Get_KS_D_F_P_Arr=KS_RS_Obj.GetRows(-1)
			 Else
			  Get_KS_D_F_P_Arr=""
			 End If
			 KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
		   End Function
			'返回相应模型的自定义字段名称数组
		   Function Get_KS_D_F_Arr(ChannelID)
			  Get_KS_D_F_Arr=Get_KS_D_F_P_Arr(ChannelID,"")
		   End Function

		   '取得后台信息添加时的自定义字段表单
		   Function Get_KS_D_F_I(F_Arr,ChannelID,ByVal UserDefineFieldValueStr,V_Tag)
		      Dim I,K,O_Arr,F_Value
			  Dim O_Text,O_Value,BRStr,O_Len,F_V
                If UserDefineFieldValueStr<>"0" And UserDefineFieldValueStr<>""  Then UserDefineFieldValueStr=Split(UserDefineFieldValueStr,"||||")
              If IsArray(F_Arr) Then
				For I=0 To Ubound(F_Arr,2)
				    If ChannelID=101 and F_Arr(10,I)="0" Then
				    Get_KS_D_F_I=Get_KS_D_F_I & "<tr class='tdbg'[@NoDisplay(" & F_Arr(0,i) & ")]>" & vbcrlf 
					Else
				    Get_KS_D_F_I=Get_KS_D_F_I & "<tr class='tdbg'>" & vbcrlf 
					End If
					Get_KS_D_F_I=Get_KS_D_F_I & " <td width=""85"" align=""right"" class='clefttitle'>" & F_Arr(1,I) & ":</td>" & vbcrlf
					Get_KS_D_F_I=Get_KS_D_F_I & " <td>"
					If IsArray(UserDefineFieldValueStr) Then
					   F_Value=UserDefineFieldValueStr(I)
					 Else
					   F_Value=F_Arr(4,I)
					   If Instr(F_Value,"|")<>0 Then F_Value=""
					 End If
					 
				   If V_Tag=1 Then	 
				    Get_KS_D_F_I=Get_KS_D_F_I & "[@" & F_Arr(0,i) &"]"
                   ElseIf lcase(F_Arr(0,i))="province&city" Then
				   	Get_KS_D_F_I=Get_KS_D_F_I & "<script language=""javascript"" src=""" & KS.Setting(2) & "/Plus/Area.asp""></script>"
				   Else
					   Select Case F_Arr(3,I)
						 Case 2
						   Get_KS_D_F_I=Get_KS_D_F_I & "<textarea style=""width:" & F_Arr(7,i) & ";height:" & F_Arr(8,i) & "px"" rows=""5"" class=""upfile"" name=""" & F_Arr(0,i) & """>" & F_Value & "</textarea>"
						 Case 3
						  Get_KS_D_F_I=Get_KS_D_F_I & "<select class=""upfile"" style=""width:" & F_Arr(7,i) & """ name=""" & F_Arr(0,I) & """>"
						   O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
						   For K=0 To O_Len
						    If O_Arr(K)<>"" Then
							   F_V=Split(O_Arr(K),"|")
							   If Ubound(F_V)=1 Then
								O_Value=F_V(0):O_Text=F_V(1)
							   Else
								O_Value=F_V(0):O_Text=F_V(0)
							   End If						   
							 If F_Value=O_Value Then
							  Get_KS_D_F_I=Get_KS_D_F_I & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
							 Else
							  Get_KS_D_F_I=Get_KS_D_F_I & "<option value=""" & O_Value& """>" & O_Text & "</option>"
							 End If
							End If
						   Next
						  Get_KS_D_F_I=Get_KS_D_F_I & "</select>"
						 Case 6
						   O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
						   If O_Len>1 And Len(F_Arr(5,I))>50 Then BrStr="<br>" Else BrStr=""
						   For K=0 To O_Len
							   F_V=Split(O_Arr(K),"|")
							   If O_Arr(K)<>"" Then
							   If Ubound(F_V)=1 Then
								O_Value=F_V(0):O_Text=F_V(1)
							   Else
								O_Value=F_V(0):O_Text=F_V(0)
							   End If						   
							 If F_Value=O_Value Then
							  Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""radio"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
							 Else
							  Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""radio"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """>" & O_Text & BRStr
							 End If
							End If
						   Next
						 Case 7
						   O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
						   For K=0 To O_Len
						     If O_Arr(K)<>"" Then
							   F_V=Split(O_Arr(K),"|")
							   If Ubound(F_V)=1 Then
								O_Value=F_V(0):O_Text=F_V(1)
							   Else
								O_Value=F_V(0):O_Text=F_V(0)
							   End If						   
							 If KS.FoundInArr(F_Value,O_Value,",")=true Then
							  Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""checkbox"" class=""checkbox"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """ checked>" & O_Text
							 Else
							  Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""checkbox"" class=""checkbox"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """>" & O_Text
							 End If
							End If
						   Next
						 Case 10
							Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""hidden"" id=""" & F_Arr(0,I) &""" name=""" & F_Arr(0,I) &""" value="""& Server.HTMLEncode(F_Value) &""" style=""display:none"" /><input type=""hidden"" id=""" & F_Arr(0,I) &"___Config"" value="""" style=""display:none"" /><iframe id=""" & F_Arr(0,I) &"___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=" & F_Arr(0,I) &"&amp;Toolbar=" & F_Arr(11,i) & """ width=""" & F_Arr(7,i) &""" height=""" & F_Arr(8,i) & """ frameborder=""0"" scrolling=""no""></iframe>"
	
						 Case Else
						   Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""text"" class=""upfile"" style=""width:" & F_Arr(7,i) & """ name=""" & F_Arr(0,i) & """ id=""" & F_Arr(0,i) & """ value=""" & F_Value & """>"
					   End Select
				   End If
				   if F_Arr(3,I)=9 and V_Tag<>1 Then Get_KS_D_F_I=Get_KS_D_F_I & " <input class=""button""  type='button' name='Submit' value='选择...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & KS.GetUpFilesDir() & "',550,290,window,$('#" & F_Arr(0,I) & "')[0]);"">"
				   If  F_Arr(2,I)<>"" Then Get_KS_D_F_I=Get_KS_D_F_I & " <span style=""margin-top:5px"">" &  F_Arr(2,I) & "</span>"
				   if F_Arr(3,I)=9 and V_Tag<>1 Then Get_KS_D_F_I=Get_KS_D_F_I & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='UpFileForm.asp?UPType=Field&FieldID=" & F_Arr(9,I) & "&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				   Get_KS_D_F_I=Get_KS_D_F_I &" </td>" &vbcrlf
				   Get_KS_D_F_I=Get_KS_D_F_I & "</tr>" &vbcrlf
				Next
			End If
		   End Function


		   '取得后台信息添加时的自定义字段
		   Function Get_KS_D_F(ChannelID,ByVal UserDefineFieldValueStr)
		      Dim F_Arr:F_Arr=Get_KS_D_F_Arr(ChannelID)
			  Get_KS_D_F=Get_KS_D_F_I(F_Arr,ChannelID,UserDefineFieldValueStr,0)
		   End Function
		   
		   '根据sql 参数取表单
		   Function Get_KS_D_F_P(ChannelID,ByVal UserDefineFieldValueStr,Param)
		      Dim F_Arr:F_Arr=Get_KS_D_F_P_Arr(ChannelID,Param)
			  Get_KS_D_F_P=Get_KS_D_F_I(F_Arr,ChannelID,UserDefineFieldValueStr,1)
		   End Function
		   
		   '**************************************************
	'函数名：ShowError
	'作  用：显示错误信息。
	'参  数：Errmsg  ----出错信息
	'返回值：无
	'**************************************************
	Sub ShowError(Errmsg)
		Response.Write("<br><br><div align=""center"">")
		Response.Write ("  <center>")
		Response.Write ("  <table border=""0"" cellpadding='2' cellspacing='1' width=""75%"" style=""MARGIN-TOP: 3px"" class='table'>")
		Response.Write ("	 <tr class='back'>")
		Response.Write ("			  <td width=""100%"" height=""30"" align=""left"" class='xingmu'>")
		Response.Write ("				&nbsp;系统提示：")
		Response.Write ("			  </td>")
		Response.Write ("	 </tr>")
		Response.Write ("	 <tr>")
		Response.Write ("			  <td width=""100%"" height=""30"" align=""center"">")
		Response.Write ("				<b> " & Errmsg & "&nbsp; </b>")
		Response.Write ("				</b>")
		Response.Write ("			  </td>")
		Response.Write ("	 </tr>")
		Response.Write ("	 <tr>")
		Response.Write ("			  <td width=""100%"" height=""30"" align=""center"">")
		Response.Write ("				<p><b>...::: 点击左边上传目录的文件夹名继续 或 <a href=""javascript:history.go(-1)""> 点此返回 </a>:::...</b>")
		Response.Write ("			  </td>")
		Response.Write ("			</tr>")		
		Response.Write ("  </table>")
		Response.Write ("  </center>")
		Response.Write ("</div>")
		c2  ("")
    end sub
	'**************************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'**************************************************
Function JoinChar(ByVal strUrl)
    If strUrl = "" Then
        JoinChar = ""
        Exit Function
    End If
    If InStr(strUrl, "?") < Len(strUrl) Then
        If InStr(strUrl, "?") > 1 Then
            If InStr(strUrl, "&") < Len(strUrl) Then
                JoinChar = strUrl & "&"
            Else
                JoinChar = strUrl
            End If
        Else
            JoinChar = strUrl & "?"
        End If
    Else
        JoinChar = strUrl
    End If
End Function
'**************************************************
'函数名：Refresh1
'作  用：等待特定时间后跳转到指定的网址
'参  数：url ---- 跳转网址
'        refreshTime ---- 等待跳转时间
'**************************************************
Sub Refresh1(url,refreshTime)
        Response.Write "<a Name='rsfreshurl' ID='rsfreshurl' href='"& url &"'></a>" & vbCrLf
        Response.Write "<script language=""javascript""> " & vbCrLf
        Response.Write "  function nextpage(){" & vbCrLf
        Response.Write "    var url = document.getElementById('rsfreshurl');" & vbCrLf
        Response.Write "    if (document.all) {" & vbCrLf
        Response.Write "      url.click();" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "   else if (document.createEvent) {" & vbCrLf
        Response.Write "     var ev = document.createEvent('HTMLEvents');" & vbCrLf
        Response.Write "       ev.initEvent('click', false, true);" & vbCrLf
        Response.Write "       url.dispatchEvent(ev);" & vbCrLf
        Response.Write "    }" & vbCrLf
        Response.Write "  }" & vbCrLf
        Response.Write "  setTimeout(""nextpage();"","&refreshTime*1000&");" & vbCrLf
        Response.Write "</script>" & vbCrLf
End Sub
'**************************************************
'过程名：WriteSuccessMsg
'作  用：显示成功提示信息
'参  数：无
'**************************************************
Sub WriteSuccessMsg(sSuccessMsg, sComeUrl)
    Response.Write "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf
    Response.Write "<link href=""images/skin/Css_"&Session("Admin_Style_Num")&"/"&Session("Admin_Style_Num")&".css"" rel=""stylesheet"" type=""text/css""></head><body><br><br>" & vbCrLf
    Response.Write "<table cellpadding=3 cellspacing=1 border=0 width=400 class='table' align=center>" & vbCrLf
    Response.Write "  <tr align='center' class='back'><td height='22' class='xingmu'><strong>恭喜你！</strong></td></tr>" & vbCrLf
    Response.Write "  <tr class='tdbg'><td height='100' valign='top'><br>" & sSuccessMsg & "</td></tr>" & vbCrLf
    Response.Write "  <tr align='center' class='tdbg'><td>"
    If sComeUrl <> "" Then
        Response.Write "<a href='" & sComeUrl & "'>&lt;&lt; 返回上一页</a>"
    Else
        Response.Write "<a href='javascript:window.close();'>【关闭】</a>"
    End If
    Response.Write "</td></tr>" & vbCrLf
    Response.Write "</table>" & vbCrLf
    Response.Write "</body></html>" & vbCrLf
End Sub

'取得后台公共管理部分的上传目录,一般用于广告,公告设置等
	Function GetCommonUpFilesDir()
	  Dim Str
	  If C("SuperTF")="1" Then 
	    Str=pic_path & Setting(91)
	  Else
	    Str=GetUpFilesDir()
	  End If
	  If Right(Str,1)="/" Then Str=Left(Str,Len(Str)-1)
	  GetCommonUpFilesDir=Str
	End Function
	
	
	Public Sub GetSetting()
		    Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		    RSObj.Open "SELECT top 1 Setting,TbSetting,SpaceSetting,JobSetting,AskSetting from [T_Config] where lan='"&LanID&"'",conn,1,1
		    Dim i,node,xml,j,DataArray,rs
			Set xml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			xml.appendChild(xml.createElement("xml"))
			If Not RSObj.EOF Then
						DataArray=RSObj.GetRows(1)
						For i=0 To UBound(DataArray,2)
							Set Node=xml.createNode(1,"config","")
							j=0
							For Each rs in RSObj.Fields
								node.attributes.setNamedItem(xml.createNode(2,LCase(rs.name),"")).text= Replace(DataArray(j,i),vbcrlf,"$br$")& ""
								j=j+1
							Next
							xml.documentElement.appendChild(Node)
						Next
			End If
			DataArray=Null
		   Set Application(SiteSN&"_Config")=Xml
		   RSObj.Close:Set RSObj=Nothing
	 End Sub

	 Sub InitialConfig()
		GetSetting
		Setting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@setting").text,"$br$",vbcrlf),"^%^")
		TbSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@tbsetting").text,"$br$",vbcrlf),"^%^")
        SSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@spacesetting").text,"$br$",vbcrlf),"^%^")
		JSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@jobsetting").text,"$br$",vbcrlf),"^%^")
		ASetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@asksetting").text,"$br$",vbcrlf),"^%^")
	 End Sub
	 
	 '取上传目录按日期存放
	Function GetUpFilesDir()
	   Dim DateFolder:DateFolder=pic_path & Setting(91) & Year(Now) & "-" & Right("0"&Month(Now),2)
	   CreateListFolder(DateFolder)
	   GetUpFilesDir=DateFolder
	End Function
	
	'**************************************************
	'函数名：CreateListFolder
	'作  用：不限分级创建目录 形如 1\2\3\ 则在网站根目录下创建分级目录
	'参  数：Folder要创建的目录
	'返回值：成功返回true 否则返回Flase
	'**************************************************
	Public Function CreateListFolder(ByVal Folder)
		Dim FSO, WaitCreateFolder, SplitFolder, CF, k
		 'on error resume next
		If Folder = "" Then
		 CreateListFolder = False:Exit Function
		End If
	   Folder = Replace(Folder, "\", "/")
	   If Right(Folder, 1) <> "/" Then Folder = Folder & "/"
	   If Left(Folder, 1) <> "/" Then Folder = "/" & Folder

		 Set FSO = CreateObject(Setting(99))
		 If Not FSO.FolderExists(Server.MapPath(Folder)) Then
		   SplitFolder = Split(Folder, "/")
		 For k = 0 To UBound(SplitFolder) - 1
		  If k = 0 Then
		   CF = SplitFolder(k) & "/"
		  Else
		   CF = CF & SplitFolder(k) & "/"
		  End If
		  If (Not FSO.FolderExists(Server.MapPath(CF))) Then
			 FSO.CreateFolder (Server.MapPath(CF))
			 CreateListFolder = True
		  End If
		 Next
	   End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear
	   CreateListFolder = False
	   Else
	   CreateListFolder = True
	   End If
	 End Function
	
	 '**************************************************
	'函数名：DeleteFolder
	'作  用：删除指定目录
	'参  数：FolderStr要删除的目录
	'返回值：成功返回true 否则返回Flase
	'**************************************************
	Public Function DeleteFolder(FolderStr)
	   Dim FSO
	   'on error resume next
	   FolderStr = Replace(FolderStr, "\", "/")
	   Set FSO = CreateObject(Setting(99))
		If FSO.FolderExists(Server.MapPath(FolderStr)) Then
			  FSO.DeleteFolder (Server.MapPath(FolderStr))
		Else
		DeleteFolder = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFolder = False
	   Else
	   DeleteFolder = True
	   End If
	End Function
	 '**************************************************
	'函数名：DeleteFile
	'作  用：删除指定文件
	'参  数：FileStr要删除的文件
	'返回值：成功返回true 否则返回Flase
	'**************************************************
	Public Function DeleteFile(FileStr)
	   Dim FSO
	   'on error resume next
	   Set FSO = CreateObject(Setting(99))
		If FSO.FileExists(Server.MapPath(FileStr)) Then
			FSO.DeleteFile Server.MapPath(FileStr), True
		Else
		DeleteFile = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFile = False
	   Else
	   DeleteFile = True
	   End If
	End Function
	'水印组件列表
	Function ExpiredStr(I)
		   Dim ComponentName(3)
			ComponentName(0) = "Persits.Jpeg"
			ComponentName(1) = "wsImage.Resize"
			ComponentName(2) = "SoftArtisans.ImageGen"
			ComponentName(3) = "CreatePreviewImage.cGvbox"
			If IsObjInstalled(ComponentName(I)) Then
				If IsExpired(ComponentName(I)) Then
					ExpiredStr = "，但已过期"
				Else
					ExpiredStr = ""
				End If
			  ExpiredStr = " √支持" & ExpiredStr
			Else
			  ExpiredStr = "×不支持"
			End If
	End Function
	
	Function IsObjInstalled(strClassString)
		'on error resume next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function
	Public Function IsExpired(strClassString)
		'on error resume next
		IsExpired = True
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
	
		If 0 = Err Then
			Select Case strClassString
				Case "Persits.Jpeg"
					If xTestObjResponse.Expires > Now Then
						IsExpired = False
					End If
				Case "wsImage.Resize"
					If InStr(xTestObj.errorinfo, "已经过期") = 0 Then
						IsExpired = False
					End If
				Case "SoftArtisans.ImageGen"
					xTestObj.CreateImage 500, 500, RGB(255, 255, 255)
					If Err = 0 Then
						IsExpired = False
					End If
			End Select
		End If
		Set xTestObj = Nothing
		Err = 0
	End Function
	
	'**************************************************
	'函数名：GetAutoDoMain()
	'作  用：取得当前服务器IP 如：http://127.0.0.1
	'参  数：无
	'**************************************************
	Public Function GetAutoDomain()
		Dim TempPath
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			GetAutoDomain = Request.ServerVariables("SERVER_NAME")
		Else
			GetAutoDomain = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
		 If Instr(UCASE(GetAutoDomain),"/W3SVC")<>0 Then
			   GetAutoDomain=Left(GetAutoDomain,Instr(GetAutoDomain,"/W3SVC"))
		 End If
		 GetAutoDomain = "http://" & GetAutoDomain
	End Function
	
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove(MyCaheName)
		Application.unLock
	End Sub
	Function ChkClng(ByVal str)
	    'on error resume next
		If IsNumeric(str) Then
			ChkClng = CLng(str)
		Else
			ChkClng = 0
		End If
		If Err Then ChkClng=0
	End Function
 	 Function InitialObject(str)
		'iis5创建对象方法Server.CreateObject(ObjectName);
		'iis6创建对象方法CreateObject(ObjectName);
		'默认为iis6，如果在iis5中使用，需要改为Server.CreateObject(str);
		Set InitialObject=CreateObject(str)
	 End Function
	 '==================================================
		'函数名：GetHttpPage
		'作  用：获取网页源码
		'参  数：HttpUrl ------网页地址
		'==================================================
		Function GetHttpPage(HttpUrl,CharsetCode)
		   If IsNull(HttpUrl) = True Or Len(HttpUrl) < 18 Or HttpUrl = "Error" Then
			  GetHttpPage = "Error"
			  Exit Function
		   End If
		   Dim Http
		  Set Http = InitialObject("MSXML2.ServerXMLHTTP") 
		   Http.Open "GET", HttpUrl, False
		  'on error resume next
		   Http.Send
		   If Http.Readystate <> 4 Then
		  'If Http.Status<>200 then 
			  Set Http = Nothing
			  GetHttpPage = "Error"
			  Exit Function
		   End If
		   if CharsetCode="auto" Or CharsetCode="" then CharsetCode=GetEncodeing(HttpUrl)
		   GetHttpPage = BytesToBstr(Http.ResponseBody, CharsetCode)
		   Set Http = Nothing
		   If Err.Number <> 0 Then
			  Err.Clear
		   End If
		End Function
		'自动取得编码格式
		function GetEncodeing(sUrl)
		'on error resume next
		dim http,re,encodeing
		Set http=InitialObject("Microsoft.XMLHTTP")
		http.Open "GET",sUrl,False
		http.send
		if http.status="200" then
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="encoding=\""utf-8"
			if re.test(http.responseText) then
				encodeing="utf-8"
			else
				encodeing="utf-8"
			end if
			set re=nothing
		end if
		If Err Then
			Err.Clear
			GetEncodeing="utf-8"
		else
			GetEncodeing=encodeing
		End If
		set http=nothing
	end function
		'==================================================
		'函数名：BytesToBstr
		'作  用：将获取的源码转换为中文
		'参  数：Body ------要转换的变量
		'参  数：Cset ------要转换的类型
		'==================================================
		Function BytesToBstr(Body, Cset)
		   Dim Objstream
		   Set Objstream = InitialObject("adodb.stream")
		   Objstream.Type = 1
		   Objstream.Mode = 3
		   Objstream.Open
		   Objstream.Write Body
		   Objstream.Position = 0
		   Objstream.Type = 2
		   Objstream.Charset = Cset
		   BytesToBstr = Objstream.ReadText
		   Objstream.Close
		   Set Objstream = Nothing
		End Function
		
		'**************************************************
	'函数名：WriteTOFile
	'作  用：写内容到指定的html文件
	'参  数：Filename  ----目标文件件 如 mb\index.htm
	'        Content   ------要写入目标文件的内容
	'返回值：成功返回true ,失败返回false
	'**************************************************
	Public Function WriteTOFile(FileName, Content)
	    'on error resume next
		dim stm:set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '以文本模式读取
		stm.mode=3
		stm.charset="utf-8"
		stm.open
		stm.WriteText content
		stm.SaveToFile server.MapPath(FileName),2 
		stm.flush
		stm.Close
		set stm=nothing
	  
	   If Err.Number <> 0 Then
		 WriteTOFile = False
	   Else
		 WriteTOFile = True
	   End If
	End Function
	'**************************************************
	'函数名：ReadFromFile
	'作  用：写内容到指定的html文件
	'参  数：Filename  ----目标文件件 如 mb\index.htm
	'返回值：成功返回文件内容 ,失败返回""
	'**************************************************
	Public Function ReadFromFile(FileName)
		 'on error resume next
		 dim str,stm
		set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '以本模式读取
		stm.mode=3 
		stm.charset="UTF-8"
		stm.open
		stm.loadfromfile server.MapPath(FileName)
		str=stm.readtext
		stm.Close
		set stm=nothing
		ReadFromFile=str
	End Function
	
	'取得后台公共管理部分的上传目录,一般用于广告,公告设置等
	Function GetCommonUpFilesDir()
	  Dim Str
	  If C("SuperTF")="1" Then 
	    Str=pic_path & Setting(91)
	  Else
	    Str=GetUpFilesDir()
	  End If
	  If Right(Str,1)="/" Then Str=Left(Str,Len(Str)-1)
	  GetCommonUpFilesDir=Str
	End Function
	
'**************************************************
'函数名：Array2Option
'作  用：把数组变成下拉列表项目
'参  数：arrTemp ---- 数组
'        ID ---- 数组中默认的数字
'返回值：下拉菜单
'**************************************************
Public Function Array2Option(ByVal arrTemp, ByVal ID)
    Dim strOption, i, arrValue
    strOption = "<option value=''> </option>"
    For i = 0 To UBound(arrTemp)
        arrValue = Split(arrTemp(i), "|")
        If CLng(arrValue(1)) = 1 Then
			ValueArr=split(arrValue(0),"@")
			if UBound(ValueArr)>0 then
				Color="style='COLOR:"&ValueArr(1)&"'"
				arrValues=ValueArr(0)
			else
				arrValues=arrValue(0)
			end if 
            If ID <>"" Then
                If arrValues = ID Then
                    strOption = strOption & "<option value='" & arrValues & "' "&Color&" selected>" & arrValues & "</option>"
                Else
                    strOption = strOption & "<option value='" & arrValues & "' "&Color&">" & arrValues & "</option>"
                End If
            Else
                If CLng(arrValue(2)) = 1 Then
                    strOption = strOption & "<option value='" & arrValues & "' "&Color&" selected>" & arrValues & "</option>"
                Else
                    strOption = strOption & "<option value='" & arrValues & "' "&Color&">" & arrValues & "</option>"
                End If
            End If
        End If
    Next
    Array2Option = strOption
End Function

'**************************************************
'函数名：Array2OBox
'作  用：把数组变成多项框
'参  数：arrTemp ---- 数组
'        ID ---- 数组中默认的数字
'		values----多项框id名称
'返回值：下拉菜单
'**************************************************
Public Function Array2OBox(ByVal arrTemp, ByVal ID,ByVal values)
    Dim strOption, i, arrValue
    strOption = "<option value=''> </option>"
    For i = 0 To UBound(arrTemp)
        arrValue = Split(arrTemp(i), "|")
        If CLng(arrValue(1)) = 1 Then
			arrValues=arrValue(0)
            If ID <>"" Then
                If FoundInArr(ID, arrValues, ",") Then
                    strOption = strOption & "<INPUT id='"&values&"' type=checkbox class=checkbox CHECKED value='" & arrValues & "' name='"&values&"'> " & arrValues & ""
                Else
                    strOption = strOption & "<INPUT id='"&values&"' type=checkbox class=checkbox value='" & arrValues & "' name='"&values&"'> " & arrValues & ""
                End If
            Else
                If CLng(arrValue(2)) = 1 Then
                    strOption = strOption & "<INPUT id='"&values&"' type=checkbox class=checkbox CHECKED value='" & arrValues & "' name='"&values&"'> " & arrValues & ""
                Else
                    strOption = strOption & "<INPUT id='"&values&"' type=checkbox class=checkbox value='" & arrValues & "' name='"&values&"'> " & arrValues & ""
                End If
            End If
        End If
    Next
    Array2OBox = strOption
End Function
'**************************************************
'函数名：Array2href
'作  用：把数组变成连接选择条目
'参  数：arrTemp ---- 数组
'       values-----要显示的文本框
'返回值：下拉菜单
'**************************************************
Public Function Array2href(ByVal arrTemp,ByVal values)
    Dim strOption, i, arrValue
    strOption = "<option value='-1'> </option>"
    ID = chkclng(ID)
    For i = 0 To UBound(arrTemp)
        arrValue = Split(arrTemp(i), "|")
        If CLng(arrValue(1)) = 1 Then
			ValueArr=split(arrValue(0),"@")
			if UBound(ValueArr)>0 then
				Color="style='COLOR:"&ValueArr(1)&"'"
				arrValues=ValueArr(0)
			else
				arrValues=arrValue(0)
			end if 
			strOption = strOption & "【<font color='blue'><font color='#993300' onclick=""$('#"&values&"').val('" & arrValues & "')"" style='cursor:pointer;'>" & arrValues & "</font></font>】"
        End If
    Next
    Array2href = strOption
End Function
'**************************************************
'函数名：GetArrFromDictionary
'作  用：从字典表获得区域值
'参  数：Fieldid ---- id
'返回值：查询区域值
'**************************************************
Public Function GetArrFromDictionary(Fieldid)
    Dim rsDictionary, strTemp
	if chkclng(Fieldid)=0 then exit Function
    Set rsDictionary = Conn.Execute("select FieldValue from CB_Dictionary where Fieldid=" & Fieldid & "")
    If rsDictionary.BOF And rsDictionary.EOF Then
        strTemp = ""
    Else
        strTemp = rsDictionary(0) & ""
    End If
    Set rsDictionary = Nothing
    GetArrFromDictionary = Split(strTemp, "$$$")
End Function

'取得会员组选项--多选列表 参数：SelectArr--默认选择项以","隔开,RowNum--每行显示选项数
	Public Function GetUserGroup_CheckBox(OptionName,SelectArr,RowNum)
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
			 GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<input type=""checkbox"" class=checkbox checked name=""" & OptionName & """ value=""" & RSObj(0) & """>" & RSObj(1) & "&nbsp;&nbsp;"
			Else
			 GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<input type=""checkbox"" class=checkbox name=""" & OptionName & """ value=""" & RSObj(0) & """>" & RSObj(1) & "&nbsp;&nbsp;"
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
	
	'==================================================
	'函数名：ScriptHtml
	'作  用：过滤html标记
	'参  数：ConStr ------ 要过滤的字符串
	'==================================================
	Function ScriptHtml(ByVal Constr, TagName, FType)
			Dim re
			Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			Select Case FType
			Case 1
			   re.Pattern = "<" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			Case 2
			   re.Pattern = "<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			Case 3
			   re.Pattern = "<" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			   re.Pattern = "</" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			End Select
			ScriptHtml = Constr
			Set re = Nothing
	End Function
	
	'取消HTML
		Public Function LoseHtml(ByVal ContentStr)
		    'on error resume next
			Dim TempLoseStr, regEx
			If ContentStr="" Or ContentStr=Null Then Exit Function
			TempLoseStr = HtmlCode(ContentStr)
			Set regEx = New RegExp
			regEx.Pattern = "<\/*[^<>]*>"
			regEx.IgnoreCase = True
			regEx.Global = True
			TempLoseStr = regEx.Replace(TempLoseStr, "")
			LoseHtml = TempLoseStr
		End Function
		
	Sub AddKeyTags(KeyWords)
		     dim i
			 dim trs:set trs=server.createobject("adodb.recordset")
			 dim karr:karr=split(KeyWords,",")
			 for i=0 to ubound(karr)
			 trs.open "select * from CB_keywords where keytext='" & left(karr(i),100) & "'",conn,1,3
			 if trs.eof then
			   trs.addnew
			   trs("keytext")=left(karr(i),100)
			   trs("adddate")=now
			  trs.update
		   end if
			  trs.close
		  next
		   set trs=nothing
	End Sub
	'文章自动分页
	'参数：Content-文章内容 SplitPageStr-文章分隔线 PerPageLen-每页大约字符数
	Function AutoSplitPage(Content,SplitPageStr,maxPagesize)
	    Dim sContent,ss,i,IsCount,c,iCount,strTemp,Temp_String,Temp_Array
		sContent=Content
		If maxPagesize<100 Or Len(sContent)<maxPagesize+100 Then
			AutoSplitPage=sContent
		End If
		sContent=Replace(sContent, SplitPageStr, "")
		sContent=Replace(sContent, "&nbsp;", "<&nbsp;>")
		sContent=Replace(sContent, "&gt;", "<&gt;>")
		sContent=Replace(sContent, "&lt;", "<&lt;>")
		sContent=Replace(sContent, "&quot;", "<&quot;>")
		sContent=Replace(sContent, "&#39;", "<&#39;>")
		If sContent<>"" and maxPagesize<>0 and InStr(1,sContent,SplitPageStr)=0 then
			IsCount=True:Temp_String=""
			For i= 1 To Len(sContent)
				c=Mid(sContent,i,1)
				If c="<" Then
					IsCount=False
				ElseIf c=">" Then
					IsCount=True
				Else
					If IsCount=True Then
						'If Abs(Asc(c))>255 Then
						'	iCount=iCount+2
						'Else
							iCount=iCount+1
						'End If
						If iCount>=maxPagesize And i<Len(sContent) Then
							strTemp=Left(sContent,i)
							If CheckPagination(strTemp,"table|a|b>|i>|strong|div|span") then
								Temp_String=Temp_String & Trim(CStr(i)) & "," 
								iCount=0
							End If
						End If
					End If
				End If	
			Next
			If Len(Temp_String)>1 Then Temp_String=Left(Temp_String,Len(Temp_String)-1)
			Temp_Array=Split(Temp_String,",")
			For i = UBound(Temp_Array) To LBound(Temp_Array) Step -1
				ss = Mid(sContent,Temp_Array(i)+1)
				If Len(ss) > 100 Then
					sContent=Left(sContent,Temp_Array(i)) & SplitPageStr & ss
				Else
					sContent=Left(sContent,Temp_Array(i)) & ss
				End If
			Next
		End If
		sContent=Replace(sContent, "<&nbsp;>", "&nbsp;")
		sContent=Replace(sContent, "<&gt;>", "&gt;")
		sContent=Replace(sContent, "<&lt;>", "&lt;")
		sContent=Replace(sContent, "<&quot;>", "&quot;")
		sContent=Replace(sContent, "<&#39;>", "&#39;")
		AutoSplitPage=sContent
	End Function
    '结合以上函数使用
	Private Function CheckPagination(strTemp,strFind)
		Dim i,n,m_ingBeginNum,m_intEndNum
		Dim m_strBegin,m_strEnd,FindArray
		strTemp=LCase(strTemp)
		strFind=LCase(strFind)
		If strTemp<>"" and strFind<>"" then
			FindArray=split(strFind,"|")
			For i = 0 to Ubound(FindArray)
				m_strBegin="<"&FindArray(i)
				m_strEnd  ="</"&FindArray(i)
				n=0
				do while instr(n+1,strTemp,m_strBegin)<>0
					n=instr(n+1,strTemp,m_strBegin)
					m_ingBeginNum=m_ingBeginNum+1
				Loop
				n=0
				do while instr(n+1,strTemp,m_strEnd)<>0
					n=instr(n+1,strTemp,m_strEnd)
					m_intEndNum=m_intEndNum+1
				Loop
				If m_intEndNum=m_ingBeginNum then
					CheckPagination=True
				Else
					CheckPagination=False
					Exit Function
				End If
			Next
		Else
			CheckPagination=False
		End If
	End Function
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) then
		    fString = ClearBadChr(fString)
			fString = Replace(fString, "&", "&amp;")
			fString = Replace(fString, "'", "&#39;")
			fString = Replace(fString, ">", "&gt;")
			fString = Replace(fString, "<", "&lt;")
			fString = Replace(fString, Chr(32), " ")
			fString = Replace(fString, Chr(9), " ")
			fString = Replace(fString, Chr(34), "&quot;")
			fString = Replace(fString, Chr(39), "&#39;")
			fString = Replace(fString, Chr(13), "")
			'fString = Replace(fString, " ", "&nbsp;")
			'fString = Replace(fString, Chr(10), "<br />")
		HTMLEncode = fString
		End If
	End Function
	
	Function ClearBadChr(str)
	  If Str<>"" Then
	     Dim re:Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(on(load|click|dbclick|mouseover|mouseout|mousedown|mouseup|mousewheel|keydown|submit|change|focus)=""[^""]+"")"
		str = re.Replace(str, "")
		re.Pattern="((name|id|class)=""[^""]+"")"
		str = re.Replace(str, "")
		re.Pattern = "(<s+cript[^>]*?>([\w\W]*?)<\/s+cript>)"
		str = re.Replace(str, "")
		re.Pattern = "(<iframe[^>]*?>([\w\W]*?)<\/iframe>)"
		str = re.Replace(str, "")
		re.Pattern = "(<p>&nbsp;<\/p>)"
		str = re.Replace(str, "")
		Set re=Nothing
		ClearBadChr = str
	 End If	
	End Function

	
	Public Function HTMLCode(HtmlStr)
		If Not IsNul(HtmlStr) then
		'HtmlStr = Replace(HtmlStr, "&nbsp;", " ")
		HtmlStr = Replace(HtmlStr, "&quot;", Chr(34))
		HtmlStr = Replace(HtmlStr, "&#39;", Chr(39))
		HtmlStr = Replace(HtmlStr, "&#123;", Chr(123))
		HtmlStr = Replace(HtmlStr, "&#125;", Chr(125))
		HtmlStr = Replace(HtmlStr, "&#36;", Chr(36))
		HtmlStr = Replace(HtmlStr, "&amp;", "&")
		'HtmlStr = Replace(HtmlStr, vbCrLf, "")

		HtmlStr = Replace(HtmlStr, "&gt;", ">")
		HtmlStr = Replace(HtmlStr, "&lt;", "<")
		
		HTMLCode = HtmlStr
		End If
	End Function
	'**************************************************
	'函数名：SaveBeyondFile
	'作  用：保存远程文件到本地
	'参  数：LocalFile 本地文件,BFU远程文件
	'返回值：无
	'**************************************************
	Public Function ReplaceBeyondUrl(ReplaceContent, SaveFilePath)
		Dim re, BeyondFile, BFU, SaveFileName,SaveFileList
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp)))"
		Set BeyondFile = re.Execute(ReplaceContent)
		Set re = Nothing
		For Each BFU In BeyondFile
		  If Instr(SaveFileList,BFU)=0 Then
			SaveFileName = Year(Now()) & Month(Now()) & Day(Now()) & MakeRandom(10) & Mid(BFU, InStrRev(BFU, "."))
			If Instr(BFU,Setting(2))<=0 Then
			Call SaveBeyondFile(SaveFilePath&SaveFileName,BFU)
			ReplaceContent = Replace(ReplaceContent, BFU, Setting(2) & SaveFilePath & SaveFileName)
			End If
		  End If
		   SaveFileList=SaveFileList & "," & BFU
		Next
		ReplaceBeyondUrl = ReplaceContent
	End Function

	'==================================================
	'过程名：SaveBeyondFile
	'作  用：保存远程的文件到本地
	'参  数：LocalFileName ------ 本地文件名
	'参  数：RemoteFileUrl ------ 远程文件URL
	'==================================================
	Function SaveBeyondFile(LocalFileName,RemoteFileUrl)
	    'on error resume next
		Dim SaveRemoteFile:SaveRemoteFile=True
		dim Ads,Retrieval,GetRemoteData
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
			.Open "Get", RemoteFileUrl, False, "", ""
			.Send
			If .Readystate<>4 then
				SaveRemoteFile=False
				Exit Function
			End If
			GetRemoteData = .ResponseBody
		End With
		Set Retrieval = Nothing
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile server.MapPath(LocalFileName),2
			.Cancel()
			.Close()
		End With
		Set Ads=nothing
		SaveBeyondFile=SaveRemoteFile
		'加水印
		Dim T:Set T=New Thumb
		call T.AddWaterMark(LocalFileName)
		Set T=Nothing
	end Function
	'*************************************************************************************
	'函数名:GetInfoID
	'作  用:生成文章,图片或下载等的唯一ID
	'参  数:ChannelID--频道ID
	'*************************************************************************************
	Public Function GetInfoID(ChannelID)
	   'on error resume next
	   Dim RSC, TableNameStr
       Set RSC=Server.CreateObject("ADODB.RECORDSET")
	   TableNameStr = "Select ProID From CB_products Where ProID='"
	   Do While True
		 GetInfoID = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Now(), "-", ""), " ", ""), ":", ""), "PM", ""), "AM", ""), "上午", ""), "下午", "") & MakeRandom(3)
			RSC.Open TableNameStr & GetInfoID & "'", Conn, 1, 1
			If RSC.EOF And RSC.BOF Then Exit Do
	   Loop
		RSC.Close:Set RSC = Nothing
	End Function
	
	 '加载用户组缓存
	 Sub LoadUserGroup()
	   If Not IsObject(Application(SiteSN&"_UserGroup")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select id,groupname,powerlist,descript From KS_UserGroup Order by ID")
		 Set Application(SiteSN&"_UserGroup")=RsToxml(rs,"row","groupConfig")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 
	 'xmlroot跟节点名称 row记录行节点名称
	Public Function RsToxml(RSObj,row,xmlroot)
			Dim i,node,rs,j,DataArray
			If xmlroot="" Then xmlroot="xml"
			If row="" Then row="row"
			Set RsToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			RsToxml.appendChild(RsToxml.createElement(xmlroot))
			If Not RSObj.EOF Then
						DataArray=RSObj.GetRows(-1)
						For i=0 To UBound(DataArray,2)
							Set Node=RsToxml.createNode(1,row,"")
							j=0
							For Each rs in RSObj.Fields
								node.attributes.setNamedItem(RsToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
								j=j+1
							Next
							RsToxml.documentElement.appendChild(Node)
						Next
			End If
			DataArray=Null
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
'从数据表添加数据到option选项 参数:表名,字段,查询条件
	  Function Get_O_F_D(Table,FieldStr,Param)
	       Dim KS_RS_Obj,Arr,I
		      If Instr(lcase(FieldStr),"distinct")<=0 and Instr(lcase(FieldStr),"top")<=0 Then FieldStr=" top 50 " &FieldStr
			  Set KS_RS_Obj = conn.Execute("Select " & FieldStr & " FROM "  & Table & " Where " & Param)
			  If Not KS_RS_Obj.Eof Then
			    Arr=KS_RS_Obj.GetRows(-1)
				KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
				For I=0 To Ubound(Arr,2)
					Get_O_F_D = Get_O_F_D & "<option value=""" & Arr(0,i) & """>" & Arr(0,i) & "</option>"
				Next
			   End If
	  End Function
 %>