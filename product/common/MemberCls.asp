<%
'****************************************************
' Email: xiaoyxh@163.com . QQ:438888382
' Web: http://www.swcn.net.cn
' Copyright (C) swcn.net.cn All Rights Reserved.
'****************************************************
'-----------------------------------------------------------------------------------------------
'思纬网站管理系统,会员系统函数类
'开发:林文仲 版本 v5.0
'-----------------------------------------------------------------------------------------------
Class UserCls
			Private KS,I  
			'---------定义会员全局变量开始---------------
			Public ID,GroupID,UserName,PassWord,Question,Answer,Email
			Public RealName,Sex,Birthday,IDCard,OfficeTel,HomeTel,Mobile,Fax
			Public Province,City,Address,Zip,HomePage,QQ,ICQ,MSN,UC,UserFace,FaceWidth,FaceHeight,Sign,Privacy,CheckNum,RegDate
			Public JoinDate,LastLoginTime,LastLoginIP,LoginTimes,Money,Score,Point,locked,RndPassword,UserType,SpaceSize
			Public ChargeType,Edays,BeginDate
			'---------定义会员全局变量结束---------------
			
		   '**************************************************
			'函数名：UserLoginChecked
			'作  用：判断用户是否登录
			'返回值：true或false
			'**************************************************
			Public Function UserLoginChecked()
                 'on error resume next
				UserName = R(Trim(C("UserName")))
				PassWord= R(Trim(C("Password")))
				RndPassword=R(Trim(C("RndPassword")))
				
				IF UserName="" Then
				   UserLoginChecked=false
				   Exit Function
				Else
					Dim UserRs:Set UserRS=Server.CreateOBject("ADODB.RECORDSET")
					UserRS.Open "Select a.* From [User] a inner join ks_UserGroup b on a.groupid=b.id Where UserName='" & UserName & "' And PassWord='" & PassWord & "'",Conn,2,3
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
			
			Public Property Get GetEdays()
					GetEdays = Edays-DateDiff("D",BeginDate,now())
			End Property

			
		    Public Function ForValue(s)
					If trim(s)="" Then ForValue=""
					ForValue=s
			End Function
			Sub InnerLocation(msg)
				Response.Write "<script language=""JavaScript"">$('locationid').innerHTML = """ & msg & """;</script>"
			End Sub
			
			'用户上传目录
			Function GetUserFolder(UserName)
			   Dim Ce:Set Ce=new CtoeCls
			   UserName=Ce.CTOE(R(UserName))
			   Set Ce=Nothing
			   GetUserFolder=pic_path&Setting(91)&"User/" & username & "/"
			End Function
			'返回专栏选择框
		  Function UserClassOption(TypeID,Sel)
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select ClassID,ClassName From KS_UserClass Where UserName='" & UserName & "' And TypeID="&TypeID,Conn,1,1
			Do While Not RS.Eof
			  If Sel=RS(0) Then
			  UserClassOption=UserClassOption & "<option value=""" & RS(0) & """ selected>" & RS(1) & "</option>"
			  Else
			  UserClassOption=UserClassOption & "<option value=""" & RS(0) & """>" & RS(1) & "</option>"
			  End iF
			  RS.MoveNext
			Loop
			RS.Close:Set RS=Nothing
		  End Function
			
			'返回相应模型的自定义字段名称数组(仅限会员中心调用)
		   Function KS_D_F_Arr(ChannelID)
		      Dim KS_RS_Obj:Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
			  KS_RS_Obj.Open "Select FieldName,Title,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,Height,FieldID,EditorType From KS_Field Where ChannelID=" & ChannelID &" And ShowOnForm=1 And ShowOnUserForm=1 Order By OrderID Asc",Conn,1,1
			 If Not KS_RS_Obj.Eof Then
			  KS_D_F_Arr=KS_RS_Obj.GetRows(-1)
			 Else
			  KS_D_F_Arr=""
			 End If
			 KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
		   End Function

		   '取得会员中心信息添加时的自定义字段
		   Function KS_D_F(ChannelID,ByVal UserDefineFieldValueStr)
		      Dim I,K,F_Arr,O_Arr,F_Value
			  Dim O_Text,O_Value,BRStr,O_Len,F_V
			    F_Arr=KS_D_F_Arr(ChannelID)
                If UserDefineFieldValueStr<>"0" And UserDefineFieldValueStr<>""  Then UserDefineFieldValueStr=Split(UserDefineFieldValueStr,"||||")
              If IsArray(F_Arr) Then
				For I=0 To Ubound(F_Arr,2)
				    KS_D_F=KS_D_F & "<tr  class=""tdbg"" height=""25""><td align=""center"">" & F_Arr(1,I) & "：</td>"
					KS_D_F=KS_D_F & " <td>"
					If IsArray(UserDefineFieldValueStr) Then
					   F_Value=UserDefineFieldValueStr(I)
					 Else
					   F_Value=F_Arr(4,I)
					  If Instr(F_Value,"|")<>0 Then
					    F_Value=LFCls.GetSingleFieldValue("select " & Split(F_Value,"|")(1) & " from " & Split(F_Value,"|")(0) & " where username='" & UserName & "'") 
					   End If
					 End If

				   Select Case F_Arr(3,I)
				     Case 2
				       KS_D_F=KS_D_F & "　 <textarea style=""width:" & F_Arr(7,i) & ";height:" & F_Arr(8,i) & "px"" rows=""5"" class=""textbox"" name=""" & F_Arr(0,i) & """>" & F_Value & "</textarea>"
					 Case 3
					  KS_D_F=KS_D_F & "　 <select class=""textbox"" style=""width:" & F_Arr(7,i) & """ name=""" & F_Arr(0,I) & """>"
						O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
						For K=0 To O_Len
						   F_V=Split(O_Arr(K),"|")
						   If O_Arr(K)<>"" Then
							   If Ubound(F_V)=1 Then
								O_Value=F_V(0):O_Text=F_V(1)
							   Else
								O_Value=F_V(0):O_Text=F_V(0)
							   End If						   
							 If F_Value=O_Value Then
							  KS_D_F=KS_D_F & "<option value=""" &O_Value& """ selected>" & O_Text & "</option>"
							 Else
							  KS_D_F=KS_D_F & "<option value=""" & O_Value& """>" &O_Text & "</option>"
							 End If
						   End If
					   Next
					  KS_D_F=KS_D_F & "</select>"
					 Case 6
						O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
						For K=0 To O_Len
						   F_V=Split(O_Arr(K),"|")
						  If O_Arr(K)<>"" Then
						   If Ubound(F_V)=1 Then
							O_Value=F_V(0):O_Text=F_V(1)
						   Else
							O_Value=F_V(0):O_Text=F_V(0)
						   End If						   
					     If F_Value=O_Value Then
						  KS_D_F=KS_D_F & "<input type=""radio"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """ checked>" & O_Text
						 Else
						  KS_D_F=KS_D_F & "<input type=""radio"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """>" & O_Text
						 End If
						 End If
					   Next
					 Case 7
						O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
						For K=0 To O_Len
						 F_V=Split(O_Arr(K),"|")
						 If O_Arr(K)<>"" Then
						 If Ubound(F_V)=1 Then
							O_Value=F_V(0):O_Text=F_V(1)
						 Else
							O_Value=F_V(0):O_Text=F_V(0)
						 End If						   
					     If FoundInArr(F_Value,O_Value,",")=true Then
						  KS_D_F=KS_D_F & "<input type=""checkbox"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """ checked>" & O_Text
						 Else
						  KS_D_F=KS_D_F & "<input type=""checkbox"" name=""" & F_Arr(0,I) & """ value=""" &O_Value& """>" & O_Text
						 End If
						 End If
					   Next
					 Case 10
					 	KS_D_F=KS_D_F & "<input type=""hidden"" id=""" & F_Arr(0,I) &""" name=""" & F_Arr(0,I) &""" value="""& Server.HTMLEncode(F_Value) &""" style=""display:none"" /><input type=""hidden"" id=""" & F_Arr(0,I) &"___Config"" value="""" style=""display:none"" />　 <iframe id=""" & F_Arr(0,I) &"___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=" & F_Arr(0,I) &"&amp;Toolbar=" & F_Arr(10,i) & """ width=""" & F_Arr(7,i) &""" height=""" & F_Arr(8,i) & """ frameborder=""0"" scrolling=""no""></iframe>"

					 Case Else
					   KS_D_F=KS_D_F & "　 <input type=""text"" class=""textbox"" style=""width:" & F_Arr(7,i) & """ name=""" & F_Arr(0,i) & """ value=""" & F_Value & """>"
				   End Select
				   If F_Arr(6,I)=1 Then KS_D_F=KS_D_F & "<font color=red> * </font>"
				   KS_D_F=KS_D_F & " <span style=""color:blue;margin-top:5px"">" &  F_Arr(2,I) & "</span>"
				   if F_Arr(3,I)=9 Then KS_D_F=KS_D_F & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & F_Arr(9,I) & "&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				   KS_D_F=KS_D_F & "   </td>"
				   KS_D_F=KS_D_F & "</tr>"
				Next
			End If
		   End Function
		   
		   '返回格式化后的时间
		   Function GetTimeFormat(DateTime)
		      if DateDiff("n",DateTime,now)<5 then
			   GetTimeFormat="刚刚"
			  elseif DateDiff("n",DateTime,now)<60 then
			   GetTimeFormat=DateDiff("n",DateTime,now) & " 分钟前"
			  elseif DateDiff("h",DateTime,now)<5 Then
			   GetTimeFormat=DateDiff("h",DateTime,now) & " 小时前"
			  else
			   GetTimeFormat=formatdatetime(DateTime,2)
			  end if
		   End Function
End Class
%> 
