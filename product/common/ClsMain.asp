<%
 
'GetLng
function GetLong(tstrSource,Min,Max,DefaultValue)
	dim lngRet
	if IsNumeric(tstrSource) = false or tstrSource = "" then
		GetLong = DefaultValue
		exit function
	end if
	
	lngRet = CLng(Left(tstrSource,9))
	if not Min = "" then _
		if lngRet < Min then lngRet = Min
	

	if not Max = "" then _
		if lngRet > Max then lngRet = Max
		
		GetLong = lngRet
end function 


' by cgy, 1999-08-24
'	edit by cgy,19991013, s is null?
' 
' 函数 ValidChar( s ,sType ),传入字符串 s ,
' 返回值:
' sType       返回值
'  0:		 返回有效字符串(只含 A-Z,a-z,0-9,_)
'	 1:		 将其中的 HTML 符号转换
'	 2:		 将回车转换为 <br>
'	 4:		 将回车转换为 <p>
'	 8:		 将单引号转换为两个单引号
'	 3:		 将其中的 HTML 符号转换 并将回车转换为 <br>
'	 5:		 将其中的 HTML 符号转换 并将回车转换为 <p>
'	 9:		 将其中的 HTML 符号转换 并将单引号转换为两个单引号
'	11:		 将其中的 HTML 符号转换 并将单引号转换为两个单引号 将回车转换为 <br>
'	13:		 将其中的 HTML 符号转换 并将单引号转换为两个单引号 将回车转换为 <p>


Function ValidChar(s, sType)
	If IsNull(s) Or IsEmpty(s) Then
		ValidChar = ""
		Exit Function
	End If
	
	Select Case sType
		Case 1:		' 将其中的 HTML 符号转换
			ValidChar = Server.HTMLEncode( s )
			
		Case 2:		' 将回车转换为 <br>
			ValidChar = Replace( s, Chr(13) & Chr(10), "<br>" )
			
		Case 4:		' 将回车转换为 <p>
			ValidChar = Replace( s, Chr(13) & Chr(10), "<p>" )
			
		Case 8:		' 将单引号转换为两个单引号
			ValidChar = Replace( s, "'", "''" )
			
		Case 3:		' 将其中的 HTML 符号转换 并将回车转换为 <br>
			ValidChar = Server.HTMLEncode( s )
			ValidChar = Replace( ValidCHar, Chr(13) & Chr(10), "<br>" )
		
		Case 5:		' 将其中的 HTML 符号转换 并将回车转换为 <p>
			ValidChar = Server.HTMLEncode( s )
			ValidChar = Replace( ValidCHar, Chr(13) & Chr(10), "<p>" )
		
		Case 9:		' 将其中的 HTML 符号转换 并将单引号转换为两个单引号
			ValidChar = Server.HTMLEncode( s )
			ValidChar = Replace( ValidChar, "'", "''" )
		
		Case 11:		' 将其中的 HTML 符号转换 并将单引号转换为两个单引号 将回车转换为 <br>
			ValidChar = Server.HTMLEncode( s )
			ValidChar = Replace( ValidChar, "'", "''" )
			ValidChar = Replace( ValidChar, Chr(13) & Chr(10), "<br>" )
			
		Case 13:		' 将其中的 HTML 符号转换 并将单引号转换为两个单引号 将回车转换为 <p>
			ValidChar = Server.HTMLEncode( s )
			ValidChar = Replace( ValidChar, "'", "''" )
			ValidChar = Replace( ValidChar, Chr(13) & Chr(10), "<p>" )

    Case 0:			' 返回有效字符串(只含 A-Z,a-z,_)
       len_s = Len(s)

       For i = 1 To len_s
         A = Mid(s, i, 1)
         If ((UCase(A) >= "A" And UCase(A) <= "Z") Or (A >= "0" And A <= "9") Or A = "_") Then
					ValidChar = ValidChar & A
			   End If
			 Next
		Case Else:
			ValidChar = s
	End Select
End Function


'UBB格式化
Function HTMLFormat(fString)
   fString=replace(fString,"<","&lt;")
   fString=replace(fString,">","&gt;")
   fString=replace(fString,chr(34),"&quot;")
   fString=replace(fString,chr(13),"<br>")
   fString=replace(fString,chr(32),"&nbsp;")
   fString=replace(fString,chr(124),"&nbsp;&nbsp;")
   fString=replace(fString,chr(9),"&nbsp;&nbsp;&nbsp;&nbsp;")
   fString=trim(fString)
   HTMLFormat=fString
End Function

'清除html标记
Function NoHtml(TestString)
    Dim re
    Set re=new RegExp
    re.IgnoreCase =true
    re.Global=True
    re.Pattern="(\<.[^\<]*\>)"
    TestString=re.replace(TestString,"")
    re.Pattern="(\<\/[^\<]*\>)"
    TestString=re.replace(TestString,"")
    NoHtml=TestString
    Set re=Nothing
End Function

'树型大
Function ListTreeBig(Layer,Islast,Prefix)
	Dim lastsignal,I,Rs	
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL      ="Select CatalogID,CatalogName From eBusiness_Catalog Order by CatalogID"
    Rs.OPEN SQL,cn,1,1
	IF NOT Rs.EOF Then	 
	 While NOT Rs.EOF
		I	=	0		
		Response.Write("<option value=P"&Rs(0)&">")	
		Response.Write(Rs(1)&"</option>"&VBCRLF)		
		Call ListTreeSmall(Rs(0),"eBusiness_Type")
		Rs.MoveNext()
	 Wend
	End IF
	Rs.Close()
	Set Rs = Nothing
End Function

'树型小
Function ListTreeSmall(byval ListClsId,DateTable)
	Dim lastsignal,I,Rs
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL="Select TypeID,TypeName From "&DateTable&" Where CatalogID=" & ListClsId &" Order by TypeID"
    Rs.OPEN SQL,cn,1,1
	IF NOT Rs.EOF Then
	 Prefix = Prefix & "　" 
	 While NOT Rs.EOF
		Response.Write("<option value="&Rs(0)&">")		
		Response.Write(Prefix)
		Response.Write("├")
		Response.Write(Rs(1)&"</option>"&VBCRLF)	
		Rs.MoveNext()
	 Wend
	End IF
	Rs.Close()
	Set Rs = Nothing
End Function

%>