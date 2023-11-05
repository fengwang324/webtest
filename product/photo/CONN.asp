﻿<%
Response.CodePage=65001
Response.Charset="utf-8"
Dim SqlNowString,DataPart_D,DataPart_Y,DataPart_H
Dim Conn,DBPath,CollectDBPath,DataServer,DataUser,DataBaseName,DataBasePsw,ConnStr,CollcetConnStr
Const DataBaseType=0                   '系统数据库类型，"1"为MS SQL2000数据库，"0"为MS ACCESS 2000数据库
Const SiteSn="cenbel"
Const MsxmlVersion=".3.0"                '系统采用XML版本设置 
dim pic_path,nowpath

'pic_path="/allpic2/product/"


nowpath = Request.ServerVariables("SCRIPT_NAME")
if instr(nowpath,"Admin_UploadFile_Main.asp") >0 then
	pic_shu= instr(nowpath,"/photo/")
	pic_path = left(nowpath,pic_shu)
	response.Cookies("pic_path_cook")=pic_path
	response.cookies("pic_path_cook").Expires="2027-12-30"
else
	pic_path=request.Cookies("pic_path_cook")
end if

'response.write pic_path
'response.end
'pic_path="/allpic2/product2/"


If DataBaseType=0 then
	'如果是ACCESS数据库，请认真修改好下面的数据库的文件名
	DBPath       = DBPaths&"../DB/D#B.asa"      'ACCESS数据库的文件名，请使用相对于网站根目录的的绝对路径
Else
	 '如果是SQL数据库，请认真修改好以下数据库选项
	 DataServer   = "(local)"                                  '数据库服务器IP
	 DataUser     = "sa"                                       '访问数据库用户名
	 DataBaseName = "cenbel"                                '数据库名称
	 DataBasePsw  = ""                                   '访问数据库密码 
End if

 '采集数据库路径
 CollectDBPath="/KS_Collect.Mdb"

'=============================================================== 以下代码请不要自行修改========================================
Call OpenConn
Sub OpenConn()
    'on error resume next
    If DataBaseType = 1 Then
       ConnStr="Provider = Sqloledb; User ID = " & datauser & "; Password = " & databasepsw & "; Initial Catalog = " & databasename & "; Data Source = " & dataserver & ";"
	   SqlNowString = "getdate()"
	   DataPart_D   = "d"
	   DataPart_Y   = "y"
	   DataPart_H   = "hour"
    Else
       ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBPath)
	   SqlNowString = "Now()"
	   DataPart_D   = "'d'"
	   DataPart_Y   = "'yyyy'"
	   DataPart_H   = "'h'"
    End If
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.open ConnStr
    If Err Then Err.Clear:Set conn = Nothing:Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。":Response.End
	CollcetConnStr ="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(CollectDBPath)
End Sub
Sub CloseConn()
    'on error resume next
	Conn.close:Set Conn=nothing
End sub

'==============================================全局临时变量类开始==============================
Dim GCls:Set GCls=New GlobalVarCls
Class GlobalVarCls
    Public StaticPreList,StaticPreContent,StaticExtension
	Private Sub Class_Initialize()
	   StaticPreList    = "list"                 rem 伪静态列表前缀 不能包含"?"及"-"
	   staticPreContent = "thread"              rem 伪静态内容前缀 
	   StaticExtension  = ".html"                rem 伪静态扩展名
	End Sub
    Private Sub Class_Terminate()
		 Set GCls=Nothing
	End Sub
	
	Public Function Execute(Command)
		If Not IsObject(Conn) Then OpenConn()
		'on error resume next
		Set Execute = Conn.Execute(Command)
		If Err Then
				Response.Write("查询语句为：" & Command & "<br>")
				Response.Write("错误信息为：" & Err.Description & "<br>")
			Err.Clear
			Set Execute = Nothing
			Response.End()
		End If
		'Sql_Use = Sql_Use + 1
	End Function
	
	Function GetUrl() 
		'on error resume next 
		Dim strTemp 
		If LCase(Request.ServerVariables("HTTPS")) = "off" Then 
		 strTemp = "http://"
		Else 
		 strTemp = "https://"
		End If 
		strTemp = strTemp & Request.ServerVariables("SERVER_NAME") 
		If Request.ServerVariables("SERVER_PORT") <> 80 Then 
		 strTemp = strTemp & ":" & Request.ServerVariables("SERVER_PORT") 
		end if
		strTemp = strTemp & Request.ServerVariables("URL") 
		If Trim(Request.QueryString) <> "" Then 
		 strTemp = strTemp & "?" & Trim(Request.QueryString) 
		end if
		GetUrl = strTemp 
	End Function

	'====================标志来访地址================
	Public Property Let ComeUrl(ByVal strVar) 
			Session("M_ComeUrl") = strVar 
	End Property 
			
	Public Property Get ComeUrl
			ComeUrl= Session("M_ComeUrl")
	End Property 
	'================================================
	
End Class
'==============================================全局临时变量类结束==============================
sub c2(str)
Response.Write(str)
Response.end
end sub
Function KillSQLCode(s)
	KillSQLCode = LCase(s)
	if Trim(KillSQLCode)<>"" then 
	    KillSQLCode = Replace(KillSQLCode,"'","")     
        KillSQLCode = Replace(KillSQLCode,"or","")
		KillSQLCode = Replace(KillSQLCode,"=","")     
        KillSQLCode = Replace(KillSQLCode,"insert","")
		KillSQLCode = Replace(KillSQLCode,"update","")
		KillSQLCode = Replace(KillSQLCode,"delete","")
		KillSQLCode = Replace(KillSQLCode,"exec","")
		KillSQLCode = Replace(KillSQLCode,"select","")
        KillSQLCode = Replace(KillSQLCode,"--","")
    end if
End Function

Function KillServerCode(s)
	KillServerCode = LCase(s)
	if Trim(KillServerCode)<>"" then  
		KillServerCode = Replace(KillServerCode,"<","&lt")     
        KillServerCode = Replace(KillServerCode,">","&gt")
		KillServerCode = Replace(KillServerCode,"script","&#115cript")
    end if
End Function

Function ClearStr(s)
	ClearStr = s
	if Trim(ClearStr)<>"" then 
	    ClearStr = Replace(ClearStr,"<","")     
        ClearStr = Replace(ClearStr,">","")     
        ClearStr = Replace(ClearStr,";","")   
        ClearStr = Replace(ClearStr,"'","")    
        ClearStr = Replace(ClearStr,"""","")   
        ClearStr = Replace(ClearStr,Chr(9),"")
        ClearStr = Replace(ClearStr,Chr(10),"")   
        ClearStr = Replace(ClearStr,Chr(13),"")    
        ClearStr = Replace(ClearStr,Chr(32),"")   
        ClearStr = Replace(ClearStr,Chr(34),"")  
        ClearStr = Replace(ClearStr,Chr(39),"")   
        ClearStr = Replace(ClearStr,"script","")   
        ClearStr = Replace(ClearStr,"(","")
        ClearStr = Replace(ClearStr,")","") 
        ClearStr = Replace(ClearStr,"--","") 
    end if
End Function

Function EnCodeStr(s)
	EnCodeStr = s
	if Trim(EnCodeStr)<>"" then 
	    EnCodeStr = Replace(EnCodeStr,"<","&lt")     
        EnCodeStr = Replace(EnCodeStr,">","&gt")     
        EnCodeStr = Replace(EnCodeStr,";","&#59;")   
        EnCodeStr = Replace(EnCodeStr,"'","&#39;")    
        EnCodeStr = Replace(EnCodeStr,"""","&quot;")   
        EnCodeStr = Replace(EnCodeStr,Chr(9),"&nbsp;")
        EnCodeStr = Replace(EnCodeStr,Chr(10),"<br>")   
        EnCodeStr = Replace(EnCodeStr,Chr(13),"")    
        EnCodeStr = Replace(EnCodeStr,Chr(32),"&nbsp;")   
        EnCodeStr = Replace(EnCodeStr,Chr(34),"&quot;")  
        EnCodeStr = Replace(EnCodeStr,Chr(39),"&#39;")   
        EnCodeStr = Replace(EnCodeStr,"script","&#115cript")   
        EnCodeStr = Replace(EnCodeStr,"(","&#40;")
        EnCodeStr = Replace(EnCodeStr,")","&#41;") 
        EnCodeStr = Replace(EnCodeStr,"--","&#45;&#45;") 
    end if
End Function

Function DeCodeStr(s)
	DeCodeStr = s
	if Trim(DeCodeStr)<>"" then 
	    DeCodeStr = Replace(DeCodeStr,"&lt","<")     
        DeCodeStr = Replace(DeCodeStr,"&gt",">")     
        DeCodeStr = Replace(DeCodeStr,"&#59;",";")   
        DeCodeStr = Replace(DeCodeStr,"&#39;","'")    
        DeCodeStr = Replace(DeCodeStr,"&quot;","""")   
        DeCodeStr = Replace(DeCodeStr,"&nbsp;",Chr(9))
        DeCodeStr = Replace(DeCodeStr,"<br>",Chr(10))   
        DeCodeStr = Replace(DeCodeStr,"",Chr(13))    
        DeCodeStr = Replace(DeCodeStr,"&nbsp;",Chr(32))   
        DeCodeStr = Replace(DeCodeStr,"&quot;",Chr(34))  
        DeCodeStr = Replace(DeCodeStr,"&#39;",Chr(39))   
        DeCodeStr = Replace(DeCodeStr,"&#115cript","script")   
        DeCodeStr = Replace(DeCodeStr,"&#40;","(")
        DeCodeStr = Replace(DeCodeStr,"&#41;",")") 
        DeCodeStr = Replace(DeCodeStr,"&#45;&#45;","--") 
    end if
End Function 

%>

