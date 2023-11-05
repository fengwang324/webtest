<!--#include file="CONST.asp"-->
<%
Dim Tim_Get,Tim_In,Tim_Inf,Tim_Xh,SiteSn
Tim_In="'※;※and※exec※insert※select※delete※update※count※*※%※chr※mid※master※truncate※char※declare※script" 
Tim_Inf = Split(Tim_In,"※") 
SiteSn="StarTeacher"
If Request.QueryString<>"" Then 
	For Each Tim_Get In Request.QueryString 
		For Tim_Xh=0 To Ubound(Tim_Inf) 
			If Instr(LCase(Request.QueryString(Tim_Get)),Tim_Inf(Tim_Xh)) Then Response.End 
		Next 
	Next 
End If

Dim ConnStr,conn
    ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath(dbPaths)
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open ConnStr
	   SqlNowString = "Now()"
	   DataPart_D   = "'d'"
	   DataPart_Y   = "'yyyy'"
	   DataPart_H   = "'h'"
If Err Then
	err.Clear
	Set conn = Nothing
	Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。"
	Response.End
End If
Sub CloseConn()
	'on error resume next
	If IsObject(rs) Then
		set rs=nothing
	end if
	If IsObject(conn) Then
		conn.close
		set conn=nothing
	end if
End Sub

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
sub c2(str)
Response.Write(str)
Response.end
end sub%>