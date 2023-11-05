<%
''on error resume next
Response.Charset = "UTF-8"
dim strDir,strAdminDir
			strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
			strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
			InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
			InstallDir=InstallDir&strAdminDir
Dim dbPaths:dbPaths = right(Replace(InstallDir,"cn/","") & "DB/D#B.asa",len(InstallDir & "DB/D#B.asa")-1)
if left(dbPaths,1)<>"/" then dbPaths="/"&dbPaths
%>
