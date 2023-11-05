<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Sub Create_Folder(str)
	dim newstr, fso9
	newstr = Server.Mappath(str)
	set fso9=server.createobject("scripting.filesystemobject")
    ' 检查文件夹是否存在
    If fso9.FolderExists(strTestFolder) Then
      'objFSO.DeleteFolder(strTestFolder)  '已存在,不用创建
    Else
      fso9.CreateFolder(newstr)
    End If
    Set fso9 = Nothing
End Sub

Function Format_Time(s_Time, n_Flag)
	Dim y, m, d, h, mi, s
	Format_Time = ""
	If IsDate(s_Time) = False Then Exit Function
	y = cstr(year(s_Time))
	m = cstr(month(s_Time))
	If len(m) = 1 Then m = "0" & m
	d = cstr(day(s_Time))
	If len(d) = 1 Then d = "0" & d
	h = cstr(hour(s_Time))
	If len(h) = 1 Then h = "0" & h
	mi = cstr(minute(s_Time))
	If len(mi) = 1 Then mi = "0" & mi
	s = cstr(second(s_Time))
	If len(s) = 1 Then s = "0" & s
	Select Case n_Flag
	Case 1
	' yyyy-mm-dd hh:mm:ss
	Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
	Case 2
	' yyyy-mm-dd
	Format_Time = y & "-" & m & "-" & d
	Case 3
	' hh:mm:ss
	Format_Time = h & ":" & mi & ":" & s
	Case 4
	' yyyy年mm月dd日
	Format_Time = y & "年" & m & "月" & d & "日"
	Case 5
	' yyyymmdd
	Format_Time = y & m & d
	case 6
	'yyyymmddhhmmss
	format_time= y & m & d & h & mi & s
	End Select
End Function

dim foldername
foldername = format_time(now,6)
foldername = mid(foldername,3)
foldername = left(foldername,6)&"_"&mid(foldername,7)

if Session("AdminID")<>"" then
	call Create_Folder("../uploadimages/"&foldername)	
end if

response.Redirect("Admin_UploadFile_Left.asp")
%>


