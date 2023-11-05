<%
Dim ShowFileStyle
ShowFileStyle = request("ShowFileStyle")
If ShowFileStyle = "" Or Not IsNumeric(ShowFileStyle) Then
    ShowFileStyle = 1
Else
    ShowFileStyle = Int(ShowFileStyle)
End If
Response.cookies("ShowFileStyle") = ShowFileStyle
adminurl=request.servervariables("http_referer")
Response.redirect adminurl
%>
