<%
Sub Actionbutton(menuid,rightnum,addurl)
	dim Action_sql,Action_rs,rs_right

			if rightnum="3" then

			if menuid="100" and InStr(nowpath,"billlist.asp")>0 then
			else
				if addurl="1" then
					response.write "<input name='menuid' type='hidden' id='menuid' value='"&request("menuid")&"'>"
					response.write "<input type='submit' name='Button91' value='增 加'>"
				else
					response.write "<input type='button' name='Button91' value='增 加' onClick=window.location.href='"&addurl&"'>"
				end if
			end if

			elseif rightnum="4" then
				response.write "<input name='menuid' type='hidden' id='menuid' value='"&request("menuid")&"'>"
			%>
				<input type='submit' name='submit91' value='删除选中' onclick="return confirm('删除记录有可能导致数据不完整，请确保所选记录还没有与其他记录相关联。\n现真的删除所选中记录吗？')">
			<%
			elseif rightnum="5" then
				response.write "<input name='menuid' type='hidden' id='menuid' value='"&request("menuid")&"'>"
				response.write "<input type='submit' name='submit91' value='修 改'>"
			else
				call alert("出错啦。")
			end if

End Sub
%>


