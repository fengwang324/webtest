  <!--#include file="conn.asp"--> 


<style type="text/css">
<!--
.style1 {font-size: 12px}
.input1 {	BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 9pt; BORDER-LEFT: #000000 1px solid; COLOR: rgb(0,0,0); BORDER-BOTTOM: #000000 1px solid; BACKGROUND-COLOR: #ffffff
}
body {
	background-color: #EFF7FF;
	margin-top: 0px;
}
.style4 {
	font-size: 12px;
	font-weight: bold;
	color: #000066}
.style5 {color: #000066}
.style3 {font-size: 16px;
	font-weight: bold;
}
-->
</style>

<div align="center">
  <table width="80%"  border=0 cellpadding=0 cellspacing=1 bgcolor=#818181 style="word-break:break-all">
    <tr>
      <td  valign="top" bgcolor=#EBF4FF>
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
          <tr>
            <td background="../../img/1-g1.gif" height=19 >
              <table width=548 height="18" border=0 cellpadding=0 cellspacing=0>
                <tr>
                  <td width=11>
                    <div align=center></div></td>
                  <td width=223 class=style1 style3><strong class="style4">
                   &nbsp;留言管理&nbsp;&nbsp;</strong>
                       
                     <%
					 Set rs=Server.CreateObject("ADODB.Recordset")
Sqln="select * from khly order by ntime desc,id desc"
rs.open Sqln,conn,1,1
 page=Request.QueryString("page")
If page="" Then page="1" 
page=cint(page)
rs.PageSize=20  '设置每页的记录数
tPagesize=rs.PageSize
tcount=rs.RecordCount '取得所有记录的总数
tpage=rs.PageCount'取得页面的总数

%>                  </td>
                  <td width=133 class=style1>目前共有[<font color="#FF0000"><%=tcount%></font>]条</td>
                  <td width=108 class=style1></td>
                  <td width=9 class=style1><div align="center"></div></td>
                  <td width=64 class=style1 style3>&nbsp;</td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td bgcolor=#0072bc height=2></td>
          </tr>
        </table>
        <div align="center">
          <table width="100%"  border="0" cellspacing="1" cellpadding="0" bgcolor=#818181>
            <tr class="style1">
			  <td width="12%" height="26" bgcolor="#EBF4FF"><div align="center" class="style5">姓名</div></td>
              <td width="22%" bgcolor="#EBF4FF"><div align="center" class="style5">电话</div></td>
              <td width="37%" bgcolor="#EBF4FF"><div align="center" class="style5">地址</div></td>
              <td width="15%" bgcolor="#EBF4FF"><div align="center" class="style5">留言日期</div></td>
              <td width="7%" bgcolor="#EBF4FF"><div align="center" class="style5">查看</div></td>
              <td width="7%" bgcolor="#EBF4FF"><div align="center" class="style5">删除</div></td>
            </tr>
             <%If Not rs.BOF Then
                %><% 
               rs.Move (page-1)*rs.PageSize 
                For ikhly=1 to tPagesize
             ' 将记录定位在所翻到的页面的首个记录上
                %>
			<tr>
			  <td height="20" bgcolor="#EBF4FF" align="center"><font  class="style1" color="#000000"><%=rs("xingming")%></font></td>
              <td align="left" bgcolor="#EBF4FF"><font  class="style1" color="#000000">&nbsp;<%=rs("dianhua")%></font></td>
              <td align="left" bgcolor="#EBF4FF"><font  class="style1" color="#000000">&nbsp;<%=rs("dizhi")%></font></td>
              <td align="center" bgcolor="#EBF4FF"><font  class="style1" color="#000000"><%=year(rs("ntime"))%>-<%=month(rs("ntime"))%>-<%=day(rs("ntime"))%></font></td>
              <td bgcolor="#EBF4FF"><div align="center" class="style1"><a href='khlyshow.asp?id=<%=rs("id")%>'>查看</a></div></td>
              <td bgcolor="#EBF4FF"><div align="center" class="style1"><a href='inc_khlydel.asp?id=<%=rs("id")%>' onClick="{if(confirm('您真的要删除这条信息吗?')){this.document.form1.submit();return true;}return false;}">删除</a></div></td>
            </tr>
			
			
			<% rs.MoveNext
If rs.EOF Then Exit For  '移到下一个记录
Next '如果到了数据库的末尾，就退出for
end if%>
           <tr class="style1">
              <td height="17" colspan="6" bgcolor="#EBF4FF"><div align="right" class="style5">
                <font color="#FF0000"><%=tPagesize%></font>个/页&nbsp;[<%=tpage%>]页&nbsp;当前<%=page%>/<%=tpage%>&nbsp;&nbsp;<% '第一页
If page>1 Then %>
                          <a href="khly.asp?page=1">首页</a>
                          <% Else %>
        首页
        <% End If %>
        <% '上下页面设置
If page>1 Then %>
        <a href="khly.asp?page=<%=page-1%>">上页</a>
        <% Else %>
        上页
        <% End If %>
        <% If page<tpage Then %>
        <a href="khly.asp?page=<%=page+1%>">下页</a>
        <% Else %>
        下页
        <% End If %>
        <% '最后一页
If page<tpage Then %>
        <a href="khly.asp?page=<%=tpage%>">尾页</a>
        <% Else %>
        尾页
        <% End If %>转到&nbsp;<SELECT name="goto" size="1"  onchange="javascript:location.href='khly.asp?page='+goto.options[this.selectedIndex].value">
                      <% for n=1 to tpage
	   if n=cint(request.QueryString("page")) then
	      response.Write "<option value='" & n &"' selected>第" & n & "页</option>"
	   else
	      response.Write "<option value='" & n &"'>第" & n & "页</option>"
	   end if
	 next%>
                          </Select>&nbsp;&nbsp;<%rs.close
set rs=nothing
conn.close
set conn=nothing%>
                </div>
              </td>
            </tr>
		  
		  </table>
        </div></td>
    </tr>
  </table>
</div>
