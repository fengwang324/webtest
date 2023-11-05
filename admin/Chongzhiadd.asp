<!--#include file="conn.asp"-->
<!--#include file="check4.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link REL='stylesheet' HREF='szcrm.css' TYPE='text/css'>
<SCRIPT language=JavaScript>
<!--
function Juge(theForm)
{

   if (theForm.customname.value == "")
  {
    alert("请填写会员。");
    theForm.customname.focus();
    return (false);
  } 

  if (theForm.czdate.value == "" || theForm.czdate.value == "null")
  {
    alert("请填写充值日期。");
    theForm.czdate.focus();
    return (false);
  }  

  if (theForm.czby.value == "")
  {
    alert("请填写充值方式。");
    theForm.czby.focus();
    return (false);
  }

  if (theForm.czmoney.value == "")
  {
    alert("请填写充值金额。");
    theForm.czmoney.focus();
    return (false);
  }
  if (isNaN(theForm.czmoney.value))
  {
    alert("请充值金额要为数值。");
    theForm.czmoney.focus();
    return (false);
  }  

  if (theForm.confirmflag.value == "")
  {
    alert("请选择是否确认。");
    theForm.confirmflag.focus();
    return (false);
  }
   if(confirm("确认新增吗?"))
      return true
   else
      return false;

}
-->
</script>
</head>
<body topmargin="10">
  <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>

  <table width="95%" align="center" border=1 cellspacing=0 cellpadding=1 bordercolorlight=#5BD77D bordercolordark=#ffffff>
   <form name="theForm" method="post" action="chongzhiaddadmin.asp?menuid=<%=menuid%>" onsubmit="return Juge(this)"><tr>
      <td class=titlehead height=20>增加充值信息：</td>
    </tr>
  </table>
  <table width="95%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableborder">

    <tr> 
      <td width="12%" height="5"></td>
      <td width="26%"></td>
      <td width="10%"></td>
      <td width="51%"></td>
    </tr>
    <tr height=35>
	<td align='right'>充值日期</td>
	<td><input type='text' name='czdate' maxlength='50' value='<%=date()%>'></td>
	<td align="right">充值会员</td>
	<td ><input type="text" name="customname" value="<%=customname%>" size="16" class="input1" readonly>
	<input type=button value='…' onClick="javascript:window.open('openhy.asp?tan=1&menuid=28','','top=72,left=100,scrollbars=yes,width=670,height=530');">
	<input type="hidden" name="customid" value="<%=customid%>">
	</td>	
	</tr>

	
    <tr height=35> 
	<td align='right'>充值方式</td>
	<td><input type='text' name='czby' maxlength='50' value=''></td>
	<td align='right'>充值金额</td>
	<td><input type='text' name='czmoney' maxlength='50' value=''></td>
    </tr>


    <tr height=35> 
	<td align='right'>是否确认</td>
	<td> <select name='confirmflag' size='1' ><option value=''></option><option value='否'>否</option><option value='是'>是</option></select></td> 
	<td align='right'>确认人/时间</td> 
	<td><input type='text' name='operator' maxlength='50' value='' class='input1' readonly></td>  
    </tr> 
    <tr  height=35>
	<td align='right'>备注</td>
	<td colspan=3><input type='text' name='czmemo' size='64' maxlength='70' value=''></td>
    </tr>
  </table>
  <table width="82%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td height="40"> 
		<div align="center">
		<input type="Submit" name="Submit2" value="确 定">
          &nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="reset" name="Submit2" value="清 除">
          &nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="button" name="Submit22" value="返 回" onclick=javascript:history.back()>
        </div></td>
    </tr>
  </table>
</form>
</body>
</html>
<%conn.close()
set conn=nothing%>







