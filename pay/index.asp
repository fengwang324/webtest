
<!--#include file="alipayto/alipay_payto.asp"-->
<%
   shijian=now()
   dingdan=year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
    '�ͻ���վ�����ţ�����ȡϵͳʱ�䣬�ɸĳ���վ�Լ��ı�����
	
	subject			=request.Cookies("subject") '	"���Զ�����"	'��Ʒ���ƣ�����ͻ��߹��ﳵ���̿�����Ϊ  "�����ţ�"&request("�ͻ���վ����")
	body			=request.Cookies("body") 	'	"��Ʒ����"		'��Ʒ����
	out_trade_no    =request.Cookies("subject") '   dingdan         '��ʱ���ȡ�Ķ����ţ������޸ĳ��Լ���վ�Ķ����ţ���֤ÿ���ύ�Ķ���Ψһ����
	price		    =request.Cookies("price") 	'	"0.01"			'��Ʒ����			0.01��100000000.00  ��ע����Ҫ����3,000.00���۸�֧��","��
    quantity        =   "1"             		'��Ʒ����,����߹��ﳵĬ��Ϊ1
	discount        =   "0"             		'��Ʒ�ۿ�
    seller_email    =request.Cookies("seller_email") '    seller_email   '���ҵ�֧�����ʺţ�c2c�ͻ������Ը��Ĵ˲�����

	'�����޸ĵ�:

	'alipay_Config.asp
	'ȫ��
	
	'Alipay_Notify.asp
	'key=request.Cookies("key") '""         		'֧������ȫ������
	'partner=request.Cookies("partner") '"" 

	
	Set AlipayObj	= New creatAlipayItemURL
	itemUrl=AlipayObj.creatAlipayItemURL(subject,body,out_trade_no,price,quantity,seller_email)
	response.Redirect(itemUrl)
%>