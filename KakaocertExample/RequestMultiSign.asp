<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 

<%
	'**************************************************************
	' īī���� ����ڿ��� ���ڼ����� ��û�մϴ�.(����)
    '**************************************************************

	' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ�, (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
	Dim clientCode : clientCode = "023040000001"		
	
	' ���ڼ��� ��û ���� ��ü
    Dim reqMultiSign : Set reqMultiSign = new MultiSign
	' ������ �޴�����ȣ - 11�� (������ ����)
    reqMultiSign.ReceiverHP = m_KakaocertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqMultiSign.ReceiverName = m_KakaocertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqMultiSign.ReceiverBirthday = m_KakaocertService.encrypt("19700101")
	' ������û �޽��� ���� - �ִ� 40��
    reqMultiSign.ReqTitle = "���ڼ������׽�Ʈ"
	' ������û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqMultiSign.ExpireIn = 1000

	' �������� ��� - �ִ� 20 ��
    Set tokens = CreateObject("Scripting.Dictionary")
    For i=0 To 2
        Set token = New MultiSignTokens
        ' ������û �޽��� ���� - �ִ� 40��
		token.ReqTitle = "���ڼ����������׽�Ʈ1"
		' ���� ���� - ���� 2,800�� ���� �Է°���
		token.Token = m_KakaocertService.encrypt("���ڼ������׽�Ʈ������"+CStr(i))
		reqMultiSign.addToken i, token
    Next

	' ���� ���� ����
	' TEXT - �Ϲ� �ؽ�Ʈ, HASH - HASH ������
    reqMultiSign.TokenType = "TEXT"
	' AppToApp ������û ����
	' true - AppToApp �������, false - Talk Message �������
    reqMultiSign.AppUseYN = false
	' App to App ��� �̿��, ������ ȣ���� URL
	' reqMultiSign.ReturnURL = "https://www.kakaocert.com"

	On Error Resume Next

		Dim result : Set result = m_KakaocertService.RequestMultiSign(clientCode, reqMultiSign)

		If Err.Number <> 0 then
			Dim code : code = Err.Number
			Dim message : message =  Err.Description
			Err.Clears
		End If

	On Error GoTo 0

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>īī�� ���ڼ��� ��û(����)</legend>
				<% 
				If code = 0 Then %>
					<ul>
						<li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
						<li>�۽�Ŵ (scheme) : <%=result.scheme %></li>
					</ul>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>