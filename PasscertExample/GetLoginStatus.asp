<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ����α��� ��û �� ��ȯ���� �������̵�� ���� ���¸� Ȯ���մϴ�.
	' ����Ȯ�� �Լ��� ����α��� ��û �Լ��� ȣ���� ���� 23�� 59�� 59�ʱ����� ȣ�� �����մϴ�.
	' ����α��� ��û �Լ��� ȣ���� ���� 23�� 59�� 59�� ���� ����Ȯ�� �Լ��� ȣ���� ��� ������ ��ȯ�˴ϴ�.
	' https://developers.barocert.com/reference/pass/asp/login/api#GetLoginStatus
	'**************************************************************

	' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
	Dim clientCode : clientCode = "023040000001"	

	' ����α��� ��û�� ��ȯ�� �������̵�
	Dim receiptID : receiptID = "02307040230400000010000000000027"
	
	On Error Resume Next

	Dim result : Set result = m_PasscertService.GetLoginStatus(clientCode, receiptID)

	If Err.Number <> 0 Then
		Dim code : code = Err.Number
		Dim message : message = Err.Description
		Err.Clears
	End If	
	On Error GoTo 0 
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>�н� ����α��� ����Ȯ��</legend>
				<% 
					If code = 0 Then 
				%>
					<ul>
						<li>�̿��� �ڵ� (ClientCode) : <%=result.clientCode %></li>
						<li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
						<li>���� (State) : <%=result.state %></li>
						<li>��û ����ð� (ExpireIn) : <%=result.expireIn %></li>
						<li>�̿��� �� (CallCenterName) : <%=result.callCenterName %></li>
						<li>�̿��� ����ó (CallCenterNum) : <%=result.callCenterNum %></li>
						<li>������û �޽��� ���� (ReqTitle) : <%=result.reqTitle %></li>
						<li>������û �޽��� (ReqMessage) : <%=result.reqMessage %></li>
						<li>�����û�Ͻ� (RequestDT) : <%=result.requestDT %></li>
						<li>����Ϸ��Ͻ� (CompleteDT) : <%=result.completeDT %></li>
						<li>�������Ͻ� (ExpireDT) : <%=result.expireDT %></li>
						<li>��������Ͻ� (RejectDT) : <%=result.rejectDT %></li>
						<li>���� ���� (TokenType) : <%=result.tokenType %></li>
						<li>����ڵ����ʿ俩�� (UserAgreementYN) : <%=result.userAgreementYN %></li>
						<li>������������Կ��� (ReceiverInfoYN) : <%=result.receiverInfoYN %></li>
						<li>��Ż� ���� (TelcoType) : <%=result.telcoType %></li>
						<li>�������� ���� (DeviceOSType) : <%=result.deviceOSType %></li>
						<li>�۽�Ŵ (Scheme) : <%=result.scheme %></li>
						<li>�ۻ������ (AppUseYN) : <%=result.appUseYN %></li>
					</ul>	
					<%	
						Else
					%>
						<ul>
							<li>Response.code: <%=code%> </li>
							<li>Response.message: <%=message%> </li>
						</ul>	
					<%	
						End If
					%>
			</fieldset>
		</div>
	</body>
</html>