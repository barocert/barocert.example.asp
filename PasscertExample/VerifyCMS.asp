<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert ASP Example</title>
	</head>
<!--#include file="common.asp"--> 

<%
	'**************************************************************
	' �Ϸ�� ���ڼ����� �����ϰ� ���ڼ���(signedData)�� ��ȯ �޽��ϴ�.
	' ���� �Լ��� �ڵ���ü ��ݵ��� ��û �Լ��� ȣ���� ���� 23�� 59�� 59�ʱ����� ȣ�� �����մϴ�.
	' �ڵ���ü ��ݵ��� ��û �Լ��� ȣ���� ���� 23�� 59�� 59�� ���� ���� �Լ��� ȣ���� ��� ������ ��ȯ�˴ϴ�.
	' https://developers.barocert.com/reference/pass/asp/cms/api#VerifyCMS
	'**************************************************************

	' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
	Dim clientCode : clientCode = "023070000014"	

	' �ڵ���ü ��ݵ��� ��û�� ��ȯ�� �������̵�
	Dim receiptID : receiptID = "02309180230700000140000000000006"

	Dim verifyCMS : Set verifyCMS = New CMSVerify

	verifyCMS.receiverHP = m_PasscertService.encrypt("01012341234")
	
	verifyCMS.receiverName = m_PasscertService.encrypt("ȫ�浿")

	On Error Resume Next

		Dim result : Set result = m_PasscertService.VerifyCMS(clientCode, receiptID, verifyCMS)

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
				<legend>�н� ��ݵ��� ����</legend>
				<% If code = 0 Then %>
					<ul>
						<li>���� ���̵� (ReceiptID) : <%=result.receiptID %></li>
						<li>���� (State) : <%=result.state %></li>
						<li>������ ���� (ReceiverName) : <%=result.receiverName %></li>
						<li>������ ����⵵ (ReceiverYear) : <%=result.receiverYear %></li>
						<li>������ ������� (ReceiverDay) : <%=result.receiverDay %></li>
						<li>������ ���� (ReceiverHP) : <%=result.receiverHP %></li>
						<li>������ �޴�����ȣ (ReceiverGender) : <%=result.receiverGender %></li>
						<li>�ܱ��� ���� (ReceiverForeign) : <%=result.receiverForeign %></li>
						<li>��Ż� ���� (ReceiverTelcoType) : <%=result.receiverTelcoType %></li>
						<li>���ڼ��� ������ ���� (SignedData) : <%=result.signedData %></li>
						<li>�������� (Ci) : <%=result.ci %></li>
					</ul>
				<% Else %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<% End If %>
			</fieldset>
		</div>
	</body>
</html>