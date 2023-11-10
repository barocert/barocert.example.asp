<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �������� ��û �� ��ȯ���� �������̵�� �������� ���� ���¸� Ȯ���մϴ�.
	' https://developers.barocert.com/reference/naver/asp/identity/api#GetIdentityStatus
	'**************************************************************

	' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
	Dim clientCode : clientCode = "023090000021"	

	' �������� ��û�� ��ȯ�� �������̵�
	Dim receiptID : receiptID = "02311090230900000210000000000012"
	
	On Error Resume Next

	Dim result : Set result = m_NavercertService.GetIdentityStatus(clientCode, receiptID)

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
				<legend>���̹� �������� ����Ȯ��</legend>
				<% If code = 0 Then %>
					<ul>
						<li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
						<li>�̿��� �ڵ� (ClientCode) : <%=result.clientCode %></li>
						<li>���� (State) : <%=result.state %></li>
						<li>��û ����ð� (ExpireIn) : <%=result.expireIn %></li>
						<li>�̿��� �� (CallCenterName) : <%=result.callCenterName %></li>
						<li>�̿��� ����ó (CallCenterNum) : <%=result.callCenterNum %></li>
						<li>����URL (ReturnURL) : <%=result.returnURL %></li>
						<li>�������Ͻ� (ExpireDT) : <%=result.expireDT %></li>
						<li>�۽�Ŵ (Scheme) : <%=result.scheme %></li>
						<li>�ۻ������ (AppUseYN) : <%=result.appUseYN %></li>
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