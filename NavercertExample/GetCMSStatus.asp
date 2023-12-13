<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' �ڵ���ü ��ݵ��� ��û �� ��ȯ���� �������̵�� ���� ���� ���¸� Ȯ���մϴ�.
	' https://developers.barocert.com/reference/naver/asp/cms/api#GetCMSStatus
	'**************************************************************

	' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
	Dim clientCode : clientCode = "023090000021"	

	' ��ݵ��� ��û�� ��ȯ�� �������̵�
	Dim receiptID : receiptID = "02312130230900000210000000000028"
	
	On Error Resume Next

	Dim result : Set result = m_NavercertService.GetCMSStatus(clientCode, receiptID)

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
				<legend>���̹� ��ݵ��� ����Ȯ��</legend>
				<% If code = 0 Then %>
					<ul>
						<li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
						<li>�̿��� �ڵ� (ClientCode) : <%=result.clientCode %></li>
						<li>���� (State) : <%=result.state %></li>
						<li>�������Ͻ� (ExpireDT) : <%=result.expireDT %></li>
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