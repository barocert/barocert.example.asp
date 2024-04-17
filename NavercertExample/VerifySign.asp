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
	' ���̹� ������å�� ���� ���� API�� 1ȸ�� ȣ���� �� �ֽ��ϴ�. ��õ��� ������ ��ȯ�˴ϴ�.
	' ���ڼ��� �����Ͻ� ���Ŀ� ���� API�� ȣ���ϸ� ������ ��ȯ�˴ϴ�.
	' https://developers.barocert.com/reference/naver/asp/sign/api-single#VerifySign
	'**************************************************************

	' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
	Dim clientCode : clientCode = "023090000021"	

	' ���ڼ���(�ܰ�) ��û�� ��ȯ�� �������̵�
	Dim receiptID : receiptID = "02311090230900000210000000000010"

	On Error Resume Next

		Dim result : Set result = m_NavercertService.VerifySign(clientCode, receiptID)

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
				<legend>���̹� ���ڼ���(�ܰ�) ����</legend>
				<% If code = 0 Then %>
					<ul>
						<li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
						<li>���� (State) : <%=result.state %></li>
						<li>���ڼ��� ������ ���� (SignedData) : <%=result.signedData %></li>
						<li>�������� (Ci) : <%=result.ci %></li>
						<li>������ ���� (ReceiverName) : <%=result.receiverName %></li>
						<li>������ ����⵵ (ReceiverYear) : <%=result.receiverYear %></li>
						<li>������ ������� (ReceiverDay) : <%=result.receiverDay %></li>
						<li>������ �޴�����ȣ (ReceiverHP) : <%=result.receiverHP %></li>
						<li>������ ���� (ReceiverGender) : <%=result.receiverGender %></li>
						<li>������ �̸��� (ReceiverEmail) : <%=result.receiverEmail %></li>
						<li>�ܱ��� ���� (ReceiverForeign) : <%=result.receiverForeign %></li>
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