<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 

<%
	'**************************************************************
	' ���ڼ��� ��û�� ��ȯ�� �������̵� ���� ������ �����մϴ�. (�ܰ�)
	' �����ϱ� API�� �Ϸ�� ���ڼ��� ��û�� 1ȸ�� ��û �����ϸ�, ����ڰ� ������ �Ϸ��� ��ȿ�ð�(10��)�̳����� ��û���� �մϴ�.
	'**************************************************************

	' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ�, (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
	Dim clientCode : clientCode = "023040000001"	

	' ���ڼ��� ��û�� ��ȯ�� �������̵�
	Dim receiptID : receiptID = "02307040230400000010000000000027"

	On Error Resume Next

		Dim result : Set result = m_KakaocertService.VerifyMultiSign(clientCode, receiptID)

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
				<legend>īī�� ���ڼ��� ����(����)</legend>
				<% If code = 0 Then %>
					<ul>
						<li>���� ���̵� (ReceiptID) : <%=result.receiptID %></li>
						<li>���� (State) : <%=result.state %></li>
						<li>�������� (Ci) : <%=result.ci %></li>
					</ul>
					<%
						For i=0 To UBound(result.multiSignedData) -1
					%>
						<fieldset class="fieldset2">
							<ul>
								<li>���ڼ��� ������ ���� (SignedData) : <%=result.multiSignedData(i)%></li>
							</ul>
						</fieldset>
					<%
						Next
					Else
					%>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>