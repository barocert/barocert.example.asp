<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'************************************************************** 
	' 자동이체 출금동의 요청 후 반환받은 접수아이디로 인증 진행 상태를 확인합니다.
	' 상태확인 함수는 자동이체 출금동의 요청 함수를 호출한 당일 23시 59분 59초까지만 호출 가능합니다.
	' 자동이체 출금동의 요청 함수를 호출한 당일 23시 59분 59초 이후 상태확인 함수를 호출할 경우 오류가 반환됩니다.
	' https://developers.barocert.com/reference/pass/asp/cms/api#GetCMSStatus
	'**************************************************************

	' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
	Dim clientCode : clientCode = "023070000014"	

	' 출금동의 요청시 반환된 접수아이디
	Dim receiptID : receiptID = "02312130230700000140000000000007"
	
	On Error Resume Next

	Dim result : Set result = m_PasscertService.GetCMSStatus(clientCode, receiptID)

	If Err.Number <> 0 Then
		Dim code : code = Err.Number
		Dim message : message = Err.Description
		Err.Clears
	End If	
'	On Error GoTo 0 
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>패스 출금동의 상태확인 </legend>
				<% If code = 0 Then %>
					<ul>
						<li>이용기관 코드 (ClientCode) : <%=result.clientCode %></li>
						<li>접수아이디 (ReceiptID) : <%=result.receiptID %></li>
						<li>상태 (State) : <%=result.state %></li>
						<li>서명요청일시 (RequestDT) : <%=result.requestDT %></li>
						<li>서명완료일시 (CompleteDT) : <%=result.completeDT %></li>
						<li>서명만료일시 (ExpireDT) : <%=result.expireDT %></li>
						<li>서명거절일시 (RejectDT) : <%=result.rejectDT %></li>
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