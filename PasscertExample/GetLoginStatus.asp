<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 간편로그인 요청 후 반환받은 접수아이디로 진행 상태를 확인합니다.
	' 상태확인 함수는 간편로그인 요청 함수를 호출한 당일 23시 59분 59초까지만 호출 가능합니다.
	' 간편로그인 요청 함수를 호출한 당일 23시 59분 59초 이후 상태확인 함수를 호출할 경우 오류가 반환됩니다.
	' https://developers.barocert.com/reference/pass/asp/login/api#GetLoginStatus
	'**************************************************************

	' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
	Dim clientCode : clientCode = "023040000001"	

	' 간편로그인 요청시 반환된 접수아이디
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
				<legend>패스 간편로그인 상태확인</legend>
				<% 
					If code = 0 Then 
				%>
					<ul>
						<li>이용기관 코드 (ClientCode) : <%=result.clientCode %></li>
						<li>접수아이디 (ReceiptID) : <%=result.receiptID %></li>
						<li>상태 (State) : <%=result.state %></li>
						<li>요청 만료시간 (ExpireIn) : <%=result.expireIn %></li>
						<li>이용기관 명 (CallCenterName) : <%=result.callCenterName %></li>
						<li>이용기관 연락처 (CallCenterNum) : <%=result.callCenterNum %></li>
						<li>인증요청 메시지 제목 (ReqTitle) : <%=result.reqTitle %></li>
						<li>인증요청 메시지 (ReqMessage) : <%=result.reqMessage %></li>
						<li>서명요청일시 (RequestDT) : <%=result.requestDT %></li>
						<li>서명완료일시 (CompleteDT) : <%=result.completeDT %></li>
						<li>서명만료일시 (ExpireDT) : <%=result.expireDT %></li>
						<li>서명거절일시 (RejectDT) : <%=result.rejectDT %></li>
						<li>원문 구분 (TokenType) : <%=result.tokenType %></li>
						<li>사용자동의필요여부 (UserAgreementYN) : <%=result.userAgreementYN %></li>
						<li>사용자정보포함여부 (ReceiverInfoYN) : <%=result.receiverInfoYN %></li>
						<li>통신사 유형 (TelcoType) : <%=result.telcoType %></li>
						<li>모바일장비 유형 (DeviceOSType) : <%=result.deviceOSType %></li>
						<li>앱스킴 (Scheme) : <%=result.scheme %></li>
						<li>앱사용유무 (AppUseYN) : <%=result.appUseYN %></li>
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