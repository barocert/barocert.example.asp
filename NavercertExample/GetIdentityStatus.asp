<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 본인인증 요청 후 반환받은 접수아이디로 본인인증 진행 상태를 확인합니다.
	' https://developers.barocert.com/reference/naver/asp/identity/api#GetIdentityStatus
	'**************************************************************

	' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
	Dim clientCode : clientCode = "023090000021"	

	' 본인인증 요청시 반환된 접수아이디
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
				<legend>네이버 본인인증 상태확인</legend>
				<% If code = 0 Then %>
					<ul>
						<li>접수아이디 (ReceiptID) : <%=result.receiptID %></li>
						<li>이용기관 코드 (ClientCode) : <%=result.clientCode %></li>
						<li>상태 (State) : <%=result.state %></li>
						<li>요청 만료시간 (ExpireIn) : <%=result.expireIn %></li>
						<li>이용기관 명 (CallCenterName) : <%=result.callCenterName %></li>
						<li>이용기관 연락처 (CallCenterNum) : <%=result.callCenterNum %></li>
						<li>복귀URL (ReturnURL) : <%=result.returnURL %></li>
						<li>서명만료일시 (ExpireDT) : <%=result.expireDT %></li>
						<li>앱스킴 (Scheme) : <%=result.scheme %></li>
						<li>앱사용유무 (AppUseYN) : <%=result.appUseYN %></li>
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