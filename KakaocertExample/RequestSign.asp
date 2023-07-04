<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 

<%
	'**************************************************************
	' 카카오톡 사용자에게 전자서명을 요청합니다.(단건)
    '**************************************************************

	' 이용기관코드, 파트너가 등록한 이용기관의 코드, (파트너 사이트에서 확인가능)
	Dim clientCode : clientCode = "023040000001"		
	
	' 전자서명 요청 정보 객체
    Dim reqSign : Set reqSign = new Sign
	' 수신자 휴대폰번호 - 11자 (하이픈 제외)
    reqSign.ReceiverHP = m_KakaocertService.encrypt("01012341234")
    ' 수신자 성명 - 80자
    reqSign.ReceiverName = m_KakaocertService.encrypt("홍길동")
    ' 수신자 생년월일 - 8자 (yyyyMMdd)
    reqSign.ReceiverBirthday = m_KakaocertService.encrypt("19700101")
    ' 인증요청 메시지 제목 - 최대 40자
    reqSign.ReqTitle = "전자서명단건테스트"
    ' 인증요청 만료시간 - 최대 1,000(초)까지 입력 가능
    reqSign.ExpireIn = 1000
    ' 서명 원문 - 원문 2,800자 까지 입력가능
    reqSign.Token = m_KakaocertService.encrypt("전자서명단건테스트데이터")
    ' 서명 원문 유형
    ' TEXT - 일반 텍스트, HASH - HASH 데이터
    reqSign.TokenType = "TEXT"
    ' AppToApp 인증요청 여부
    ' true - AppToApp 인증방식, false - Talk Message 인증방식
    reqSign.AppUseYN = false
    ' App to App 방식 이용시, 호출할 URL
    ' reqSign.ReturnURL = "https://www.kakaocert.com"

	On Error Resume Next

		Dim result : Set result = m_KakaocertService.RequestSign(clientCode, reqSign)

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
				<legend>카카오 전자서명 요청(단건)</legend>
				<% 
				If code = 0 Then %>
					<ul>
						<li>접수아이디 (ReceiptID) : <%=result.receiptID %></li>
						<li>앱스킴 (scheme) : <%=result.scheme %></li>
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