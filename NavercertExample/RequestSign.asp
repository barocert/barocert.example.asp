<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' 네이버 이용자에게 단건(1건) 문서의 전자서명을 요청합니다
    ' https://developers.barocert.com/reference/naver/asp/sign/api-single#RequestSign
    '**************************************************************

    ' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
    Dim clientCode : clientCode = "023090000021"
    
    ' 전자서명 요청 정보 객체
    Dim reqSign : Set reqSign = new Sign
    ' 수신자 휴대폰번호 - 11자 (하이픈 제외)
    reqSign.ReceiverHP = m_NavercertService.encrypt("01012341234")
    ' 수신자 성명 - 80자
    reqSign.ReceiverName = m_NavercertService.encrypt("홍길동")
    ' 수신자 생년월일 - 8자 (yyyyMMdd)
    reqSign.ReceiverBirthday = m_NavercertService.encrypt("19700101")
    ' 인증요청 메시지 제목 - 최대 40자
    reqSign.ReqTitle = "전자서명(단건) 요청 메시지 제목"
    ' 인증요청 메시지 - 최대 500자
    reqSign.ReqMessage = m_NavercertService.encrypt("전자서명(단건) 요청 메시지")
    ' 고객센터 연락처 - 최대 12자
    reqSign.CallCenterNum = "1600-9854"
    ' 인증요청 만료시간 - 최대 1,000(초)까지 입력 가능
    reqSign.ExpireIn = 1000
    ' 서명 원문 유형
    ' TEXT - 일반 텍스트, HASH - HASH 데이터
    reqSign.TokenType = "TEXT"
    ' 서명 원문 - 원문 2,800자 까지 입력가능
    reqSign.Token = m_NavercertService.encrypt("전자서명(단건) 요청 원문")
    ' 서명 원문 유형
    ' reqSign.TokenType = "HASH"
    ' 서명 원문 유형이 HASH인 경우, 원문은 SHA-256, Base64 URL Safe No Padding을 사용
    ' reqSign.Token = m_NavercertService.encrypt(m_NavercertService.sha256_base64url("전자서명(단건) 요청 원문"))
    ' AppToApp 인증요청 여부
    ' true - AppToApp 인증방식, false - Talk Message 인증방식
    reqSign.AppUseYN = false
    ' AppToApp 인증방식에서 사용
    ' 모바일장비 유형('ANDROID', 'IOS'), 대문자 입력(대소문자 구분)
    ' reqSign.DeviceOSType = "ANDROID";
    ' AppToApp 방식 이용시, 호출할 URL
    ' "http", "https"등의 웹프로토콜 사용 불가
    ' reqSign.ReturnURL = "navercert://sign";

    On Error Resume Next

        Dim result : Set result = m_NavercertService.RequestSign(clientCode, reqSign)

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
                <legend>네이버 전자서명(단건) 요청</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>접수아이디 (ReceiptID) : <%=result.receiptID %></li>
                        <li>앱스킴 (scheme) : <%=result.scheme %></li>
                        <li>앱다운로드URL (marketUrl) : <%=result.marketUrl %></li>
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