<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert ASP Example</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' 네이버 이용자에게 복수(최대 50건) 문서의 전자서명을 요청합니다.
    ' https://developers.barocert.com/reference/naver/asp/sign/api-multi#RequestMultiSign
    '**************************************************************

    ' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
    Dim clientCode : clientCode = "023090000021"
    
    ' 전자서명 요청 정보 객체
    Dim reqMultiSign : Set reqMultiSign = new MultiSign
    ' 수신자 휴대폰번호 - 11자 (하이픈 제외)
    reqMultiSign.ReceiverHP = m_NavercertService.encrypt("01012341234")
    ' 수신자 성명 - 80자
    reqMultiSign.ReceiverName = m_NavercertService.encrypt("홍길동")
    ' 수신자 생년월일 - 8자 (yyyyMMdd)
    reqMultiSign.ReceiverBirthday = m_NavercertService.encrypt("19700101")
    ' 인증요청 메시지 제목 - 최대 40자
    reqMultiSign.ReqTitle = "전자서명(복수) 요청 메시지 제목"
    ' 고객센터 연락처 - 최대 12자
    reqMultiSign.CallCenterNum = "1600-9854"
    ' 인증요청 메시지 - 최대 500자
    reqMultiSign.ReqMessage = m_NavercertService.encrypt("전자서명(복수) 요청 메시지")
    ' 인증요청 만료시간 - 최대 1,000(초)까지 입력 가능
    reqMultiSign.ExpireIn = 1000

    ' 개별문서 등록 - 최대 50 건
    Set tokens = CreateObject("Scripting.Dictionary")
    For i=0 To 2
        Set token = New MultiSignTokens
        ' 서명 원문 유형
        ' TEXT - 일반 텍스트, HASH - HASH 데이터 
        token.tokenType = "TEXT"
        ' 서명 원문 - 원문 2,800자 까지 입력가능
        token.Token = m_NavercertService.encrypt("전자서명(복수) 요청 원문 "+CStr(i))
        ' 서명 원문 유형
        ' token.tokenType = "HASH"
        ' 서명 원문 유형이 HASH인 경우, 원문은 SHA-256, Base64 URL Safe No Padding을 사용
        ' token.Token = m_NavercertService.encrypt(m_NavercertService.sha256_base64url("전자서명(복수) 요청 원문 "+CStr(i)))
        reqMultiSign.addToken i, token
    Next
    
    ' AppToApp 인증요청 여부
    ' true - AppToApp 인증방식, false - 푸시(Push) 인증방식
    reqMultiSign.AppUseYN = false
    ' AppToApp 인증방식에서 사용
    ' 모바일장비 유형('ANDROID', 'IOS'), 대문자 입력(대소문자 구분)
    ' reqMultiSign.DeviceOSType = "ANDROID";
    ' AppToApp 방식 이용시, 호출할 URL
    ' "http", "https"등의 웹프로토콜 사용 불가
    ' reqMultiSign.ReturnURL = "navercert://sign";

    On Error Resume Next

        Dim result : Set result = m_NavercertService.RequestMultiSign(clientCode, reqMultiSign)

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
                <legend>네이버 전자서명(복수) 요청</legend>
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