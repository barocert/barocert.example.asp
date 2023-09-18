<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' 패스 이용자에게 간편로그인을 요청합니다.
    ' https://developers.barocert.com/reference/pass/asp/login/api#RequestLogin
    '**************************************************************

    ' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
    Dim clientCode : clientCode = "023040000001"
    
    ' 간편로그인 요청 정보 객체
    Dim reqLogin : Set reqLogin = new Login
    ' 수신자 휴대폰번호 - 11자 (하이픈 제외)
    reqLogin.ReceiverHP = m_PasscertService.encrypt("01067668440")
    ' 수신자 성명 - 80자
    reqLogin.ReceiverName = m_PasscertService.encrypt("정우석")
    ' 수신자 생년월일 - 8자 (yyyyMMdd)
    reqLogin.ReceiverBirthday = m_PasscertService.encrypt("19900911")
    ' 요청 메시지 제목 - 최대 40자
    reqLogin.ReqTitle = "간편로그인 메시지 제목란"
    ' 요청 메시지 - 최대 500자
    reqLogin.ReqMessage = m_PasscertService.encrypt("간편로그인 요청 메시지 내용")
    ' 고객센터 연락처 - 최대 12자
    reqLogin.CallCenterNum = "1600-9854"
    ' 요청 만료시간 - 최대 1,000(초)까지 입력 가능
    reqLogin.ExpireIn = 1000
    ' 서명 원문 - 원문 2,800자 까지 입력가능 
    reqLogin.Token = m_PasscertService.encrypt("간편로그인 요청 토큰")
    ' 사용자 동의 필요 여부
    reqLogin.UserAgreementYN = true
    ' 사용자 정보 포함 여부
    reqLogin.ReceiverInfoYN = true
    ' AppToApp 요청 여부
    ' true - AppToApp 인증방식, false - Push 인증방식
    reqCms.AppUseYN = false
    ' ApptoApp 인증방식에서 사용
    ' 통신사 유형('SKT', 'KT', 'LGU'), 대문자 입력(대소문자 구분)
    ' reqCms.TelcoType = 'SKT'
    ' ApptoApp 인증방식에서 사용
    ' 모바일장비 유형('ANDROID', 'IOS'), 대문자 입력(대소문자 구분)
    ' reqCms.DeviceOSType = 'IOS'

    On Error Resume Next

        Dim result : Set result = m_PasscertService.RequestLogin(clientCode, reqLogin)

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
                <legend>패스 간편로그인 요청</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>접수아이디 (ReceiptID) : <%=result.receiptId %></li>
                        <li>앱스킴 (scheme) : <%=result.scheme %></li>
                        <li>앱다운로드URL (MarketUrl) : <%=result.marketUrl %></li>
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