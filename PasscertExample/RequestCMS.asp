<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' 패스 이용자에게 자동이체 출금동의를 요청합니다.
    ' https://developers.barocert.com/reference/pass/asp/cms/api#RequestCMS
    '**************************************************************

    ' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
    Dim clientCode : clientCode = "023070000014"    
    
    ' 출금동의 요청 정보 객체
    Dim reqCms : Set reqCms = New CMS
    ' 수신자 휴대폰번호 - 11자 (하이픈 제외)
    reqCms.ReceiverHP = m_PasscertService.encrypt("01012341234")
    ' 수신자 성명 - 80자
    reqCms.ReceiverName = m_PasscertService.encrypt("홍길동")
    ' 수신자 생년월일 - 8자 (yyyyMMdd)
    reqCms.ReceiverBirthday = m_PasscertService.encrypt("19700101")
    ' 요청 메시지 제목 - 최대 40자
    reqCms.ReqTitle = "출금동의 요청 메시지 제목"
    ' 요청 메시지 - 최대 500자
    reqCms.ReqMessage = m_PasscertService.encrypt("출금동의 요청 메시지")
    ' 고객센터 연락처 - 최대 12자
    reqCms.CallCenterNum = "1600-9854"
    ' 요청 만료시간 - 최대 1,000(초)까지 입력 가능
    reqCms.ExpireIn = 1000
    ' 사용자 동의 필요 여부
    reqCms.UserAgreementYN = true
    ' 사용자 정보 포함 여부
    reqCms.ReceiverInfoYN = true
    ' 출금은행명 - 최대 100자
    reqCms.BankName = m_PasscertService.encrypt("국민은행")
    ' 출금계좌번호 - 최대 31자
    reqCms.BankAccountNum = m_PasscertService.encrypt("9-****-5117-58")
    ' 출금계좌 예금주명 - 최대 100자
    reqCms.BankAccountName = m_PasscertService.encrypt("홍길동")
    ' 출금유형
    ' CMS - 출금동의, OPEN_BANK - 오픈뱅킹
    reqCms.BankServiceType = m_PasscertService.encrypt("CMS")
    ' 출금액
    reqCms.BankWithdraw = m_PasscertService.encrypt("1,000,000원")
    ' AppToApp 요청 여부
    ' true - AppToApp 인증방식, false - 푸시(Push) 인증방식
    reqCms.AppUseYN = false
    ' ApptoApp 인증방식에서 사용
    ' 통신사 유형('SKT', 'KT', 'LGU'), 대문자 입력(대소문자 구분)
    ' reqCms.TelcoType = 'SKT'
    ' ApptoApp 인증방식에서 사용
    ' 모바일장비 유형('ANDROID', 'IOS'), 대문자 입력(대소문자 구분)
    ' reqCms.DeviceOSType = 'IOS'
    
    On Error Resume Next

    Dim result : Set result = m_PasscertService.RequestCMS(clientCode, reqCms)

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
                <legend>패스 출금동의 요청</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>접수아이디 (ReceiptID) : <%=result.receiptID %></li>
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