<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' 카카오톡 이용자에게 자동이체 출금동의를 요청합니다.
    ' https://developers.barocert.com/reference/kakao/asp/cms/api#RequestCMS
    '**************************************************************

    ' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
    Dim clientCode : clientCode = "023040000001"    
    
    ' 출금동의 요청 정보 객체
    Dim reqCms : Set reqCms = New CMS
    ' 수신자 휴대폰번호 - 11자 (하이픈 제외)
    reqCms.ReceiverHP = m_KakaocertService.encrypt("01067668440")
    ' 수신자 성명 - 80자
    reqCms.ReceiverName = m_KakaocertService.encrypt("정우석")
    ' 수신자 생년월일 - 8자 (yyyyMMdd)
    reqCms.ReceiverBirthday = m_KakaocertService.encrypt("19900911")
    ' 인증요청 메시지 제목 - 최대 40자
    reqCms.ReqTitle = "인증요청 메시지 제공란"
    ' 인증요청 만료시간 - 최대 1,000(초)까지 입력 가능
    reqCms.ExpireIn = 1000
    ' 청구기관명 - 최대 100자
    reqCms.RequestCorp = m_KakaocertService.encrypt("청구기관명란")
    ' 출금은행명 - 최대 100자
    reqCms.BankName = m_KakaocertService.encrypt("출금은행명란")
    ' 출금계좌번호 - 최대 32자
    reqCms.BankAccountNum = m_KakaocertService.encrypt("9-4324-5117-58")
    ' 출금계좌 예금주명 - 최대 100자
    reqCms.BankAccountName = m_KakaocertService.encrypt("예금주명 입력란")
    ' 출금계좌 예금주 생년월일 - 8자
    reqCms.BankAccountBirthday = m_KakaocertService.encrypt("19930112")
    ' 출금유형
    ' CMS - 출금동의용, FIRM - 펌뱅킹, GIRO - 지로용
    reqCms.BankServiceType = m_KakaocertService.encrypt("CMS")
    ' AppToApp 인증요청 여부
    ' true - AppToApp 인증방식, false - Talk Message 인증방식
    reqCms.AppUseYN = false
    ' App to App 방식 이용시, 에러시 호출할 URL
    ' reqCms.ReturnURL("https://www.kakaocert.com")

    On Error Resume Next

    Dim result : Set result = m_KakaocertService.RequestCMS(clientCode, reqCms)

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
    <legend>카카오 출금동의 요청</legend>
    <% If code = 0 Then %>
        <ul>
            <li>접수아이디 (ReceiptID) : <%=result.receiptID %></li>
            <li>앱스킴 (scheme) : <%=result.scheme %></li>
        </ul>
    <%    Else  %>
        <ul>
            <li>Response.code: <%=code%> </li>
            <li>Response.message: <%=message%> </li>
        </ul>    
    <%    End If    %>
    </fieldset>
    </div>
    </body>
</html>