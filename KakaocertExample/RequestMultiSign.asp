<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert ASP Example</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' 카카오톡 이용자에게 복수(최대 20건) 문서의 전자서명을 요청합니다.
    ' https://developers.barocert.com/reference/kakao/asp/sign/api-multi#RequestMultiSign
    '**************************************************************

    ' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
    Dim clientCode : clientCode = "023040000001"
    
    ' 전자서명 요청 정보 객체
    Dim reqMultiSign : Set reqMultiSign = new MultiSign
    ' 수신자 휴대폰번호 - 11자 (하이픈 제외)
    reqMultiSign.ReceiverHP = m_KakaocertService.encrypt("01012341234")
    ' 수신자 성명 - 80자
    reqMultiSign.ReceiverName = m_KakaocertService.encrypt("홍길동")
    ' 수신자 생년월일 - 8자 (yyyyMMdd)
    reqMultiSign.ReceiverBirthday = m_KakaocertService.encrypt("19700101")
    ' 인증요청 메시지 제목 - 최대 40자
    reqMultiSign.ReqTitle = "전자서명(복수) 요청 메시지 제목"
    ' 커스텀 메시지 - 최대 500자
    reqMultiSign.ExtraMessage = m_KakaocertService.encrypt("전자서명(복수) 커스텀 메시지")
    ' 인증요청 만료시간 - 최대 1,000(초)까지 입력 가능
    reqMultiSign.ExpireIn = 1000

    ' 개별문서 등록 - 최대 20 건
    Set tokens = CreateObject("Scripting.Dictionary")
    For i=0 To 2
        Set token = New MultiSignTokens
        ' 서명 요청 제목 - 최대 40자
        token.SignTitle = "전자서명(복수) 서명 요청 제목 " + CStr(i)
        ' 서명 원문 - 원문 2,800자 까지 입력가능
        token.Token = m_KakaocertService.encrypt("전자서명(복수) 요청 원문 "+CStr(i))
        reqMultiSign.addToken i, token
    Next

    ' 서명 원문 유형
    ' TEXT - 일반 텍스트, HASH - HASH 데이터
    reqMultiSign.TokenType = "TEXT"
    ' AppToApp 인증요청 여부
    ' true - AppToApp 인증방식, false - Talk Message 인증방식
    reqMultiSign.AppUseYN = false
    ' App to App 방식 이용시, 에러시 호출할 URL
    ' reqMultiSign.ReturnURL = "https://www.kakaocert.com"


    On Error Resume Next

        Dim result : Set result = m_KakaocertService.RequestMultiSign(clientCode, reqMultiSign)

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
                <legend>카카오 전자서명(복수) 요청</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>접수아이디 (ReceiptID) : <%=result.receiptID %></li>
                        <li>앱스킴 (scheme) : <%=result.scheme %></li>
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