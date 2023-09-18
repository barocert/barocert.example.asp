<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' 완료된 전자서명을 검증하고 전자서명값(signedData)을 반환 받습니다.
    ' 반환받은 전자서명값(signedData)과 [1. RequestIdentity] 함수 호출에 입력한 Token의 동일 여부를 확인하여 이용자의 본인인증 검증을 완료합니다.
    ' 검증 함수는 본인인증 요청 함수를 호출한 당일 23시 59분 59초까지만 호출 가능합니다.
    ' 본인인증 요청 함수를 호출한 당일 23시 59분 59초 이후 검증 함수를 호출할 경우 오류가 반환됩니다.
    ' https://developers.barocert.com/reference/pass/asp/identity/api#VerifyIdentity
    '**************************************************************

    ' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
    Dim clientCode : clientCode = "023040000001"

    ' 본인인증 요청시 반환된 접수아이디
    Dim receiptID : receiptID = "02309180230700000140000000000008"

    Dim verifyIdentity : Set verifyIdentity = new IdentityVerify

	verifyIdentity.ReceiverHP = m_PasscertService.encrypt("01012341234")
	
	verifyIdentity.ReceiverName = m_PasscertService.encrypt("홍길동")

    On Error Resume Next

        Dim result : Set result = m_PasscertService.VerifyIdentity(clientCode, receiptID, verifyIdentity)

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
                <legend>패스 본인인증 검증</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>접수 아이디 (ReceiptID) : <%=result.receiptID %></li>
						<li>상태 (State) : <%=result.state %></li>
						<li>수신자 성명 (ReceiverName) : <%=result.receiverName %></li>
						<li>수신자 출생년도 (ReceiverYear) : <%=result.receiverYear %></li>
						<li>수신자 출생월일 (ReceiverDay) : <%=result.receiverDay %></li>
						<li>수신자 휴대폰번호 (ReceiverGender) : <%=result.receiverGender %></li>
						<li>외국인 여부 (ReceiverForeign) : <%=result.receiverForeign %></li>
						<li>통신사 유형 (ReceiverTelcoType) : <%=result.receiverTelcoType %></li>
						<li>전자서명 데이터 전문 (SignedData) : <%=result.signedData %></li>
						<li>연계정보 (Ci) : <%=result.ci %></li>
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