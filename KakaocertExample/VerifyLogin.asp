<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert ASP Example</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' 완료된 전자서명을 검증하고 전자서명 데이터 전문(signedData)을 반환 받습니다.
    ' 카카오 보안정책에 따라 검증 API는 1회만 호출할 수 있습니다. 재시도시 오류가 반환됩니다.
    ' 전자서명 완료일시로부터 10분 이후에 검증 API를 호출하면 오류가 반환됩니다.
    ' https://developers.barocert.com/reference/kakao/asp/login/api#VerifyLogin
    '**************************************************************

    ' 이용기관코드, 파트너가 등록한 이용기관의 코드 (파트너 사이트에서 확인가능)
    Dim clientCode : clientCode = "023040000001"

    ' 간편로그인 요청시 반환된 트랜잭션 아이디
    Dim txID : txID = "018aa84ea3-2a16-4e08-b3c6-07e235aa273f"

    On Error Resume Next

        Dim result : Set result = m_KakaocertService.VerifyLogin(clientCode, txID)

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
                <legend>카카오 간편로그인 검증</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>트랜잭션 아이디 (TxID) : <%=result.txID %></li>
                        <li>상태 (State) : <%=result.state %></li>
                        <li>전자서명 데이터 전문 (SignedData) : <%=result.signedData %></li>
                        <li>연계정보 (Ci) : <%=result.ci %></li>
                        <li>수신자 성명 (ReceiverName) : <%=result.receiverName %></li>
						<li>수신자 출생년도 (ReceiverYear) : <%=result.receiverYear %></li>
						<li>수신자 출생월일 (ReceiverDay) : <%=result.receiverDay %></li>
						<li>수신자 휴대폰번호 (ReceiverHP) : <%=result.receiverHP %></li>
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