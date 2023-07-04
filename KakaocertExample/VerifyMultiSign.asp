<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>Barocert SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 

<%
	'**************************************************************
	' 전자서명 요청시 반환된 접수아이디를 통해 서명을 검증합니다. (단건)
	' 검증하기 API는 완료된 전자서명 요청당 1회만 요청 가능하며, 사용자가 서명을 완료후 유효시간(10분)이내에만 요청가능 합니다.
	'**************************************************************

	' 이용기관코드, 파트너가 등록한 이용기관의 코드, (파트너 사이트에서 확인가능)
	Dim clientCode : clientCode = "023040000001"	

	' 전자서명 요청시 반환된 접수아이디
	Dim receiptID : receiptID = "02307040230400000010000000000027"

	On Error Resume Next

		Dim result : Set result = m_KakaocertService.VerifyMultiSign(clientCode, receiptID)

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
				<legend>카카오 전자서명 검증(복수)</legend>
				<% If code = 0 Then %>
					<ul>
						<li>접수 아이디 (ReceiptID) : <%=result.receiptID %></li>
						<li>상태 (State) : <%=result.state %></li>
						<li>연계정보 (Ci) : <%=result.ci %></li>
					</ul>
					<%
						For i=0 To UBound(result.multiSignedData) -1
					%>
						<fieldset class="fieldset2">
							<ul>
								<li>전자서명 데이터 전문 (SignedData) : <%=result.multiSignedData(i)%></li>
							</ul>
						</fieldset>
					<%
						Next
					Else
					%>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>