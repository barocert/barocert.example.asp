<!--#include virtual="Barocert/Barocert.asp"-->
<!--#include virtual="Barocert/Passcert.asp"-->
<%
	'**************************************************************
	' Passcert API ASP SDK Example
	'
	' - 업데이트 일자 : 2023-12-11
	' - 연동 기술지원 연락처 : 1600-9854
	' - 연동 기술지원 이메일 : code@linkhubcorp.com
	'
	' <테스트 연동개발 준비사항>
	' 링크아이디(LinkID)와 비밀키(SecretKey)를 메일로 발급받은 인증정보를 참조하여 변경합니다.
	'**************************************************************

	' 링크아이디 
	Dim LinkID : LinkID = "TESTER"
	
	' 비밀키
	Dim SecretKey : SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="
	
	Dim m_PasscertService : set m_PasscertService = New PasscertService
	
	' Passcert API 서비스 모듈 초기화
	m_PasscertService.Initialize LinkID, SecretKey

	' 인증토큰 IP제한기능 사용여부, True-사용, False-미사용, 기본값(True)
	m_PasscertService.IPRestrictOnOff = True

	' 패스써트 API 서비스 고정 IP 사용여부, True-사용, False-미사용, 기본값(False)
	m_PasscertService.useStaticIP = False
%>