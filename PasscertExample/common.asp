<!--#include virtual="Barocert/Barocert.asp"-->
<!--#include virtual="Barocert/Passcert.asp"-->
<%
	'**************************************************************
	' Barocert PASS API ASP SDK Example
	'
	' 업데이트 일자 : 2024-04-17
	' 연동기술지원 연락처 : 1600-9854
	' 연동기술지원 이메일 : code@linkhubcorp.com
	'         
	' <테스트 연동개발 준비사항>
	'   1) API Key 변경 (연동신청 시 메일로 전달된 정보)
	'       - LinkID : 링크허브에서 발급한 링크아이디
	'       - SecretKey : 링크허브에서 발급한 비밀키
	'   2) SDK 환경설정 필수 옵션 설정
	'       - IPRestrictOnOff : 인증토큰 IP 검증 설정, true-사용, false-미사용, (기본값:true)
	'       - UseStaticIP : 통신 IP 고정, true-사용, false-미사용, (기본값:false)
	'**************************************************************

	' 링크아이디 
	Dim LinkID : LinkID = "TESTER"
	
	' 비밀키
	Dim SecretKey : SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="
	
	Dim m_PasscertService : set m_PasscertService = New PasscertService
	
	' Passcert API 서비스 모듈 초기화
	m_PasscertService.Initialize LinkID, SecretKey

	' 인증토큰 IP 검증 설정, true-사용, false-미사용, (기본값:true)
	m_PasscertService.IPRestrictOnOff = True

	' 통신 IP 고정, true-사용, false-미사용, (기본값:false)
	m_PasscertService.useStaticIP = False
%>