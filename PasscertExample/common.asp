<!--#include virtual="Barocert/Barocert.asp"-->
<!--#include virtual="Barocert/Passcert.asp"-->
<%
	'**************************************************************
	' Passcert API ASP SDK Example
	'
	' - ������Ʈ ���� : 2023-12-11
	' - ���� ������� ����ó : 1600-9854
	' - ���� ������� �̸��� : code@linkhubcorp.com
	'
	' <�׽�Ʈ �������� �غ����>
	' ��ũ���̵�(LinkID)�� ���Ű(SecretKey)�� ���Ϸ� �߱޹��� ���������� �����Ͽ� �����մϴ�.
	'**************************************************************

	' ��ũ���̵� 
	Dim LinkID : LinkID = "TESTER"
	
	' ���Ű
	Dim SecretKey : SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="
	
	Dim m_PasscertService : set m_PasscertService = New PasscertService
	
	' Passcert API ���� ��� �ʱ�ȭ
	m_PasscertService.Initialize LinkID, SecretKey

	' ������ū IP���ѱ�� ��뿩��, True-���, False-�̻��, �⺻��(True)
	m_PasscertService.IPRestrictOnOff = True

	' �н���Ʈ API ���� ���� IP ��뿩��, True-���, False-�̻��, �⺻��(False)
	m_PasscertService.useStaticIP = False
%>