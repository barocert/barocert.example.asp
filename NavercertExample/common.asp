<!--#include virtual="Barocert/Barocert.asp"-->
<!--#include virtual="Barocert/Navercert.asp"-->
<%
	'**************************************************************
	' Navercert API ASP SDK Example
	'
	' - ������Ʈ ���� : 2023-12-13
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
	
	Dim m_NavercertService : set m_NavercertService = New NavercertService
	
	' Navercert API ���� ��� �ʱ�ȭ
	m_NavercertService.Initialize LinkID, SecretKey

	' ������ū IP���ѱ�� ��뿩��, ����(True)
	m_NavercertService.IPRestrictOnOff = True

	' ���̹���Ʈ API ���� ���� IP ��뿩��, True-���, False-�̻��, �⺻��(False)
	m_NavercertService.useStaticIP = False
%>