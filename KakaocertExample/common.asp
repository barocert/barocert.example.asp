<!--#include virtual="Barocert/Barocert.asp"-->
<!--#include virtual="Barocert/Kakaocert.asp"-->
<%
	'**************************************************************
	' Barocert KAKAO API ASP SDK Example
	'
	' ������Ʈ ���� : 2024-04-17
	' ����������� ����ó : 1600-9854
	' ����������� �̸��� : code@linkhubcorp.com
	'         
	' <�׽�Ʈ �������� �غ����>
	'   1) API Key ���� (������û �� ���Ϸ� ���޵� ����)
	'       - LinkID : ��ũ��꿡�� �߱��� ��ũ���̵�
	'       - SecretKey : ��ũ��꿡�� �߱��� ���Ű
	'   2) SDK ȯ�漳�� �ʼ� �ɼ� ����
	'       - IPRestrictOnOff : ������ū IP ���� ����, true-���, false-�̻��, (�⺻��:true)
	'       - UseStaticIP : ��� IP ����, true-���, false-�̻��, (�⺻��:false)
	'**************************************************************

	' ��ũ���̵� 
	Dim LinkID : LinkID = "TESTER"
	
	' ���Ű
	Dim SecretKey : SecretKey = "SwWxqU+0TErBXy/9TVjIPEnI0VTUMMSQZtJf3Ed8q3I="
	
	Dim m_KakaocertService : set m_KakaocertService = New KakaocertService
	
	' Kakaocert API ���� ��� �ʱ�ȭ
	m_KakaocertService.Initialize LinkID, SecretKey

	' ������ū IP ���� ����, true-���, false-�̻��, (�⺻��:true)
	m_KakaocertService.IPRestrictOnOff = True

	' UseStaticIP : ��� IP ����, true-���, false-�̻��, (�⺻��:false)
	m_KakaocertService.useStaticIP = False
%>