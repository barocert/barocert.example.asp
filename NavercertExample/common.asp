<!--#include virtual="Barocert/Barocert.asp"-->
<!--#include virtual="Barocert/Navercert.asp"-->
<%
	'**************************************************************
	' Barocert NAVER API ASP SDK Example
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
	
	Dim m_NavercertService : set m_NavercertService = New NavercertService
	
	' Navercert API ���� ��� �ʱ�ȭ
	m_NavercertService.Initialize LinkID, SecretKey

	' ������ū IP ���� ����, true-���, false-�̻��, (�⺻��:true)
	m_NavercertService.IPRestrictOnOff = True

	' ��� IP ����, true-���, false-�̻��, (�⺻��:false)
	m_NavercertService.useStaticIP = False
%>