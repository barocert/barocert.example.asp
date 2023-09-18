<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' �н� �̿��ڿ��� ���������� ��û�մϴ�.
    ' https://developers.barocert.com/reference/pass/asp/identity/api#RequestIdentity
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023040000001"        
    
    ' �������� ��û ���� ��ü
    Dim reqIdentity : Set reqIdentity = new Identity
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqIdentity.ReceiverHP = m_PasscertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqIdentity.ReceiverName = m_PasscertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqIdentity.ReceiverBirthday = m_PasscertService.encrypt("19700101")
    ' ��û �޽��� ���� - �ִ� 40��
    reqIdentity.ReqTitle = "�������� �޽��� �����"
    ' ��û �޽��� - �ִ� 500��
    reqIdentity.ReqMessage = m_PasscertService.encrypt("�������� ��û �޽��� ����")
    ' ������ ����ó - �ִ� 12��
    reqIdentity.CallCenterNum = "1600-9854"
    ' ��û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqIdentity.ExpireIn = 1000
    ' ���� ���� - ���� 2,800�� ���� �Է°��� 
    reqIdentity.Token = m_PasscertService.encrypt("�������� ��û ��ū")
    ' ����� ���� �ʿ� ����
    reqIdentity.UserAgreementYN = true
    ' ����� ���� ���� ����
    reqIdentity.ReceiverInfoYN = true
    ' AppToApp ��û ����
    ' true - AppToApp �������, false - Push �������
    reqIdentity.AppUseYN = false
    ' ApptoApp ������Ŀ��� ���
    ' ��Ż� ����('SKT', 'KT', 'LGU'), �빮�� �Է�(��ҹ��� ����)
    ' reqIdentity.TelcoType = 'SKT'
    ' ApptoApp ������Ŀ��� ���
    ' �������� ����('ANDROID', 'IOS'), �빮�� �Է�(��ҹ��� ����)
    ' reqIdentity.DeviceOSType = 'IOS'
    
    On Error Resume Next

        Dim result : Set result = m_PasscertService.RequestIdentity(clientCode, reqIdentity)

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
                <legend>�н� �������� ��û</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>�������̵� (ReceiptID) : <%=result.receiptId %></li>
                        <li>�۽�Ŵ (scheme) : <%=result.scheme %></li>
                        <li>�۴ٿ�ε�URL (MarketUrl) : <%=result.marketUrl %></li>
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