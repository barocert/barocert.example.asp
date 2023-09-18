<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' �н� �̿��ڿ��� ����α����� ��û�մϴ�.
    ' https://developers.barocert.com/reference/pass/asp/login/api#RequestLogin
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023040000001"
    
    ' ����α��� ��û ���� ��ü
    Dim reqLogin : Set reqLogin = new Login
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqLogin.ReceiverHP = m_PasscertService.encrypt("01067668440")
    ' ������ ���� - 80��
    reqLogin.ReceiverName = m_PasscertService.encrypt("���켮")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqLogin.ReceiverBirthday = m_PasscertService.encrypt("19900911")
    ' ��û �޽��� ���� - �ִ� 40��
    reqLogin.ReqTitle = "����α��� �޽��� �����"
    ' ��û �޽��� - �ִ� 500��
    reqLogin.ReqMessage = m_PasscertService.encrypt("����α��� ��û �޽��� ����")
    ' ������ ����ó - �ִ� 12��
    reqLogin.CallCenterNum = "1600-9854"
    ' ��û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqLogin.ExpireIn = 1000
    ' ���� ���� - ���� 2,800�� ���� �Է°��� 
    reqLogin.Token = m_PasscertService.encrypt("����α��� ��û ��ū")
    ' ����� ���� �ʿ� ����
    reqLogin.UserAgreementYN = true
    ' ����� ���� ���� ����
    reqLogin.ReceiverInfoYN = true
    ' AppToApp ��û ����
    ' true - AppToApp �������, false - Push �������
    reqCms.AppUseYN = false
    ' ApptoApp ������Ŀ��� ���
    ' ��Ż� ����('SKT', 'KT', 'LGU'), �빮�� �Է�(��ҹ��� ����)
    ' reqCms.TelcoType = 'SKT'
    ' ApptoApp ������Ŀ��� ���
    ' �������� ����('ANDROID', 'IOS'), �빮�� �Է�(��ҹ��� ����)
    ' reqCms.DeviceOSType = 'IOS'

    On Error Resume Next

        Dim result : Set result = m_PasscertService.RequestLogin(clientCode, reqLogin)

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
                <legend>�н� ����α��� ��û</legend>
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