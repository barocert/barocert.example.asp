<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' �н� �̿��ڿ��� �ڵ���ü ��ݵ��Ǹ� ��û�մϴ�.
    ' https://developers.barocert.com/reference/pass/asp/cms/api#RequestCMS
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023070000014"    
    
    ' ��ݵ��� ��û ���� ��ü
    Dim reqCms : Set reqCms = New CMS
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqCms.ReceiverHP = m_PasscertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqCms.ReceiverName = m_PasscertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqCms.ReceiverBirthday = m_PasscertService.encrypt("19700101")
    ' ��û �޽��� ���� - �ִ� 40��
    reqCms.ReqTitle = "��ݵ��� ��û �޽��� ����"
    ' ��û �޽��� - �ִ� 500��
    reqCms.ReqMessage = m_PasscertService.encrypt("��ݵ��� ��û �޽���")
    ' ������ ����ó - �ִ� 12��
    reqCms.CallCenterNum = "1600-9854"
    ' ��û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqCms.ExpireIn = 1000
    ' ����� ���� �ʿ� ����
    reqCms.UserAgreementYN = true
    ' ����� ���� ���� ����
    reqCms.ReceiverInfoYN = true
    ' �������� - �ִ� 100��
    reqCms.BankName = m_PasscertService.encrypt("��������")
    ' ��ݰ��¹�ȣ - �ִ� 31��
    reqCms.BankAccountNum = m_PasscertService.encrypt("9-****-5117-58")
    ' ��ݰ��� �����ָ� - �ִ� 100��
    reqCms.BankAccountName = m_PasscertService.encrypt("ȫ�浿")
    ' �������
    ' CMS - ��ݵ���, OPEN_BANK - ���¹�ŷ
    reqCms.BankServiceType = m_PasscertService.encrypt("CMS")
    ' ��ݾ�
    reqCms.BankWithdraw = m_PasscertService.encrypt("1,000,000��")
    ' AppToApp ��û ����
    ' true - AppToApp �������, false - Ǫ��(Push) �������
    reqCms.AppUseYN = false
    ' ApptoApp ������Ŀ��� ���
    ' ��Ż� ����('SKT', 'KT', 'LGU'), �빮�� �Է�(��ҹ��� ����)
    ' reqCms.TelcoType = 'SKT'
    ' ApptoApp ������Ŀ��� ���
    ' �������� ����('ANDROID', 'IOS'), �빮�� �Է�(��ҹ��� ����)
    ' reqCms.DeviceOSType = 'IOS'
    
    On Error Resume Next

    Dim result : Set result = m_PasscertService.RequestCMS(clientCode, reqCms)

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
                <legend>�н� ��ݵ��� ��û</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
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