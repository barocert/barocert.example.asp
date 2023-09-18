<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' �н� �̿��ڿ��� ������ ���ڼ����� ��û�մϴ�.
    ' https://developers.barocert.com/reference/pass/asp/sign/api#RequestSign
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023040000001"
    
    ' ���ڼ��� ��û ���� ��ü
    Dim reqSign : Set reqSign = new Sign
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqSign.ReceiverHP = m_PasscertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqSign.ReceiverName = m_PasscertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqSign.ReceiverBirthday = m_PasscertService.encrypt("19700101")
    ' ��û �޽��� ���� - �ִ� 40��
    reqSign.ReqTitle = "���ڼ��� �޽��� �����"
    ' ��û �޽��� - �ִ� 500��
    reqSign.ReqMessage = m_PasscertService.encrypt("���ڼ��� ��û �޽��� ����")
    ' ������ ����ó - �ִ� 12��
    reqSign.CallCenterNum = "1600-9854"
    ' ��û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqSign.ExpireIn = 1000
    ' ���� ���� - ���� 2,800�� ���� �Է°��� 
    reqSign.Token = m_PasscertService.encrypt("���ڼ��� ��û ��ū")
    ' ���� ���� ����
    ' 'TEXT' - �Ϲ� �ؽ�Ʈ, 'HASH' - HASH ������, 'URL' - URL ������
    ' ����������(originalTypeCode, originalURL, originalFormatCode) �Է½� 'TEXT'��� �Ұ�
    reqSign.tokenType = "URL"
    ' ����� ���� �ʿ� ����
    reqSign.UserAgreementYN = true
    ' ����� ���� ���� ����
    reqSign.ReceiverInfoYN = true
    ' ���������ڵ�
    ' 'AG' - ���Ǽ�, 'AP' - ��û��, 'CT' - ��༭, 'GD' - �ȳ���, 'NT' - ������, 'TR' - ���
    reqSign.originalTypeCode = "TR"
    ' ������ȸURL
    reqSign.originalURL = "https://www.passcert.co.kr"
    ' ���������ڵ�
    ' ('TEXT', 'HTML', 'DOWNLOAD_IMAGE', 'DOWNLOAD_DOCUMENT')
    reqSign.originalFormatCode = "HTML"
    ' AppToApp ��û ����
    ' true - AppToApp �������, false - Push �������
    reqSign.AppUseYN = false
    ' ApptoApp ������Ŀ��� ���
    ' ��Ż� ����('SKT', 'KT', 'LGU'), �빮�� �Է�(��ҹ��� ����)
    ' reqSign.TelcoType = "SKT"
    ' ApptoApp ������Ŀ��� ���
    ' �������� ����('ANDROID', 'IOS'), �빮�� �Է�(��ҹ��� ����)
    ' reqSign.DeviceOSType = "IOS"

    On Error Resume Next

        Dim result : Set result = m_PasscertService.RequestSign(clientCode, reqSign)

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
                <legend>�н� ���ڼ��� ��û</legend>
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