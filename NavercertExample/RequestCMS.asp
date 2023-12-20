<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' ���̹� �̿��ڿ��� �ڵ���ü ��ݵ��Ǹ� ��û�մϴ�.
    ' https://developers.barocert.com/reference/naver/asp/cms/api#RequestCMS
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023090000021"        
    
    ' ��ݵ��� ��û ���� ��ü
    Dim reqCMS : Set reqCMS = new CMS
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqCMS.ReceiverHP = m_NavercertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqCMS.ReceiverName = m_NavercertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqCMS.ReceiverBirthday = m_NavercertService.encrypt("19700101")
    ' ������û �޽��� ���� - �ִ� 40��
    reqCMS.ReqTitle = "��ݵ��� ��û �޽��� ����"
    ' ������û �޽��� - �ִ� 500��
    reqCMS.ReqMessage = m_NavercertService.encrypt("��ݵ��� ��û �޽���")
    ' ������ ����ó - �ִ� 12��
    reqCMS.CallCenterNum = "1600-9854"
    ' ������û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqCMS.ExpireIn = 1000
    ' û�������
    reqCMS.requestCorp = m_NavercertService.encrypt("û�����")
    ' ��������
    reqCMS.bankName = m_NavercertService.encrypt("�������")
    ' ��ݰ��¹�ȣ
    reqCMS.bankAccountNum = m_NavercertService.encrypt("123-456-7890")
    ' ��ݰ��� �����ָ�
    reqCMS.bankAccountName = m_NavercertService.encrypt("ȫ�浿")
    ' ��ݰ��� ������ �������
    reqCMS.bankAccountBirthday = m_NavercertService.encrypt("19700101")
    ' AppToApp ������û ����
    ' true - AppToApp �������, false - Ǫ��(Push) �������
    reqCMS.AppUseYN = false
    ' AppToApp ������Ŀ��� ���
    ' �������� ����('ANDROID', 'IOS'), �빮�� �Է�(��ҹ��� ����)
    ' reqCMS.DeviceOSType = "ANDROID";
    ' AppToApp ��� �̿��, ȣ���� URL
    ' "http", "https"���� ���������� ��� �Ұ�
    ' reqCMS.ReturnURL = "navercert://cms";

    On Error Resume Next

        Dim result : Set result = m_NavercertService.RequestCMS(clientCode, reqCMS)

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
                <legend>���̹� ��ݵ��� ��û</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
                        <li>�۽�Ŵ (scheme) : <%=result.scheme %></li>
                        <li>�۴ٿ�ε�URL (marketUrl) : <%=result.marketUrl %></li>
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