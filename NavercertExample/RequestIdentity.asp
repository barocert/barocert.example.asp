<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' ���̹� �̿��ڿ��� ���������� ��û�մϴ�.
    ' https://developers.barocert.com/reference/naver/asp/identity/api#RequestIdentity
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023090000021"        
    
    ' �������� ��û ���� ��ü
    Dim reqIdentity : Set reqIdentity = new Identity
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqIdentity.ReceiverHP = m_NavercertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqIdentity.ReceiverName = m_NavercertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqIdentity.ReceiverBirthday = m_NavercertService.encrypt("19700101")
    ' ������ ����ó - �ִ� 12��
    reqIdentity.CallCenterNum = "1600-9854"
    ' ������û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqIdentity.ExpireIn = 1000
    ' AppToApp ������û ����
    ' true - AppToApp �������, false - Talk Message �������
    reqIdentity.AppUseYN = false
    ' AppToApp ������Ŀ��� ���
    ' �������� ����('ANDROID', 'IOS'), �빮�� �Է�(��ҹ��� ����)
    ' reqIdentity.DeviceOSType = "ANDROID";
    ' AppToApp ��� �̿��, ȣ���� URL
    ' "http", "https"���� ���������� ��� �Ұ�
    ' reqIdentity.ReturnURL = "navercert://Identity";

    On Error Resume Next

        Dim result : Set result = m_NavercertService.RequestIdentity(clientCode, reqIdentity)

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
                <legend>���̹� �������� ��û</legend>
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