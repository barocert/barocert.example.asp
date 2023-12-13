<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' ���̹� �̿��ڿ��� �ܰ�(1��) ������ ���ڼ����� ��û�մϴ�
    ' https://developers.barocert.com/reference/naver/asp/sign/api-single#RequestSign
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023090000021"
    
    ' ���ڼ��� ��û ���� ��ü
    Dim reqSign : Set reqSign = new Sign
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqSign.ReceiverHP = m_NavercertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqSign.ReceiverName = m_NavercertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqSign.ReceiverBirthday = m_NavercertService.encrypt("19700101")
    ' ������û �޽��� ���� - �ִ� 40��
    reqSign.ReqTitle = "���ڼ���(�ܰ�) ��û �޽��� ����"
    ' ������û �޽��� - �ִ� 500��
    reqSign.ReqMessage = m_NavercertService.encrypt("���ڼ���(�ܰ�) ��û �޽���")
    ' ������ ����ó - �ִ� 12��
    reqSign.CallCenterNum = "1600-9854"
    ' ������û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqSign.ExpireIn = 1000
    ' ���� ���� ����
    ' TEXT - �Ϲ� �ؽ�Ʈ, HASH - HASH ������
    reqSign.TokenType = "TEXT"
    ' ���� ���� - ���� 2,800�� ���� �Է°���
    reqSign.Token = m_NavercertService.encrypt("���ڼ���(�ܰ�) ��û ����")
    ' ���� ���� ����
    ' reqSign.TokenType = "HASH"
    ' ���� ���� ������ HASH�� ���, ������ SHA-256, Base64 URL Safe No Padding�� ���
    ' reqSign.Token = m_NavercertService.encrypt(m_NavercertService.sha256_base64url("���ڼ���(�ܰ�) ��û ����"))
    ' AppToApp ������û ����
    ' true - AppToApp �������, false - Talk Message �������
    reqSign.AppUseYN = false
    ' AppToApp ������Ŀ��� ���
    ' �������� ����('ANDROID', 'IOS'), �빮�� �Է�(��ҹ��� ����)
    ' reqSign.DeviceOSType = "ANDROID";
    ' AppToApp ��� �̿��, ȣ���� URL
    ' "http", "https"���� ���������� ��� �Ұ�
    ' reqSign.ReturnURL = "navercert://sign";

    On Error Resume Next

        Dim result : Set result = m_NavercertService.RequestSign(clientCode, reqSign)

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
                <legend>���̹� ���ڼ���(�ܰ�) ��û</legend>
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