<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert ASP Example</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' ���̹� �̿��ڿ��� ����(�ִ� 50��) ������ ���ڼ����� ��û�մϴ�.
    ' https://developers.barocert.com/reference/naver/asp/sign/api-multi#RequestMultiSign
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023090000021"
    
    ' ���ڼ��� ��û ���� ��ü
    Dim reqMultiSign : Set reqMultiSign = new MultiSign
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqMultiSign.ReceiverHP = m_NavercertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqMultiSign.ReceiverName = m_NavercertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqMultiSign.ReceiverBirthday = m_NavercertService.encrypt("19700101")
    ' ������û �޽��� ���� - �ִ� 40��
    reqMultiSign.ReqTitle = "���ڼ���(����) ��û �޽��� ����"
    ' ������ ����ó - �ִ� 12��
    reqMultiSign.CallCenterNum = "1600-9854"
    ' ������û �޽��� - �ִ� 500��
    reqMultiSign.ReqMessage = m_NavercertService.encrypt("���ڼ���(����) ��û �޽���")
    ' ������û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqMultiSign.ExpireIn = 1000

    ' �������� ��� - �ִ� 50 ��
    Set tokens = CreateObject("Scripting.Dictionary")
    For i=0 To 2
        Set token = New MultiSignTokens
        ' ���� ���� ����
        ' TEXT - �Ϲ� �ؽ�Ʈ, HASH - HASH ������ 
        token.tokenType = "TEXT"
        ' ���� ���� - ���� 2,800�� ���� �Է°���
        token.Token = m_NavercertService.encrypt("���ڼ���(����) ��û ���� "+CStr(i))
        ' ���� ���� ����
        ' token.tokenType = "HASH"
        ' ���� ���� ������ HASH�� ���, ������ SHA-256, Base64 URL Safe No Padding�� ���
        ' token.Token = m_NavercertService.encrypt(m_NavercertService.sha256_base64url("���ڼ���(����) ��û ���� "+CStr(i)))
        reqMultiSign.addToken i, token
    Next
    
    ' AppToApp ������û ����
    ' true - AppToApp �������, false - Ǫ��(Push) �������
    reqMultiSign.AppUseYN = false
    ' AppToApp ������Ŀ��� ���
    ' �������� ����('ANDROID', 'IOS'), �빮�� �Է�(��ҹ��� ����)
    ' reqMultiSign.DeviceOSType = "ANDROID";
    ' AppToApp ��� �̿��, ȣ���� URL
    ' "http", "https"���� ���������� ��� �Ұ�
    ' reqMultiSign.ReturnURL = "navercert://sign";

    On Error Resume Next

        Dim result : Set result = m_NavercertService.RequestMultiSign(clientCode, reqMultiSign)

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
                <legend>���̹� ���ڼ���(����) ��û</legend>
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