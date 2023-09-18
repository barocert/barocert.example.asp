<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' īī���� �̿��ڿ��� �ܰ�(1��) ������ ���ڼ����� ��û�մϴ�.
    ' https://developers.barocert.com/reference/kakao/asp/sign/api-single#RequestSign
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023040000001"
    
    ' ���ڼ��� ��û ���� ��ü
    Dim reqSign : Set reqSign = new Sign
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqSign.ReceiverHP = m_KakaocertService.encrypt("01067668440")
    ' ������ ���� - 80��
    reqSign.ReceiverName = m_KakaocertService.encrypt("���켮")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqSign.ReceiverBirthday = m_KakaocertService.encrypt("19900911")
    ' ������û �޽��� ���� - �ִ� 40��
    reqSign.ReqTitle = "���ڼ���ܰ��׽�Ʈ"
    ' ������û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqSign.ExpireIn = 1000
    ' ���� ���� - ���� 2,800�� ���� �Է°���
    reqSign.Token = m_KakaocertService.encrypt("���ڼ���ܰ��׽�Ʈ������")
    ' ���� ���� ����
    ' TEXT - �Ϲ� �ؽ�Ʈ, HASH - HASH ������
    reqSign.TokenType = "TEXT"
    ' AppToApp ������û ����
    ' true - AppToApp �������, false - Talk Message �������
    reqSign.AppUseYN = false
    ' App to App ��� �̿��, ȣ���� URL
    ' reqSign.ReturnURL = "https://www.kakaocert.com"

    On Error Resume Next

        Dim result : Set result = m_KakaocertService.RequestSign(clientCode, reqSign)

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
                <legend>īī�� ���ڼ���(�ܰ�) ��û</legend>
                <% 
                If code = 0 Then %>
                    <ul>
                        <li>�������̵� (ReceiptID) : <%=result.receiptId %></li>
                        <li>�۽�Ŵ (scheme) : <%=result.scheme %></li>
                    </ul>
                <%    Else  %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>    
                <%    End If    %>
            </fieldset>
        </div>
    </body>
</html>