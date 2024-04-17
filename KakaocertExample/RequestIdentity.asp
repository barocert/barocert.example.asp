<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert ASP Example</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' īī���� �̿��ڿ��� ���������� ��û�մϴ�.
    ' https://developers.barocert.com/reference/kakao/asp/identity/api#RequestIdentity
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023040000001"        
    
    ' �������� ��û ���� ��ü
    Dim reqIdentity : Set reqIdentity = new Identity
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqIdentity.ReceiverHP = m_KakaocertService.encrypt("01012341234")
    ' ������ ���� - 80��
    reqIdentity.ReceiverName = m_KakaocertService.encrypt("ȫ�浿")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqIdentity.ReceiverBirthday = m_KakaocertService.encrypt("19700101")
    ' ������û �޽��� ���� - �ִ� 40��
    reqIdentity.ReqTitle = "�������� ��û �޽��� ����"
    ' Ŀ���� �޽��� - �ִ� 500��
    reqIdentity.ExtraMessage = m_KakaocertService.encrypt("�������� Ŀ���� �޽���")
    ' ������û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqIdentity.ExpireIn = 1000
    ' ���� ���� - �ִ� 40�� ���� �Է°���
    reqIdentity.Token = m_KakaocertService.encrypt("�������� ��û ����")
    ' AppToApp ������û ����
    ' true - AppToApp �������, false - Talk Message �������
    reqIdentity.AppUseYN = false
    ' App to App ��� �̿��, ȣ���� URL
    ' reqIdentity.ReturnURL = "https://www.kakaocert.com"

    On Error Resume Next

        Dim result : Set result = m_KakaocertService.RequestIdentity(clientCode, reqIdentity)

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
                <legend>īī�� �������� ��û</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
                        <li>�۽�Ŵ (scheme) : <%=result.scheme %></li>
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