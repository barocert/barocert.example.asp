<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>Barocert SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 

<%
    '**************************************************************
    ' īī���� �̿��ڿ��� �ڵ���ü ��ݵ��Ǹ� ��û�մϴ�.
    ' https://developers.barocert.com/reference/kakao/asp/cms/api#RequestCMS
    '**************************************************************

    ' �̿����ڵ�, ��Ʈ�ʰ� ����� �̿����� �ڵ� (��Ʈ�� ����Ʈ���� Ȯ�ΰ���)
    Dim clientCode : clientCode = "023040000001"    
    
    ' ��ݵ��� ��û ���� ��ü
    Dim reqCms : Set reqCms = New CMS
    ' ������ �޴�����ȣ - 11�� (������ ����)
    reqCms.ReceiverHP = m_KakaocertService.encrypt("01067668440")
    ' ������ ���� - 80��
    reqCms.ReceiverName = m_KakaocertService.encrypt("���켮")
    ' ������ ������� - 8�� (yyyyMMdd)
    reqCms.ReceiverBirthday = m_KakaocertService.encrypt("19900911")
    ' ������û �޽��� ���� - �ִ� 40��
    reqCms.ReqTitle = "������û �޽��� ������"
    ' ������û ����ð� - �ִ� 1,000(��)���� �Է� ����
    reqCms.ExpireIn = 1000
    ' û������� - �ִ� 100��
    reqCms.RequestCorp = m_KakaocertService.encrypt("û��������")
    ' �������� - �ִ� 100��
    reqCms.BankName = m_KakaocertService.encrypt("���������")
    ' ��ݰ��¹�ȣ - �ִ� 32��
    reqCms.BankAccountNum = m_KakaocertService.encrypt("9-4324-5117-58")
    ' ��ݰ��� �����ָ� - �ִ� 100��
    reqCms.BankAccountName = m_KakaocertService.encrypt("�����ָ� �Է¶�")
    ' ��ݰ��� ������ ������� - 8��
    reqCms.BankAccountBirthday = m_KakaocertService.encrypt("19930112")
    ' �������
    ' CMS - ��ݵ��ǿ�, FIRM - �߹�ŷ, GIRO - ���ο�
    reqCms.BankServiceType = m_KakaocertService.encrypt("CMS")
    ' AppToApp ������û ����
    ' true - AppToApp �������, false - Talk Message �������
    reqCms.AppUseYN = false
    ' App to App ��� �̿��, ������ ȣ���� URL
    ' reqCms.ReturnURL("https://www.kakaocert.com")

    On Error Resume Next

    Dim result : Set result = m_KakaocertService.RequestCMS(clientCode, reqCms)

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
    <legend>īī�� ��ݵ��� ��û</legend>
    <% If code = 0 Then %>
        <ul>
            <li>�������̵� (ReceiptID) : <%=result.receiptID %></li>
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