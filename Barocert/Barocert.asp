<!--#include file="Linkhub/Linkhub.asp"--> 
<!--#include file="Crypto.asp"--> 
<%

Application("LINKHUB_TOKEN_SCOPE_BAROCERT") = Array("partner")
Const ServiceID = "BAROCERT"
Const ServiceURL = "https://barocert.linkhub.co.kr"
Const ServiceURL_Static = "https://static-barocert.linkhub.co.kr"

Const APIVersion = "2.1"
Const adTypeBinary = 1
Const adTypeText = 2

Class BarocertBase

	Private m_TokenDic
	Private m_Encryptor
	Private m_Linkhub
	Private m_IPRestrictOnOff
	Private m_useStaticIP
	Private m_UseLocalTimeYN
	Private m_ServiceURL

	Public Property Let IPRestrictOnOff(ByVal value)
		m_IPRestrictOnOff = value
	End Property
	Public Property Let useStaticIP(ByVal value)
		m_useStaticIP = value
	End Property
	Public Property Let UseLocalTimeYN(ByVal value)
		m_UseLocalTimeYN = value
	End Property
	Public Property Let ServiceURL(ByVal value)
		m_ServiceURL = value
	End Property
	Public Property Let AuthURL(ByVal value)
		m_Linkhub.AuthURL = value
	End Property

	Public Sub Class_Initialize
		On Error Resume next
		If  Not(BAROCERT_TOKEN_CACHE Is Nothing) Then
			Set m_TokenDic = BAROCERT_TOKEN_CACHE
		Else
			Set m_TokenDic = server.CreateObject("Scripting.Dictionary")
		End If
		On Error GoTo 0
		If isEmpty( m_TokenDic) Then
			Set m_TokenDic = server.CreateObject("Scripting.Dictionary")
		End If
		
		m_IPRestrictOnOff = True
		m_UseStaticIP = False
		m_UseLocalTimeYN = True
		Set m_Linkhub = New Linkhub
		Set m_Encryptor = New Encryptor
	End Sub

	Public Sub Initialize(linkID, SecretKey )
		m_Linkhub.LinkID = linkID
		m_Linkhub.SecretKey = SecretKey
		m_Encryptor.Initialize SecretKey
	End Sub

	Public Sub Class_Terminate
		Set m_Linkhub = Nothing 
	End Sub 

	Private Property Get m_scope
		m_scope = Application("LINKHUB_TOKEN_SCOPE_BAROCERT")
	End Property

	Public Sub AddScope(scope)
		t = Application("LINKHUB_TOKEN_SCOPE_BAROCERT")
		ReDim Preserve t(Ubound(t)+1)
		t(Ubound(t)) = scope
		Application("LINKHUB_TOKEN_SCOPE_BAROCERT") = t
	End Sub

	Private Function getTargetURL() 
		If IsNull(m_ServiceURL) or m_ServiceURL = "" Then
			If m_UseStaticIP Then
				getTargetURL = ServiceURL_Static
			Else
				getTargetURL = ServiceURL
			End If
		Else
			If InStr(m_ServiceURL, "https://") = 0 and InStr(m_ServiceURL, "http://") = 0 Then
				Err.raise -99999999, "BAROCERT", "ServiceURL에 전송 프로토콜(HTTP 또는 HTTPS)을 포함하여 주시기 바랍니다."
			Else
				getTargetURL = m_ServiceURL
			End if
		End If
	End Function

	Public Function getSession_token()
		Dim refresh : refresh = False
		Dim m_Token : Set m_Token = Nothing
		If m_TokenDic.Exists(m_Linkhub.LinkID) Then 
			Set m_Token = m_TokenDic.Item(m_Linkhub.LinkID)
		End If
		
		If m_Token Is Nothing Then
			refresh = True
		Else
			'CheckScope
			Dim scope
			For Each scope In m_scope
				If InStr(m_Token.strScope,scope) = 0 Then
					refresh = True
					Exit for
				End if
			Next
			If refresh = False Then
				Dim utcnow : utcnow = CDate(Replace(left(m_linkhub.getTime(m_useStaticIP, m_useLocalTimeYN, false),19),"T" , " " ))
				refresh = CDate(Replace(left(m_Token.expiration,19),"T" , " " )) < utcnow
			End if
		End If
		
		If refresh Then
			If m_TokenDic.Exists(m_Linkhub.LinkID) Then m_TokenDic.remove m_Linkhub.LinkID
			Set m_Token = m_Linkhub.getToken(ServiceID, "", m_scope, IIf(m_IPRestrictOnOff, "", "*"), m_useStaticIP, m_useLocalTimeYN, false)
			m_Token.set "strScope", Join(m_scope,"|")
			m_TokenDic.Add m_Linkhub.LinkID, m_Token
		End If
		
		getSession_token = m_Token.session_token
	End Function


	'Private Functions
	Public Function httpGET(url , BearerToken )
		Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

		Call winhttp1.Open("GET", getTargetURL() + url, false)
		
		Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
		Call winhttp1.setRequestHeader("x-bc-version", APIVersion)
		
		winhttp1.Send
		winhttp1.WaitForResponse
		Dim result : result = winhttp1.responseText

		If winhttp1.Status <> 200 Then
			Set winhttp1 = Nothing
			Dim parsedDic : Set parsedDic = m_Linkhub.parse(result)
			Err.raise parsedDic.code, "BAROCERT", parsedDic.message
		End If
		
		Set winhttp1 = Nothing
		
		Set httpGET = m_Linkhub.parse(result)
	End Function


	Public Function httpPOST(url, BearerToken, postdata)

		Dim winhttp1 : Set winhttp1 = CreateObject("WinHttp.WinHttpRequest.5.1")

		Call winhttp1.Open("POST", getTargetURL() + url)
		Call winhttp1.setRequestHeader("x-bc-version", APIVersion)
		Call winhttp1.setRequestHeader("Content-Type", "Application/json")
		
		If BearerToken <> "" Then
			Call winhttp1.setRequestHeader("Authorization", "Bearer " + BearerToken)
		End If

		Dim xDate : xDate = m_linkhub.getTime(m_useStaticIP, m_useLocalTimeYN, false)
		Call winhttp1.setRequestHeader("x-bc-date", xDate)
	

		Dim target : target = "POST" + Chr(10)
		If postdata <> "" Then
			target = target + m_Linkhub.b64_sha256(postData) + Chr(10)
		End If	
		target = target + xDate + Chr(10)
		target = target + url + Chr(10)
		
		Dim auth_target : auth_target =  m_Linkhub.b64_hmac_sha256(m_Linkhub.SecretKey, target)

		Call winhttp1.setRequestHeader("x-bc-auth", auth_target)
		Call winhttp1.setRequestHeader("x-bc-encryptionmode", "CBC")

		winhttp1.Send (postdata)
		winhttp1.WaitForResponse
		Dim result : result = winhttp1.responseText
		
		If winhttp1.Status <> 200 Then
			Set winhttp1 = Nothing
			Dim parsedDic :  Set parsedDic = m_Linkhub.parse(result)
			Err.raise parsedDic.code, "BAROCERT", parsedDic.message
		End If
		
		Set winhttp1 = Nothing
		Set httpPOST = m_Linkhub.parse(result)
	End Function

	Private Function StringToBytes(Str)
		Dim Stream : Set Stream = Server.CreateObject("ADODB.Stream")
		Stream.Type = adTypeText
		Stream.Charset = "UTF-8"
		Stream.Open
		Stream.WriteText Str
		Stream.Flush
		Stream.Position = 0
		Stream.Type = adTypeBinary
		buffer= Stream.Read
		Stream.Close
		'Remove BOM.
		Set Stream = Server.CreateObject("ADODB.Stream")
		Stream.Type = adTypeBinary
		Stream.Open
		Stream.write buffer
		Stream.Flush
		Stream.Position = 3
		StringToBytes= Stream.Read
		Stream.Close
		Set Stream = Nothing
	
	End Function

	Private Function IIf(condition , trueState, falseState)
		If condition Then 
			IIf = trueState
		Else
			IIf = falseState
		End if
	End Function

	public Function toString(object)
		toString = m_Linkhub.toString(object)
	End Function

	Public Function parse(jsonString)
		Set parse = m_Linkhub.parse(jsonString)
	End Function

	Public Function encrypt(plainText)
		encrypt = m_Encryptor.enc(plainText)
	End Function

	public Function sha256_base64url(target)
		sha256_base64url = m_Linkhub.sha256ToBase64url(target)
	End Function
	
End Class

%>