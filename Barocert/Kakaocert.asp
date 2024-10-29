<%

Class KakaocertService

	Private m_BarocertBase

	Public Sub Class_Initialize
		Set m_BarocertBase = New BarocertBase
		m_BarocertBase.AddScope("401")
		m_BarocertBase.AddScope("402")
		m_BarocertBase.AddScope("403")
		m_BarocertBase.AddScope("404")
		m_BarocertBase.AddScope("405")
	End Sub

    Public Sub Initialize(linkID, SecretKey)
        m_BarocertBase.Initialize linkID,SecretKey
    End Sub

	Public Property Let IPRestrictOnOff(ByVal value)
		m_BarocertBase.IPRestrictOnOff = value
	End Property

	Public Property Let UseStaticIP(ByVal value)
		m_BarocertBase.UseStaticIP = value
	End Property

	Public Property Let UseGAIP(ByVal value)
		m_BarocertBase.UseGAIP = value
	End Property

	Public Property Let UseLocalTimeYN(ByVal value)
		m_BarocertBase.UseLocalTimeYN = value
	End Property

	Public Property Let ServiceURL(ByVal value)
		m_BarocertBase.ServiceURL = value
	End Property
	
	Public Property Let AuthURL(ByVal value)
		m_BarocertBase.AuthURL = value
	End Property

	public Function toString(object)
		toString = m_BarocertBase.toString(object)
	End Function

	Public Function RequestIdentity(ClientCode, ByRef Identity)

		Dim tmpDic : Set tmpDic = Identity.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/Identity/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New IdentityReceipt
		infoTmp.fromJsonInfo result

		Set RequestIdentity = infoTmp
	End Function

	Public Function GetIdentityStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpGET("/KAKAO/Identity/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token)

		Dim infoTmp : Set infoTmp = New IdentityStatus
		infoTmp.fromJsonInfo result

		Set GetIdentityStatus = infoTmp
	End Function 

	Public Function VerifyIdentity(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/Identity/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), "")
		
		Dim infoTmp : Set infoTmp = New IdentityResult
		infoTmp.fromJsonInfo result

		Set VerifyIdentity = infoTmp
	End Function 

	Public Function RequestSign(ClientCode, ByRef sign)

		Dim tmpDic : Set tmpDic = sign.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/Sign/" + ClientCode , m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New SignReceipt
		infoTmp.fromJsonInfo result

		Set RequestSign = infoTmp
	End Function 

	Public Function GetSignStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpGET("/KAKAO/Sign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token())

		Dim infoTmp : Set infoTmp = New SignStatus
		infoTmp.fromJsonInfo result
		Set GetSignStatus = infoTmp
	End Function 

	Public Function VerifySign(ClientCode, ReceiptID )

		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/Sign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), "")

		Dim infoTmp : Set infoTmp = New SignResult
		infoTmp.fromJsonInfo result
		Set VerifySign = infoTmp
	End Function 

	Public Function RequestMultiSign(ClientCode, ByRef multiSign )

		Dim tmpDic : Set tmpDic = multiSign.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/MultiSign/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New MultiSignReceipt
		infoTmp.fromJsonInfo result

		Set RequestMultiSign = infoTmp
	End Function 

	Public Function GetMultiSignStatus(ClientCode, ReceiptID)
		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New MultiSignStatus
		Dim result : Set result = m_BarocertBase.httpGET("/KAKAO/MultiSign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token())

		infoTmp.fromJsonInfo result
		Set GetMultiSignStatus = infoTmp
	End Function 

	Public Function VerifyMultiSign(ClientCode, ReceiptID)
		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New MultiSignResult
		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/MultiSign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), "")

		infoTmp.fromJsonInfo result
		Set VerifyMultiSign = infoTmp
	End Function 

	Public Function RequestCMS(ClientCode, ByRef cms)

		Dim tmpDic : Set tmpDic = cms.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/CMS/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New CMSReceipt
		infoTmp.fromJsonInfo result

		Set RequestCMS = infoTmp
	End Function

	Public Function GetCMSStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New CMSStatus
		Dim result : Set result = m_BarocertBase.httpGET("/KAKAO/CMS/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token())

		infoTmp.fromJsonInfo result
		Set GetCMSStatus = infoTmp
	End Function 

	Public Function VerifyCMS(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New CMSResult
		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/CMS/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), "")

		infoTmp.fromJsonInfo result
		Set VerifyCMS = infoTmp
	End Function 

	Public Function VerifyLogin(ClientCode, txID)

		If ClientCode = "" Then
			Err.Raise -99999999, "KAKAOCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If txID = "" Then
			Err.Raise -99999999, "KAKAOCERT", "트랜잭션 아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New LoginResult
		Dim result : Set result = m_BarocertBase.httpPOST("/KAKAO/Login/" + ClientCode + "/" + txID, m_BarocertBase.getSession_token(), "")

		infoTmp.fromJsonInfo result
		Set VerifyLogin = infoTmp
	End Function 

	Public Function encrypt(PlainText)
		encrypt = m_BarocertBase.encrypt(PlainText)
	End Function
End Class

Class Identity
	Public receiverHP
	Public receiverName
	Public receiverBirthday
	Public reqTitle
	Public extraMessage
	Public expireIn
	Public token
	Public returnURL
	Public appUseYN

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "extraMessage", extraMessage
		toJsonInfo.Set "expireIn", expireIn
		toJsonInfo.Set "token", token
		toJsonInfo.Set "returnURL", returnURL
		toJsonInfo.Set "appUseYN", appUseYN
	End Function 

End Class

Class IdentityReceipt
	Public receiptID
	Public scheme

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
		On Error GoTo 0
	End Sub
End Class

Class IdentityStatus
	Public receiptID
	Public clientCode
	Public state
	Public requestDT
	Public viewDT
	Public completeDT
	Public expireDT
	Public verifyDT

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID	
			End If
			If Not isEmpty(jsonInfo.clientCode) Then
				clientCode = jsonInfo.clientCode
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.requestDT) Then
				requestDT = jsonInfo.requestDT
			End If
			If Not isEmpty(jsonInfo.viewDT) Then
				viewDT = jsonInfo.viewDT
			End If
			If Not isEmpty(jsonInfo.completeDT) Then
				completeDT = jsonInfo.completeDT
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.verifyDT) Then
				verifyDT = jsonInfo.verifyDT
			End If
		On Error GoTo 0
	End Sub
End Class

Class IdentityResult
	Public receiptID
	Public state
	Public signedData
	Public ci
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverHP
	Public receiverGender

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.signedData) Then
				signedData = jsonInfo.signedData
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
			End If
			If Not isEmpty(jsonInfo.receiverName) Then
				receiverName = jsonInfo.receiverName
			End If
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
		On Error GoTo 0
	End Sub
End Class

Class Sign
	Public receiverHP
	Public receiverName
	Public receiverBirthday
	Public reqTitle
	Public signTitle
	Public extraMessage
	Public expireIn
	Public token
    Public tokenType
	Public returnURL
	Public appUseYN

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "signTitle", signTitle
		toJsonInfo.Set "extraMessage", extraMessage
		toJsonInfo.Set "expireIn", expireIn
		toJsonInfo.Set "token", token
		toJsonInfo.Set "tokenType", tokenType
		toJsonInfo.Set "returnURL", returnURL
		toJsonInfo.Set "appUseYN", appUseYN
	End Function 

End Class

Class SignReceipt
	Public receiptID
	Public scheme

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
		On Error GoTo 0
	End Sub
End Class

Class SignStatus
	Public receiptID
	Public clientCode
	Public state
	Public requestDT
	Public viewDT
	Public completeDT
	Public expireDT
	Public verifyDT

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.clientCode) Then
				clientCode = jsonInfo.clientCode
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.requestDT) Then
				requestDT = jsonInfo.requestDT
			End If
			If Not isEmpty(jsonInfo.viewDT) Then
				viewDT = jsonInfo.viewDT
			End If
			If Not isEmpty(jsonInfo.completeDT) Then
				completeDT = jsonInfo.completeDT
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.verifyDT) Then
				verifyDT = jsonInfo.verifyDT
			End If
		On Error GoTo 0
	End Sub
End Class

Class SignResult
	Public receiptID
	Public state
	Public signedData
	Public ci
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverHP
	Public receiverGender

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.signedData) Then
				signedData = jsonInfo.signedData
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
			End If
			If Not isEmpty(jsonInfo.receiverName) Then
				receiverName = jsonInfo.receiverName
			End If
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
		On Error GoTo 0
	End Sub
End Class

Class MultiSignTokens
	Public reqTitle
	Public signTitle
	Public token

	Public Sub setValue(multiSignToken)

		If Not isEmpty(multiSignToken.reqTitle) Then
			reqTitle = multiSignToken.reqTitle
		End If
		
		If Not isEmpty(multiSignToken.signTitle) Then
			signTitle = multiSignToken.signTitle
		End If

		If Not isEmpty(multiSignToken.token) Then
			token = multiSignToken.token
		End If
	End Sub

	Public Sub fromJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		On Error Resume Next
			If Not isEmpty(reqTitle) Then
				toJsonInfo.set "reqTitle", reqTitle
			End If
			If Not isEmpty(signTitle) Then
				toJsonInfo.set "signTitle", signTitle
			End If
			If Not isEmpty(token) Then
				toJsonInfo.set "token", token
			End If
		On Error GoTo 0
	End Sub

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "signTitle", signTitle
		toJsonInfo.Set "token", token
	End Function 
End Class

Class MultiSign
	Public receiverHP
	Public receiverName
	Public receiverBirthday
	Public reqTitle
	Public extraMessage
	Public expireIn
	Public tokens
	Public tokenType
	Public returnURL
	Public appUseYN

	Public Sub Class_Initialize
		Set tokens = CreateObject("Scripting.Dictionary")
	End Sub

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "extraMessage", extraMessage
		toJsonInfo.Set "expireIn", expireIn
		Dim multiSignTokens : Set multiSignTokens = JSON.parse("[]")
		Dim i
		For i=0 To tokens.Count-1
			Dim signToken : Set signToken = New MultiSignTokens
			signToken.setValue tokens.Item(i)
			multiSignTokens.Set i, signToken.toJsonInfo
		Next
		toJsonInfo.Set "tokens", multiSignTokens
		toJsonInfo.Set "tokenType", tokenType
		toJsonInfo.Set "returnURL", returnURL
		toJsonInfo.Set "appUseYN", appUseYN
	End Function 

	Public Sub addToken(index, data)
		tokens.Add index, data
	End Sub
End Class

Class MultiSignReceipt
	Public receiptID
	Public scheme

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
		On Error GoTo 0
	End Sub
End Class

Class MultiSignedData
	Public reqTitle
	Public signTitle
	Public token

	Public Sub setValue(multiSignToken)

		If Not isEmpty(multiSignToken.reqTitle) Then
			reqTitle = multiSignToken.reqTitle
		End If
		
		If Not isEmpty(multiSignToken.signTitle) Then
			signTitle = multiSignToken.signTitle
		End If

		If Not isEmpty(multiSignToken.token) Then
			token = multiSignToken.token
		End If
	End Sub

	Public Sub fromJsonInfo(jsonInfo)
		Set toJsonInfo = JSON.parse("{}")
		On Error Resume Next
		If Not isEmpty(jsonInfo.reqTitle) Then
			reqTitle = jsonInfo.reqTitle
		End If
		If Not isEmpty(jsonInfo.signTitle) Then
			signTitle = jsonInfo.signTitle
		End If
		If Not isEmpty(jsonInfo.token) Then
			token = jsonInfo.token
		End If
		On Error GoTo 0
	End Sub
End Class

Class MultiSignStatus
	Public receiptID
	Public clientCode
	Public state
	Public requestDT
	Public viewDT
	Public completeDT
	Public expireDT
	Public verifyDT

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.clientCode) Then
				clientCode = jsonInfo.clientCode
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.requestDT) Then
				requestDT = jsonInfo.requestDT
			End If
			If Not isEmpty(jsonInfo.viewDT) Then
				viewDT = jsonInfo.viewDT
			End If
			If Not isEmpty(jsonInfo.completeDT) Then
				completeDT = jsonInfo.completeDT
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.verifyDT) Then
				verifyDT = jsonInfo.verifyDT
			End If
		On Error GoTo 0
	End Sub
End Class

Class MultiSignResult
	Public receiptID
	Public state
	Public ci
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverHP
	Public receiverGender
	Public multiSignedData()

	Public Sub Class_Initialize
		ReDim multiSignedData(-1)
	End Sub

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
			End If
			If Not isEmpty(jsonInfo.receiverName) Then
				receiverName = jsonInfo.receiverName
			End If
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If

			ReDim multiSignedData(jsonInfo.multiSignedData.length)
			Dim i
			For i = 0 To jsonInfo.multiSignedData.length -1
				Dim tmpObj : Set tmpObj = New MultiSignedData
				multiSignedData(i) = jsonInfo.multiSignedData.Get(i)
				Set multiSignedData(i) = tmpObj
			Next

		On Error GoTo 0
	End Sub
End Class

Class CMS
	public receiverHP
	public receiverName
	public receiverBirthday
	public reqTitle
	public extraMessage
	public expireIn
	public returnURL
	public requestCorp
	public bankName
	public bankAccountNum
	public bankAccountName
	public bankAccountBirthday
	public bankServiceType
	public appUseYN

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "extraMessage", extraMessage
		toJsonInfo.Set "expireIn", expireIn
		toJsonInfo.Set "returnURL", returnURL
		toJsonInfo.Set "requestCorp", requestCorp 
		toJsonInfo.Set "bankName", bankName 
		toJsonInfo.Set "bankAccountNum", bankAccountNum 
		toJsonInfo.Set "bankAccountName", bankAccountName 
		toJsonInfo.Set "bankAccountBirthday", bankAccountBirthday 
		toJsonInfo.Set "bankServiceType", bankServiceType 
		toJsonInfo.Set "appUseYN", appUseYN
	End Function 

End Class

Class CMSReceipt
	Public receiptID
	Public scheme

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
		On Error GoTo 0
	End Sub
End Class

Class CMSStatus
	Public receiptID
	Public clientCode
	Public state
	Public requestDT
	Public viewDT
	Public completeDT
	Public expireDT
	Public verifyDT

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.clientCode) Then
				clientCode = jsonInfo.clientCode
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.requestDT) Then
				requestDT = jsonInfo.requestDT
			End If
			If Not isEmpty(jsonInfo.viewDT) Then
				viewDT = jsonInfo.viewDT
			End If
			If Not isEmpty(jsonInfo.completeDT) Then
				completeDT = jsonInfo.completeDT
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.verifyDT) Then
				verifyDT = jsonInfo.verifyDT
			End If
		On Error GoTo 0
	End Sub
End Class

Class CMSResult
	Public receiptID
	Public state
	Public signedData
	Public ci
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverHP
	Public receiverGender

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.signedData) Then
				signedData = jsonInfo.signedData
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
			End If
			If Not isEmpty(jsonInfo.receiverName) Then
				receiverName = jsonInfo.receiverName
			End If
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
		On Error GoTo 0
	End Sub
End Class

Class LoginResult
	Public txID
	Public state
	Public signedData
	Public ci
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverHP
	Public receiverGender

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.txID) Then
				txID = jsonInfo.txID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.signedData) Then
				signedData = jsonInfo.signedData
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
			End If
			If Not isEmpty(jsonInfo.receiverName) Then
				receiverName = jsonInfo.receiverName
			End If
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
		On Error GoTo 0
	End Sub
End Class

%>