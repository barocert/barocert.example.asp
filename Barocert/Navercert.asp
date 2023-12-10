<%

Class NavercertService

	Private m_BarocertBase

	Public Sub Class_Initialize
		Set m_BarocertBase = New BarocertBase
		m_BarocertBase.AddScope("421")
		m_BarocertBase.AddScope("422")
		m_BarocertBase.AddScope("423")
		m_BarocertBase.AddScope("424")
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

		Dim result : Set result = m_BarocertBase.httpPOST("/NAVER/Identity/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New IdentityReceipt
		infoTmp.fromJsonInfo result

		Set RequestIdentity = infoTmp
	End Function

	Public Function GetIdentityStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "NAVERCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "NAVERCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpGET("/NAVER/Identity/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token)

		Dim infoTmp : Set infoTmp = New IdentityStatus
		infoTmp.fromJsonInfo result

		Set GetIdentityStatus = infoTmp
	End Function 

	Public Function VerifyIdentity(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "NAVERCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "NAVERCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpPOST("/NAVER/Identity/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), "")
		
		Dim infoTmp : Set infoTmp = New IdentityResult
		infoTmp.fromJsonInfo result

		Set VerifyIdentity = infoTmp
	End Function 

	Public Function RequestSign(ClientCode, ByRef sign)

		Dim tmpDic : Set tmpDic = sign.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/NAVER/Sign/" + ClientCode , m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New SignReceipt
		infoTmp.fromJsonInfo result

		Set RequestSign = infoTmp
	End Function 

	Public Function GetSignStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "NAVERCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "NAVERCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpGET("/NAVER/Sign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token())

		Dim infoTmp : Set infoTmp = New SignStatus
		infoTmp.fromJsonInfo result
		Set GetSignStatus = infoTmp
	End Function 

	Public Function VerifySign(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "NAVERCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "NAVERCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpPOST("/NAVER/Sign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), "")

		Dim infoTmp : Set infoTmp = New SignResult
		infoTmp.fromJsonInfo result
		Set VerifySign = infoTmp
	End Function 

	Public Function RequestMultiSign(ClientCode, ByRef multiSign)

		Dim tmpDic : Set tmpDic = multiSign.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/NAVER/MultiSign/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New MultiSignReceipt
		infoTmp.fromJsonInfo result

		Set RequestMultiSign = infoTmp
	End Function 

	Public Function GetMultiSignStatus(ClientCode, ReceiptID)
		If ClientCode = "" Then
			Err.Raise -99999999, "NAVERCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "NAVERCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New MultiSignStatus
		Dim result : Set result = m_BarocertBase.httpGET("/NAVER/MultiSign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token())

		infoTmp.fromJsonInfo result
		Set GetMultiSignStatus = infoTmp
	End Function 

	Public Function VerifyMultiSign(ClientCode, ReceiptID)
		If ClientCode = "" Then
			Err.Raise -99999999, "NAVERCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "NAVERCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New MultiSignResult
		Dim result : Set result = m_BarocertBase.httpPOST("/NAVER/MultiSign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), "")

		infoTmp.fromJsonInfo result
		Set VerifyMultiSign = infoTmp
	End Function 	

	Public Function RequestCMS(ClientCode, ByRef CMS)

		Dim tmpDic : Set tmpDic = CMS.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/NAVER/CMS/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New CMSReceipt
		infoTmp.fromJsonInfo result

		Set RequestCMS = infoTmp
	End Function

	Public Function GetCMSStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "NAVERCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "NAVERCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpGET("/NAVER/CMS/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token)

		Dim infoTmp : Set infoTmp = New CMSStatus
		infoTmp.fromJsonInfo result

		Set GetCMSStatus = infoTmp
	End Function 

	Public Function VerifyCMS(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "NAVERCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "NAVERCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpPOST("/NAVER/CMS/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), "")
		
		Dim infoTmp : Set infoTmp = New CMSResult
		infoTmp.fromJsonInfo result

		Set VerifyCMS = infoTmp
	End Function 

	Public Function encrypt(PlainText)
		encrypt = m_BarocertBase.encrypt(PlainText)
	End Function

	Public Function sha256(target)
		sha256 = m_BarocertBase.sha256(target)
	End Function

End Class

Class Identity
	Public receiverHP
	Public receiverName
	Public receiverBirthday
	Public callCenterNum
	Public expireIn
	Public deviceOSType
	Public returnURL
	Public appUseYN

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "callCenterNum", callCenterNum
		toJsonInfo.Set "expireIn", expireIn
		toJsonInfo.Set "deviceOSType", deviceOSType
		toJsonInfo.Set "returnURL", returnURL
		toJsonInfo.Set "appUseYN", appUseYN
	End Function 

End Class

Class IdentityReceipt
	Public receiptID
	Public scheme
	Public marketUrl

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
			If Not isEmpty(jsonInfo.marketUrl) Then
				marketUrl = jsonInfo.marketUrl
			End If
		On Error GoTo 0
	End Sub
End Class

Class IdentityStatus
	Public clientCode	
	Public receiptID
	Public state
	Public expireIn
	Public callCenterName
	Public callCenterNum
	Public returnURL
	Public expireDT
	Public deviceOSType
	Public scheme
	Public appUseYN

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.clientCode) Then
				clientCode = jsonInfo.clientCode
			End If
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.expireIn) Then
				expireIn = jsonInfo.expireIn
			End If
			If Not isEmpty(jsonInfo.callCenterName) Then
				callCenterName = jsonInfo.callCenterName
			End If
			If Not isEmpty(jsonInfo.callCenterNum) Then
				callCenterNum = jsonInfo.callCenterNum
			End If
			If Not isEmpty(jsonInfo.returnURL) Then
				returnURL = jsonInfo.returnURL
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.deviceOSType) Then
				deviceOSType = jsonInfo.deviceOSType
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
			If Not isEmpty(jsonInfo.appUseYN) Then
				appUseYN = jsonInfo.appUseYN
			End If
		On Error GoTo 0
	End Sub
End Class

Class IdentityResult
	Public receiptID
	Public state
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverHP
	Public receiverGender
	Public receiverEmail
	Public receiverForeign
	Public signedData
	Public ci
	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.receiverName) Then
				receiverName = jsonInfo.receiverName
			End If
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
			End If
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
			If Not isEmpty(jsonInfo.receiverEmail) Then
				receiverEmail = jsonInfo.receiverEmail
			End If
			If Not isEmpty(jsonInfo.receiverForeign) Then
				receiverForeign = jsonInfo.receiverForeign
			End If
			If Not isEmpty(jsonInfo.signedData) Then
				signedData = jsonInfo.signedData
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
			End If
		On Error GoTo 0
	End Sub
End Class

Class Sign
	Public receiverHP
	Public receiverName
	Public receiverBirthday
	Public reqTitle
	Public reqMessage
	Public callCenterNum
	Public expireIn
	Public token
	Public tokenType
	Public returnURL
	Public deviceOSType
	Public appUseYN

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "reqMessage", reqMessage
		toJsonInfo.Set "callCenterNum", callCenterNum
		toJsonInfo.Set "expireIn", expireIn
		toJsonInfo.Set "token", token
		toJsonInfo.Set "tokenType", tokenType
		toJsonInfo.Set "returnURL", returnURL
		toJsonInfo.Set "deviceOSType", deviceOSType
		toJsonInfo.Set "appUseYN", appUseYN
	End Function 

End Class

Class SignReceipt
	Public receiptID
	Public scheme
	Public marketUrl

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
			If Not isEmpty(jsonInfo.marketUrl) Then
				marketUrl = jsonInfo.marketUrl
			End If
		On Error GoTo 0
	End Sub
End Class

Class SignStatus
	Public clientCode	
	Public receiptID
	Public state
	Public expireIn
	Public callCenterName
	Public callCenterNum
	Public reqTitle
	Public returnURL
	Public expireDT
	Public tokenType
	Public scheme
	Public deviceOSType
	Public appUseYN

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.clientCode) Then
				clientCode = jsonInfo.clientCode
			End If
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.expireIn) Then
				expireIn = jsonInfo.expireIn
			End If
			If Not isEmpty(jsonInfo.callCenterName) Then
				callCenterName = jsonInfo.callCenterName
			End If
			If Not isEmpty(jsonInfo.callCenterNum) Then
				callCenterNum = jsonInfo.callCenterNum
			End If
			If Not isEmpty(jsonInfo.reqTitle) Then
				reqTitle = jsonInfo.reqTitle
			End If
			If Not isEmpty(jsonInfo.returnURL) Then
				returnURL = jsonInfo.returnURL
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.tokenType) Then
				tokenType = jsonInfo.tokenType
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
			If Not isEmpty(jsonInfo.deviceOSType) Then
				deviceOSType = jsonInfo.deviceOSType
			End If
			If Not isEmpty(jsonInfo.appUseYN) Then
				appUseYN = jsonInfo.appUseYN
			End If
		On Error GoTo 0
	End Sub
End Class

Class SignResult
	Public receiptID
	Public state
	Public receiverHP
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverGender
	Public receiverEmail
	Public receiverForeign
	Public signedData
	Public ci

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
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
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
			If Not isEmpty(jsonInfo.receiverEmail) Then
				receiverEmail = jsonInfo.receiverEmail
			End If
			If Not isEmpty(jsonInfo.receiverForeign) Then
				receiverForeign = jsonInfo.receiverForeign
			End If
			If Not isEmpty(jsonInfo.signedData) Then
				signedData = jsonInfo.signedData
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
			End If
		On Error GoTo 0
	End Sub
End Class

Class MultiSign
	Public receiverHP
	Public receiverName
	Public receiverBirthday
	Public callCenterNum
	Public reqTitle
	Public reqMessage
	Public expireIn
	Public tokens
	Public returnURL
	Public deviceOSType
	Public appUseYN

	Public Sub Class_Initialize
		Set tokens = CreateObject("Scripting.Dictionary")
	End Sub

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "callCenterNum", callCenterNum
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "reqMessage", reqMessage
		toJsonInfo.Set "expireIn", expireIn
		Dim multiSignTokens : Set multiSignTokens = JSON.parse("[]")
		Dim i
		For i=0 To tokens.Count-1
			Dim signToken : Set signToken = New MultiSignTokens
			signToken.setValue tokens.Item(i)
			multiSignTokens.Set i, signToken.toJsonInfo
		Next
		toJsonInfo.Set "tokens", multiSignTokens
		toJsonInfo.Set "returnURL", returnURL
		toJsonInfo.Set "deviceOSType", tokenType
		toJsonInfo.Set "appUseYN", appUseYN
	End Function 

	Public Sub addToken(index, data)
		tokens.Add index, data
	End Sub
End Class

Class MultiSignReceipt
	Public receiptID
	Public scheme
	Public marketUrl

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
			If Not isEmpty(jsonInfo.marketUrl) Then
				marketUrl = jsonInfo.marketUrl
			End If
		On Error GoTo 0
	End Sub
End Class

Class MultiSignTokens
	Public tokenType
	Public token

	Public Sub setValue(multiSignToken)

		If Not isEmpty(multiSignToken.tokenType) Then
			tokenType = multiSignToken.tokenType
		End If

		If Not isEmpty(multiSignToken.token) Then
			token = multiSignToken.token
		End If
	End Sub

	Public Sub fromJsonInfo(jsonInfo)
		Set toJsonInfo = JSON.parse("{}")
			On Error Resume Next
				If Not isEmpty(jsonInfo.tokenType) Then
				tokenType = jsonInfo.tokenType
			End If
			If Not isEmpty(jsonInfo.token) Then
				token = jsonInfo.token
			End If
		On Error GoTo 0
	End Sub

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "tokenType", tokenType
		toJsonInfo.Set "token", token
	End Function 
End Class

Class MultiSignStatus
	Public receiptID
	Public clientCode
	Public state
	Public expireIn
	Public callCenterName
	Public callCenterNum
	Public reqTitle
	Public returnURL
	Public tokenTypes()
	Public expireDT
	Public scheme
	Public deviceOSType
	Public appUseYN

	Public Sub Class_Initialize
		ReDim tokenTypes(-1)
	End Sub

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
			If Not isEmpty(jsonInfo.expireIn) Then
				expireIn = jsonInfo.expireIn
			End If
			If Not isEmpty(jsonInfo.callCenterName) Then
				callCenterName = jsonInfo.callCenterName
			End If
			If Not isEmpty(jsonInfo.callCenterNum) Then
				callCenterNum = jsonInfo.callCenterNum
			End If
			If Not isEmpty(jsonInfo.reqTitle) Then
				reqTitle = jsonInfo.reqTitle
			End If
			If Not isEmpty(jsonInfo.returnURL) Then
				returnURL = jsonInfo.returnURL
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
			If Not isEmpty(jsonInfo.deviceOSType) Then
				deviceOSType = jsonInfo.deviceOSType
			End If
			If Not isEmpty(jsonInfo.appUseYN) Then
				appUseYN = jsonInfo.appUseYN
			End If

			ReDim tokenTypes(jsonInfo.tokenTypes.length)
			Dim i
			For i = 0 To jsonInfo.tokenTypes.length -1
				Dim tmpObj : Set tmpObj = New MultiSignedData
				tokenTypes(i) = jsonInfo.tokenTypes.Get(i)
				Set tokenTypes(i) = tmpObj
			Next

		On Error GoTo 0
	End Sub
End Class

Class MultiSignResult
	Public receiptID
	Public state
	Public receiverHP
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverGender
	Public receiverEmail
	Public receiverForeign
	Public multiSignedData()
	Public ci

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
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
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
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
			If Not isEmpty(jsonInfo.receiverEmail) Then
				receiverEmail = jsonInfo.receiverEmail
			End If
			If Not isEmpty(jsonInfo.receiverForeign) Then
				receiverForeign = jsonInfo.receiverForeign
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
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
	Public receiverHP
	Public receiverName
	Public receiverBirthday
	Public reqTitle
	Public reqMessage
	Public callCenterNum
	Public expireIn
	Public requestCorp
	Public bankName
	Public bankAccountNum
	Public bankAccountName
	Public bankAccountBirthday
	Public deviceOSType
	Public returnURL
	Public appUseYN

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "reqMessage", reqMessage
		toJsonInfo.Set "callCenterNum", callCenterNum
		toJsonInfo.Set "expireIn", expireIn
		toJsonInfo.Set "requestCorp", requestCorp
		toJsonInfo.Set "bankName", bankName
		toJsonInfo.Set "bankAccountNum", bankAccountNum
		toJsonInfo.Set "bankAccountName", bankAccountName
		toJsonInfo.Set "bankAccountBirthday", bankAccountBirthday
		toJsonInfo.Set "deviceOSType", deviceOSType
		toJsonInfo.Set "returnURL", returnURL
		toJsonInfo.Set "appUseYN", appUseYN
	End Function 

End Class

Class CMSReceipt
	Public receiptID
	Public scheme
	Public marketUrl

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.scheme) Then
				scheme = jsonInfo.scheme
			End If
			If Not isEmpty(jsonInfo.marketUrl) Then
				marketUrl = jsonInfo.marketUrl
			End If
		On Error GoTo 0
	End Sub
End Class

Class CMSStatus
	Public clientCode	
	Public receiptID
	Public state
	Public expireDT

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.clientCode) Then
				clientCode = jsonInfo.clientCode
			End If
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
		On Error GoTo 0
	End Sub
End Class

Class CMSResult
	Public receiptID
	Public state
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverHP
	Public receiverGender
	Public receiverEmail
	Public receiverForeign
	Public signedData
	Public ci
	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptID) Then
				receiptID = jsonInfo.receiptID
			End If
			If Not isEmpty(jsonInfo.state) Then
				state = jsonInfo.state
			End If
			If Not isEmpty(jsonInfo.receiverName) Then
				receiverName = jsonInfo.receiverName
			End If
			If Not isEmpty(jsonInfo.receiverHP) Then
				receiverHP = jsonInfo.receiverHP
			End If
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
			If Not isEmpty(jsonInfo.receiverEmail) Then
				receiverEmail = jsonInfo.receiverEmail
			End If
			If Not isEmpty(jsonInfo.receiverForeign) Then
				receiverForeign = jsonInfo.receiverForeign
			End If
			If Not isEmpty(jsonInfo.signedData) Then
				signedData = jsonInfo.signedData
			End If
			If Not isEmpty(jsonInfo.ci) Then
				ci = jsonInfo.ci
			End If
		On Error GoTo 0
	End Sub
End Class

%>