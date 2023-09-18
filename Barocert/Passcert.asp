<%

Class PasscertService

	Private m_BarocertBase

	Public Sub Class_Initialize
		Set m_BarocertBase = New BarocertBase
		m_BarocertBase.AddScope("441")
		m_BarocertBase.AddScope("442")
		m_BarocertBase.AddScope("443")
		m_BarocertBase.AddScope("444")
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

	public Function toString(object)
		toString = m_BarocertBase.toString(object)
	End Function

	Public Function RequestIdentity(ClientCode, ByRef Identity)

		Dim tmpDic : Set tmpDic = Identity.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/PASS/Identity/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New IdentityReceipt
		infoTmp.fromJsonInfo result

		Set RequestIdentity = infoTmp
	End Function

	Public Function GetIdentityStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "PASSCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "PASSCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpGET("/PASS/Identity/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token)

		Dim infoTmp : Set infoTmp = New IdentityStatus
		infoTmp.fromJsonInfo result

		Set GetIdentityStatus = infoTmp
	End Function 

	Public Function VerifyIdentity(ClientCode, ReceiptID, ByRef IdentityVerify)

		Dim tmpDic : Set tmpDic = IdentityVerify.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		If ClientCode = "" Then
			Err.Raise -99999999, "PASSCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "PASSCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpPOST("/PASS/Identity/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), postdata)
		
		Dim infoTmp : Set infoTmp = New IdentityResult
		infoTmp.fromJsonInfo result

		Set VerifyIdentity = infoTmp
	End Function 

	Public Function RequestSign(ClientCode, ByRef sign)

		Dim tmpDic : Set tmpDic = sign.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/PASS/Sign/" + ClientCode , m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New SignReceipt
		infoTmp.fromJsonInfo result

		Set RequestSign = infoTmp
	End Function 

	Public Function GetSignStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "PASSCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "PASSCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpGET("/PASS/Sign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token())

		Dim infoTmp : Set infoTmp = New SignStatus
		infoTmp.fromJsonInfo result
		Set GetSignStatus = infoTmp
	End Function 

	Public Function VerifySign(ClientCode, ReceiptID, ByRef SignVerify)

		Dim tmpDic : Set tmpDic = SignVerify.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		If ClientCode = "" Then
			Err.Raise -99999999, "PASSCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "PASSCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim result : Set result = m_BarocertBase.httpPOST("/PASS/Sign/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New SignResult
		infoTmp.fromJsonInfo result
		Set VerifySign = infoTmp
	End Function 

	Public Function RequestCMS(ClientCode, ByRef cms)

		Dim tmpDic : Set tmpDic = cms.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/PASS/CMS/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New CMSReceipt
		infoTmp.fromJsonInfo result

		Set RequestCMS = infoTmp
	End Function

	Public Function GetCMSStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "PASSCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "PASSCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New CMSStatus
		Dim result : Set result = m_BarocertBase.httpGET("/PASS/CMS/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token())

		infoTmp.fromJsonInfo result
		Set GetCMSStatus = infoTmp
	End Function 

	Public Function VerifyCMS(ClientCode, ReceiptID, ByRef CMSVerify)

		Dim tmpDic : Set tmpDic = CMSVerify.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		If ClientCode = "" Then
			Err.Raise -99999999, "PASSCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "PASSCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New CMSResult
		Dim result : Set result = m_BarocertBase.httpPOST("/PASS/CMS/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), postdata)

		infoTmp.fromJsonInfo result
		Set VerifyCMS = infoTmp
	End Function 

	Public Function RequestLogin(ClientCode, ByRef login)

		Dim tmpDic : Set tmpDic = login.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		Dim result : Set result = m_BarocertBase.httpPOST("/PASS/Login/" + ClientCode, m_BarocertBase.getSession_token(), postdata)

		Dim infoTmp : Set infoTmp = New LoginReceipt
		infoTmp.fromJsonInfo result

		Set RequestLogin = infoTmp
	End Function

	Public Function GetLoginStatus(ClientCode, ReceiptID)

		If ClientCode = "" Then
			Err.Raise -99999999, "PASSCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "PASSCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New LoginStatus
		Dim result : Set result = m_BarocertBase.httpGET("/PASS/Login/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token())

		infoTmp.fromJsonInfo result
		Set GetLoginStatus = infoTmp
	End Function 

	Public Function VerifyLogin(ClientCode, ReceiptID, ByRef LoginVerify)

		Dim tmpDic : Set tmpDic = LoginVerify.toJsonInfo

		Dim postdata : postdata = toString(tmpDic)

		If ClientCode = "" Then
			Err.Raise -99999999, "PASSCERT", "이용기관코드가 입력되지 않았습니다."
		End If

		If ReceiptID = "" Then
			Err.Raise -99999999, "PASSCERT", "접수아이디가 입력되지 않았습니다."
		End If

		Dim infoTmp : Set infoTmp = New LoginResult
		Dim result : Set result = m_BarocertBase.httpPOST("/PASS/Login/" + ClientCode + "/" + ReceiptID, m_BarocertBase.getSession_token(), postdata)

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
	Public reqMessage
	Public callCenterNum
	Public expireIn
	Public token
	Public userAgreementYN
	Public receiverInfoYN
	Public telcoType
	Public deviceOSType
	Public appUseYN
	Public useTssYN

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
		toJsonInfo.Set "userAgreementYN", userAgreementYN
		toJsonInfo.Set "receiverInfoYN", receiverInfoYN
		toJsonInfo.Set "telcoType", telcoType
		toJsonInfo.Set "deviceOSType", deviceOSType
		toJsonInfo.Set "appUseYN", appUseYN
		toJsonInfo.Set "useTssYN", useTssYN
	End Function 

End Class

Class IdentityReceipt
	Public receiptId
	Public scheme
	Public marketUrl

	Public Sub fromJsonInfo(jsonInfo)
		On Error Resume Next
			If Not isEmpty(jsonInfo.receiptId) Then
				receiptId = jsonInfo.receiptId
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
	Public reqTitle
	Public reqMessage
	Public requestDT
	Public completeDT
	Public expireDT
	Public rejectDT
	Public tokenType
	Public userAgreementYN
	Public receiverInfoYN
	Public telcoType
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
			If Not isEmpty(jsonInfo.reqTitle) Then
				reqTitle = jsonInfo.reqTitle
			End If
			If Not isEmpty(jsonInfo.reqMessage) Then
				reqMessage = jsonInfo.reqMessage
			End If
			If Not isEmpty(jsonInfo.requestDT) Then
				requestDT = jsonInfo.requestDT
			End If
			If Not isEmpty(jsonInfo.completeDT) Then
				completeDT = jsonInfo.completeDT
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.rejectDT) Then
				rejectDT = jsonInfo.rejectDT
			End If
			If Not isEmpty(jsonInfo.tokenType) Then
				tokenType = jsonInfo.tokenType
			End If
			If Not isEmpty(jsonInfo.userAgreementYN) Then
				userAgreementYN = jsonInfo.userAgreementYN
			End If
			If Not isEmpty(jsonInfo.receiverInfoYN) Then
				receiverInfoYN = jsonInfo.receiverInfoYN
			End If
			If Not isEmpty(jsonInfo.telcoType) Then
				telcoType = jsonInfo.telcoType
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

Class IdentityVerify
	Public receiverHP
	Public receiverName

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
	End Function 
End Class

Class IdentityResult
	Public receiptID
	Public state
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverGender
	Public receiverForeign
	Public receiverTelcoType
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
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
			If Not isEmpty(jsonInfo.receiverForeign) Then
				receiverForeign = jsonInfo.receiverForeign
			End If
			If Not isEmpty(jsonInfo.receiverTelcoType) Then
				receiverTelcoType = jsonInfo.receiverTelcoType
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
	Public userAgreementYN
	Public receiverInfoYN
	Public originalTypeCode
	Public originalURL
	Public originalFormatCode
	Public telcoType
	Public deviceOSType
	Public appUseYN
	Public useTSSYN

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
		toJsonInfo.Set "userAgreementYN", userAgreementYN
		toJsonInfo.Set "receiverInfoYN", receiverInfoYN
		toJsonInfo.Set "originalTypeCode", originalTypeCode
		toJsonInfo.Set "originalURL", originalURL
		toJsonInfo.Set "originalFormatCode", originalFormatCode
		toJsonInfo.Set "telcoType", telcoType
		toJsonInfo.Set "deviceOSType", deviceOSType
		toJsonInfo.Set "appUseYN", appUseYN
		toJsonInfo.Set "useTssYN", useTssYN
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
	Public reqMessage
	Public requestDT
	Public completeDT
	Public expireDT
	Public rejectDT
	Public tokenType
	Public userAgreementYN
	Public receiverInfoYN
	Public telcoType
	Public deviceOSType
	Public originalTypeCode
	Public originalURL
	Public originalFormatCode
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
			If Not isEmpty(jsonInfo.reqTitle) Then
				reqTitle = jsonInfo.reqTitle
			End If
			If Not isEmpty(jsonInfo.reqMessage) Then
				reqMessage = jsonInfo.reqMessage
			End If
			If Not isEmpty(jsonInfo.requestDT) Then
				requestDT = jsonInfo.requestDT
			End If
			If Not isEmpty(jsonInfo.completeDT) Then
				completeDT = jsonInfo.completeDT
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.rejectDT) Then
				rejectDT = jsonInfo.rejectDT
			End If
			If Not isEmpty(jsonInfo.tokenType) Then
				tokenType = jsonInfo.tokenType
			End If
			If Not isEmpty(jsonInfo.userAgreementYN) Then
				userAgreementYN = jsonInfo.userAgreementYN
			End If
			If Not isEmpty(jsonInfo.receiverInfoYN) Then
				receiverInfoYN = jsonInfo.receiverInfoYN
			End If
			If Not isEmpty(jsonInfo.telcoType) Then
				telcoType = jsonInfo.telcoType
			End If
			If Not isEmpty(jsonInfo.deviceOSType) Then
				deviceOSType = jsonInfo.deviceOSType
			End If
			If Not isEmpty(jsonInfo.originalTypeCode) Then
				originalTypeCode = jsonInfo.originalTypeCode
			End If
			If Not isEmpty(jsonInfo.originalURL) Then
				originalURL = jsonInfo.originalURL
			End If
			If Not isEmpty(jsonInfo.originalFormatCode) Then
				originalFormatCode = jsonInfo.originalFormatCode
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

Class SignVerify
	Public receiverHP
	Public receiverName

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
	End Function 
End Class

Class SignResult
	Public receiptID
	Public state
	Public receiverHP
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverGender
	Public receiverForeign
	Public receiverTelcoType
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
			If Not isEmpty(jsonInfo.receiverForeign) Then
				receiverForeign = jsonInfo.receiverForeign
			End If
			If Not isEmpty(jsonInfo.receiverTelcoType) Then
				receiverTelcoType = jsonInfo.receiverTelcoType
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

Class CMS
	public receiverHP
	public receiverName
	public receiverBirthday
	public reqTitle
	Public reqMessage
	Public callCenterNum
	public expireIn
	Public userAgreementYN
	Public receiverInfoYN
	public bankName
	public bankAccountNum
	public bankAccountName
	public bankServiceType
	public bankWithdraw
	Public telcoType
	Public deviceOSType
	public appUseYN
	Public useTSSYN

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
		toJsonInfo.Set "receiverBirthday", receiverBirthday
		toJsonInfo.Set "reqTitle", reqTitle
		toJsonInfo.Set "reqMessage", reqMessage
		toJsonInfo.Set "callCenterNum", callCenterNum
		toJsonInfo.Set "expireIn", expireIn
		toJsonInfo.Set "userAgreementYN", userAgreementYN
		toJsonInfo.Set "receiverInfoYN", receiverInfoYN
		toJsonInfo.Set "bankName", bankName 
		toJsonInfo.Set "bankAccountNum", bankAccountNum 
		toJsonInfo.Set "bankAccountName", bankAccountName 
		toJsonInfo.Set "bankServiceType", bankServiceType 
		toJsonInfo.Set "bankWithdraw", bankWithdraw 
		toJsonInfo.Set "telcoType", telcoType
		toJsonInfo.Set "deviceOSType", deviceOSType
		toJsonInfo.Set "appUseYN", appUseYN
		toJsonInfo.Set "useTssYN", useTssYN
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
	Public expireIn
	Public callCenterName
	Public callCenterNum
	Public reqTitle
	Public reqMessage
	Public requestDT
	Public completeDT
	Public expireDT
	Public rejectDT
	Public tokenType
	Public userAgreementYN
	Public receiverInfoYN
	Public telcoType
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
			If Not isEmpty(jsonInfo.reqTitle) Then
				reqTitle = jsonInfo.reqTitle
			End If
			If Not isEmpty(jsonInfo.reqMessage) Then
				reqMessage = jsonInfo.reqMessage
			End If
			If Not isEmpty(jsonInfo.requestDT) Then
				requestDT = jsonInfo.requestDT
			End If
			If Not isEmpty(jsonInfo.completeDT) Then
				completeDT = jsonInfo.completeDT
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.rejectDT) Then
				rejectDT = jsonInfo.rejectDT
			End If
			If Not isEmpty(jsonInfo.tokenType) Then
				tokenType = jsonInfo.tokenType
			End If
			If Not isEmpty(jsonInfo.userAgreementYN) Then
				userAgreementYN = jsonInfo.userAgreementYN
			End If
			If Not isEmpty(jsonInfo.receiverInfoYN) Then
				receiverInfoYN = jsonInfo.receiverInfoYN
			End If
			If Not isEmpty(jsonInfo.telcoType) Then
				telcoType = jsonInfo.telcoType
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

Class CMSVerify
	Public receiverHP
	Public receiverName

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
	End Function 
End Class

Class CMSResult
	Public receiptID
	Public state
	Public receiverHP
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverGender
	Public receiverForeign
	Public receiverTelcoType
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
			If Not isEmpty(jsonInfo.receiverForeign) Then
				receiverForeign = jsonInfo.receiverForeign
			End If
			If Not isEmpty(jsonInfo.receiverTelcoType) Then
				receiverTelcoType = jsonInfo.receiverTelcoType
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

Class Login
	Public receiverHP
	Public receiverName
	Public receiverBirthday
	Public reqTitle
	Public reqMessage
	Public callCenterNum
	Public expireIn
	Public token
	Public userAgreementYN
	Public receiverInfoYN
	Public telcoType
	Public deviceOSType
	Public appUseYN
	Public useTSSYN

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
		toJsonInfo.Set "userAgreementYN", userAgreementYN
		toJsonInfo.Set "receiverInfoYN", receiverInfoYN
		toJsonInfo.Set "telcoType", telcoType
		toJsonInfo.Set "deviceOSType", deviceOSType
		toJsonInfo.Set "appUseYN", appUseYN
		toJsonInfo.Set "useTssYN", useTssYN
	End Function 

End Class

Class LoginReceipt
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

Class LoginStatus
Public clientCode	
	Public receiptID
	Public state
	Public expireIn
	Public callCenterName
	Public callCenterNum
	Public reqTitle
	Public reqMessage
	Public requestDT
	Public completeDT
	Public expireDT
	Public rejectDT
	Public tokenType
	Public userAgreementYN
	Public receiverInfoYN
	Public telcoType
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
			If Not isEmpty(jsonInfo.reqTitle) Then
				reqTitle = jsonInfo.reqTitle
			End If
			If Not isEmpty(jsonInfo.reqMessage) Then
				reqMessage = jsonInfo.reqMessage
			End If
			If Not isEmpty(jsonInfo.requestDT) Then
				requestDT = jsonInfo.requestDT
			End If
			If Not isEmpty(jsonInfo.completeDT) Then
				completeDT = jsonInfo.completeDT
			End If
			If Not isEmpty(jsonInfo.expireDT) Then
				expireDT = jsonInfo.expireDT
			End If
			If Not isEmpty(jsonInfo.rejectDT) Then
				rejectDT = jsonInfo.rejectDT
			End If
			If Not isEmpty(jsonInfo.tokenType) Then
				tokenType = jsonInfo.tokenType
			End If
			If Not isEmpty(jsonInfo.userAgreementYN) Then
				userAgreementYN = jsonInfo.userAgreementYN
			End If
			If Not isEmpty(jsonInfo.receiverInfoYN) Then
				receiverInfoYN = jsonInfo.receiverInfoYN
			End If
			If Not isEmpty(jsonInfo.telcoType) Then
				telcoType = jsonInfo.telcoType
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

Class LoginVerify
	Public receiverHP
	Public receiverName

	Public Function toJsonInfo()
		Set toJsonInfo = JSON.parse("{}")
		toJsonInfo.Set "receiverHP", receiverHP
		toJsonInfo.Set "receiverName", receiverName
	End Function 
End Class

Class LoginResult
	Public receiptID
	Public state
	Public receiverName
	Public receiverYear
	Public receiverDay
	Public receiverGender
	Public receiverForeign
	Public receiverTelcoType
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
			If Not isEmpty(jsonInfo.receiverYear) Then
				receiverYear = jsonInfo.receiverYear
			End If
			If Not isEmpty(jsonInfo.receiverDay) Then
				receiverDay = jsonInfo.receiverDay
			End If
			If Not isEmpty(jsonInfo.receiverGender) Then
				receiverGender = jsonInfo.receiverGender
			End If
			If Not isEmpty(jsonInfo.receiverForeign) Then
				receiverForeign = jsonInfo.receiverForeign
			End If
			If Not isEmpty(jsonInfo.receiverTelcoType) Then
				receiverTelcoType = jsonInfo.receiverTelcoType
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