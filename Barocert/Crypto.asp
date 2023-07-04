<%
Class Encryptor
    private m_Key
    private m_Mode
	Private m_encrypt

    Public Property Let Key(ByVal value)
        m_Key = value
    End Property
    Public Property Let Mode(ByVal value)
        m_Mode = value
    End Property
	
	Public Sub Class_Initialize
		Set m_encrypt = GetObject( "script:" & Request.ServerVariables("APPL_PHYSICAL_PATH") + "barocert" & "\cryptojs.wsc" )
	End Sub

	Public Sub Initialize(SecretKey)
		m_Key = SecretKey
	End Sub
	
	Public Sub Class_Terminate
		Set m_sha1 = Nothing 
	End Sub 

    Public Function enc(plainText)
        enc = m_encrypt.ivAddAES256CBC(plainText , m_Key)
    End Function

End Class
%>