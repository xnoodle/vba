Private Function GetUserFirstName()
    Dim objSysInfo, strUser, objUser

    Set objSysInfo = CreateObject("ADSystemInfo")
    strUser = objSysInfo.UserName
    Set objUser = GetObject("LDAP://" & strUser)
    GetUserFirstName = objUser.givenName

End Function
