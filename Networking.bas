Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' IP addressing and Networking-related custom formulae (user-invokable)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Version: 14 February 2018 :* :* :*
''



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Check that the string contains a valid IPv4 address (and, optionally, a CIDR mask).
'' Arguments: textual representation of an IPv4 address, with or without CIDR (depending on the second argument)
''
''   Examples of good arguments:
''    192.135.83.0/28
''    192.135.83.14
''   Examples of bad arguments (would return FALSE):
''    cc192.135.83.143
''    Next hop: 192.135.83.14
''    258.135.83.14
'' Returns:   TRUE or FALSE
'' Alexander Ivashkin, 25 January 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_Is_IP_Address(ByVal StringToSearch As String, Optional CIDRMask As Boolean = False) As Boolean

    ' A tiny regexp is all it takes...
    Dim sIsIPRegexp As String
    If CIDRMask Then
        sIsIPRegexp = "^(?:(?:(?:1?\d{1,2})|(?:2[0-4]\d)|(?:25[0-5]))\.){3}(?:(?:(?:1?\d{1,2})|(?:2[0-4]\d)|(?:25[0-5])))\/(?:\d|(?:[12]\d)|(?:3[012]))$"
    Else
        sIsIPRegexp = "^(?:(?:(?:(?:1?\d{1,2})|(?:2[0-4]\d)|(?:25[0-5]))\.){3}(?:(?:(?:1?\d{1,2})|(?:2[0-4]\d)|(?:25[0-5]))))$"
    End If

    AI_Is_IP_Address = AI_RegExp_IsMatch(StringToSearch, sIsIPRegexp)

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Extricate an IP address from a string.
'' Arguments: textual representation of a valid IPv4 address, with or without junk
''   Junk could be anything around the IP address but it can't be a digit or a dot.
''   Actually, my regexp would still work correctly,
''    but passing such jibberish to a function is wrong and shall be punished.
''
''   Examples of good arguments:
''    IP route: 192.135.83.0/28
''    192.135.83.14
''    az23f10.38.250.0
''    3dddd Next hop:  10.38.250.0eij 15
''   Examples of bad arguments (would return an empty string):
''    8192.135.83.14
''    192.135.83.14.adsfa
''    258.135.83.14
'' Returns:   textual representation of the IP address (or an empty string)
'' Alexander Ivashkin, 22 January 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_Get_IP_Address(ByVal StringToSearch As String) As String

    ' A tiny regexp is all it takes...
    ' My debugging and explanation entry: https://regex101.com/r/hdJdDD/5
    Const cGetIpAddress_rgx = "^(?:.*[^0-9.])?((?:(?:(?:1?\d{1,2})|(?:2[0-4]\d)|(?:25[0-5]))\.){3}(?:(?:(?:1?\d{1,2})|(?:2[0-4]\d)|(?:25[0-5]))))(?:[^0-9.].*)?$"

    AI_Get_IP_Address = AI_RegExp_GetSubMatch(StringToSearch, cGetIpAddress_rgx)

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Extricate a CIDR prefix length from a string containing an IPv4 address and a mask.
'' Arguments: textual representation of a valid IPv4 address with a valid mask, with or without junk
''   Junk could be anything around the IP address but it can't be a digit or a dot.
''   Actually, my regexp would still work correctly,
''    but passing such jibberish to a function is wrong and shall be punished.
''
''   Examples of good arguments:
''    192.135.83.14/23
''    az23f10.38.250.0/10
''    Subnet: 3dddd  10.38.250.0/30eij 15
''   Examples of bad arguments (would return an empty string):
''    8192.135.83.14
''    192.135.83.14.adsfa
''    258.135.83.14
''    192.135.83.14
''    192.135.83.14/38
'' Returns:   textual representation of the CIDR pfx len (or an empty string)
'' Alexander Ivashkin, 22 January 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_Get_CIDR_PfxLen(ByVal StringToSearch As String) As String

    ' A tiny regexp is all it takes...
    ' My entry on the Regex101: https://regex101.com/r/iV031f/2
    Const cGetCIDR_PfxLen_rgx = "^(?:.*[^0-9.])?(?:(?:(?:1?\d{1,2})|(?:2[0-4]\d)|(?:25[0-5]))\.){3}(?:(?:(?:1?\d{1,2})|(?:2[0-4]\d)|(?:25[0-5])))\/((?:[12]?\d)|(?:3[0-2]))(?:[^0-9.].*)?$"

    AI_Get_CIDR_PfxLen = AI_RegExp_GetSubMatch(StringToSearch, cGetCIDR_PfxLen_rgx)

End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Convert a textual CIDR into a subnet mask
'' Arguments: textual representation of the CIDR notation.
''   Could be just the prefix length (with or without slash) or with the address
''   Examples: "/27" or "123.213.132.0/27" or "27"
'' Returns:   textual representation of the subnet mask
'' Alexander Ivashkin, 22 December 2017
'' Updated version (to support mask without a slash): 16 January 2018
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_IP_CIDR_To_Mask(ByVal TextualCIDR As String) As String
    ' Let's strip the prefix length
    Dim sPfxLen As String
    sPfxLen = AI_RegExp_GetSubMatch(TextualCIDR, "^(?:(?:(?:\d{1,3}\.){3}\d{1,3}\/)?|(?:\/)?)(\d{1,2})$", False, 0, 0)
    
    ' Yes, this is going to be a lookup table.
    ' Gross? Maybe. But works and is probably much faster then binary math / array iteration.
    ' That's Bill Gates's favourite language, don't expect any high performance!
    
    Select Case sPfxLen
        Case "0"
                AI_IP_CIDR_To_Mask = "0.0.0.0"
        Case "1"
                AI_IP_CIDR_To_Mask = "128.0.0.0"
        Case "2"
                AI_IP_CIDR_To_Mask = "192.0.0.0"
        Case "3"
                AI_IP_CIDR_To_Mask = "224.0.0.0"
        Case "4"
                AI_IP_CIDR_To_Mask = "240.0.0.0"
        Case "5"
                AI_IP_CIDR_To_Mask = "248.0.0.0"
        Case "6"
                AI_IP_CIDR_To_Mask = "252.0.0.0"
        Case "7"
                AI_IP_CIDR_To_Mask = "254.0.0.0"
        Case "8"
                AI_IP_CIDR_To_Mask = "255.0.0.0"
        Case "9"
                AI_IP_CIDR_To_Mask = "255.128.0.0"
        Case "10"
                AI_IP_CIDR_To_Mask = "255.192.0.0"
        Case "11"
                AI_IP_CIDR_To_Mask = "255.224.0.0"
        Case "12"
                AI_IP_CIDR_To_Mask = "255.240.0.0"
        Case "13"
                AI_IP_CIDR_To_Mask = "255.248.0.0"
        Case "14"
                AI_IP_CIDR_To_Mask = "255.252.0.0"
        Case "15"
                AI_IP_CIDR_To_Mask = "255.254.0.0"
        Case "16"
                AI_IP_CIDR_To_Mask = "255.255.0.0"
        Case "17"
                AI_IP_CIDR_To_Mask = "255.255.128.0"
        Case "18"
                AI_IP_CIDR_To_Mask = "255.255.192.0"
        Case "19"
                AI_IP_CIDR_To_Mask = "255.255.224.0"
        Case "20"
                AI_IP_CIDR_To_Mask = "255.255.240.0"
        Case "21"
                AI_IP_CIDR_To_Mask = "255.255.248.0"
        Case "22"
                AI_IP_CIDR_To_Mask = "255.255.252.0"
        Case "23"
                AI_IP_CIDR_To_Mask = "255.255.254.0"
        Case "24"
                AI_IP_CIDR_To_Mask = "255.255.255.0"
        Case "25"
                AI_IP_CIDR_To_Mask = "255.255.255.128"
        Case "26"
                AI_IP_CIDR_To_Mask = "255.255.255.192"
        Case "27"
                AI_IP_CIDR_To_Mask = "255.255.255.224"
        Case "28"
                AI_IP_CIDR_To_Mask = "255.255.255.240"
        Case "29"
                AI_IP_CIDR_To_Mask = "255.255.255.248"
        Case "30"
                AI_IP_CIDR_To_Mask = "255.255.255.252"
        Case "31"
                AI_IP_CIDR_To_Mask = "255.255.255.254"
        Case "32"
                AI_IP_CIDR_To_Mask = "255.255.255.255"
        Case Else
                AI_IP_CIDR_To_Mask = ""
    End Select

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' NOT's a subnet mask (flips 0 and 1s, useful for ACL matching)
'' Arguments: textual representation of the subnet or wildcard mask
'' Returns:   textual representation of the subnet or wildcard mask
'' Alexander Ivashkin, 16 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_IP_NOT_Subnet(ByVal TextualSubnetMask As String) As String
    
    AI_IP_NOT_Subnet = privAI_IP_FromArrayToString(privAI_IP_BitwiseNOT_Array(AI_ConvertTextualIP_AddressToArray(TextualSubnetMask)))

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Checks whether a subnet/mask will be within another subnet/mask
'' Arguments: textual representations of the IP addresses / mask
'' Returns:   boolean
'' Alexander Ivashkin, 14-17 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_IP_IsSubnet1InSubnet2(ByVal TextualSubnet1 As String, ByVal TextualSubnetMask1 As String, _
    ByVal IsSubnetMask1Wildcard As Boolean, ByVal TextualSubnet2 As String, ByVal TextualSubnetMask2 As String, ByVal IsSubnetMask2Wildcard As Boolean) As Boolean
    
    Dim aSubnet1() As Byte
    Dim aSubnetMask1() As Byte
    Dim aSubnet2() As Byte
    Dim aSubnetMask2() As Byte
    
    aSubnet1 = AI_ConvertTextualIP_AddressToArray(TextualSubnet1)
    aSubnet2 = AI_ConvertTextualIP_AddressToArray(TextualSubnet2)
    
    If IsSubnetMask1Wildcard Then
        aSubnetMask1 = privAI_IP_BitwiseNOT_Array(AI_ConvertTextualIP_AddressToArray(TextualSubnetMask1))
    Else
        aSubnetMask1 = AI_ConvertTextualIP_AddressToArray(TextualSubnetMask1, 1)
    End If
    
    If IsSubnetMask2Wildcard Then
        aSubnetMask2 = privAI_IP_BitwiseNOT_Array(AI_ConvertTextualIP_AddressToArray(TextualSubnetMask2))
    Else
        aSubnetMask2 = AI_ConvertTextualIP_AddressToArray(TextualSubnetMask2, 1)
    End If
    
    
    Dim aFirstIP_InSubnet() As Byte
    Dim aLastIP_InSubnet() As Byte
    Dim bIsFirstIPInSubnet2 As Boolean
    Dim bIsLastIPInSubnet2 As Boolean
    
    
    aFirstIP_InSubnet = privAI_IP_FindSubnetArray(aSubnet1, aSubnetMask1)
    bIsFirstIPInSubnet2 = privAI_IP_IsAddressInSubnetArray(aFirstIP_InSubnet, aSubnet2, aSubnetMask2)
    
    aLastIP_InSubnet = privAI_IP_BitwiseOR_Array(privAI_IP_BitwiseNOT_Array(aSubnetMask1), aSubnet1)
    bIsLastIPInSubnet2 = privAI_IP_IsAddressInSubnetArray(aLastIP_InSubnet, aSubnet2, aSubnetMask2)
    
    AI_IP_IsSubnet1InSubnet2 = bIsFirstIPInSubnet2 And bIsLastIPInSubnet2
        
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Checks whether an IP address will be within a subnet with a given mask
'' Arguments: textual representation of the IP addresses / mask
'' Returns:   boolean
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_IP_IsAddressInSubnet(ByVal TextualIP_Address As String, ByVal TextualSubnet_Number As String, ByVal TextualSubnetMask As String, ByVal IsSubnetMaskWildcard As Boolean) As Boolean
    Dim aSubnetNumber1() As Byte
    Dim aSubnetNumber2() As Byte
    Dim aIP_Add() As Byte
    Dim aSubnetNumber() As Byte
    Dim aSubnetMask() As Byte
    Dim i As Integer
    
    aIP_Add = AI_ConvertTextualIP_AddressToArray(TextualIP_Address)
    aSubnetNumber = AI_ConvertTextualIP_AddressToArray(TextualSubnet_Number)
    If IsSubnetMaskWildcard Then
        aSubnetMask = privAI_IP_BitwiseNOT_Array(AI_ConvertTextualIP_AddressToArray(TextualSubnetMask))
    Else
        aSubnetMask = AI_ConvertTextualIP_AddressToArray(TextualSubnetMask, 1)
    End If
    
    'aSubnetMask = AI_ConvertTextualIP_AddressToArray(TextualSubnetMask)
    
    AI_IP_IsAddressInSubnet = privAI_IP_IsAddressInSubnetArray(aIP_Add, aSubnetNumber, aSubnetMask)
    
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Find a subnet out of textual IP address and subnet mask
'' Returns a textual representation of subnet number
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function AI_IP_CalculateSubnet(ByVal TextualIP_Address As String, ByVal TextualSubnetMask As String) As String
    Dim aSubnetNumber() As Byte
    Dim aIP_Add() As Byte
    Dim aSubnetMask() As Byte
    Dim i As Integer
    
    aIP_Add = AI_ConvertTextualIP_AddressToArray(TextualIP_Address)
    aSubnetMask = AI_ConvertTextualIP_AddressToArray(TextualSubnetMask, 1)
    aSubnetNumber = privAI_IP_FindSubnetArray(aIP_Add, aSubnetMask)
    
    AI_IP_CalculateSubnet = privAI_IP_FromArrayToString(aSubnetNumber)
    
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Supporting functions (non user-invokable)




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Checks whether an IP address will be within a subnet with a given mask
'' Arguments: arrays of Bytes
'' Returns:   boolean
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function privAI_IP_IsAddressInSubnetArray(ByRef IP_Address() As Byte, ByRef SubnetNumber() As Byte, ByRef SubnetMask() As Byte) As Boolean
    Dim aSubnetNumber1() As Byte
    Dim aSubnetNumber2() As Byte
    Dim i As Integer
    
    aSubnetNumber1 = privAI_IP_FindSubnetArray(IP_Address, SubnetMask)
    aSubnetNumber2 = privAI_IP_FindSubnetArray(SubnetNumber, SubnetMask)
    
    For i = 0 To 3
        If aSubnetNumber1(i) <> aSubnetNumber2(i) Then
            privAI_IP_IsAddressInSubnetArray = False
            Exit Function
        End If
    Next i
    
    privAI_IP_IsAddressInSubnetArray = True
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Convert an IP address from an array into string
'' Input:  an array of bytes
'' Output: a string
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function privAI_IP_FromArrayToString(ByRef IP_Address1() As Byte) As String
    Dim sResult As String
    Dim i As Integer
    
    For i = 0 To 2
        sResult = sResult + Trim(Str$(IP_Address1(i))) + "."
    Next i
    sResult = sResult + Trim(Str$(IP_Address1(i)))
    
    privAI_IP_FromArrayToString = sResult
        
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Bitwise NOT an IP address
'' Input:  an array of bytes
'' Output: an array of bytes
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function privAI_IP_BitwiseNOT_Array(ByRef IP_Address1() As Byte) As Variant
    Dim aResult(3) As Byte
    Dim i As Integer
    
    For i = 0 To 3
        aResult(i) = Not IP_Address1(i)
    Next i
    
    privAI_IP_BitwiseNOT_Array = aResult
        
End Function



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Bitwise OR two IP addresses
'' Input:  two arrays of bytes
'' Output: an array of bytes
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function privAI_IP_BitwiseOR_Array(ByRef IP_Address1() As Byte, ByRef IP_Address2() As Byte) As Variant
    Dim aResult(3) As Byte
    Dim i As Integer
    
    For i = 0 To 3
        aResult(i) = IP_Address1(i) Or IP_Address2(i)
    Next i
    
    privAI_IP_BitwiseOR_Array = aResult
        
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Bitwise AND two IP addresses
'' Input:  two arrays of bytes
'' Output: an array of bytes
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function privAI_IP_BitwiseAND_Array(ByRef IP_Address1() As Byte, ByRef IP_Address2() As Byte) As Variant
    Dim aResult(3) As Byte
    Dim i As Integer
    
    For i = 0 To 3
        aResult(i) = IP_Address1(i) And IP_Address2(i)
    Next i
    
    privAI_IP_BitwiseAND_Array = aResult
        
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Bitwise AND two textual IP addresses
'' Output: an array of bytes
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function privAI_IP_BitwiseAND_Textual(ByVal TextualIP_Address1 As String, ByVal TextualIP_Address2 As String) As Variant
    Dim aIP_Add1 As Variant
    Dim aIP_Add2 As Variant
    Dim aResult(3) As Byte
    Dim i As Integer
    
    aIP_Add1 = AI_ConvertTextualIP_AddressToArray(TextualIP_Address1)
    aIP_Add2 = AI_ConvertTextualIP_AddressToArray(TextualIP_Address2)

    For i = 0 To 3
        aResult(i) = aIP_Add1(i) And aIP_Add2(i)
    Next i
        
    privAI_IP_BitwiseAND_Textual = aResult
        
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Find a subnet out of textual IP address and subnet mask
'' Output: an array of Bytes
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function privAI_IP_FindSubnetTextual(ByVal TextualIP_Address As String, ByVal TextualSubnetMask As String) As Variant
    privAI_IP_FindSubnetTextual = privAI_IP_BitwiseAND_Textual(TextualIP_Address, TextualSubnetMask)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Find a subnet out of an IP address in array form and subnet mask
'' Output: an array of Bytes
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function privAI_IP_FindSubnetArray(ByRef IP_Address() As Byte, ByRef SubnetMask() As Byte) As Variant
    privAI_IP_FindSubnetArray = privAI_IP_BitwiseAND_Array(IP_Address, SubnetMask)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Convert a textual representation of an IP address into numerical form
'' Arguments:
''  TextualIP_Address - a string representation of an IP address
''  ArgumentType
''    0: this is an IPv4 address (default)
''    1: this is a subnet mask
''    2: this is a wildcard mask
'' Output:    an array of bytes (0,3)
'' Alexander Ivashkin, 14 November 2017
'' Upgraded: 16 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AI_ConvertTextualIP_AddressToArray(ByVal TextualIP_Address As String, Optional ByVal ArgumentType As Byte = 0) As Variant

    Const cRegExp_IsIP_Add = "^([1-2]??\d{1,2})\.([1-2]??\d{1,2})\.([1-2]??\d{1,2})\.([1-2]??\d{1,2})$"
    
    If AI_RegExp_IsMatch(TextualIP_Address, cRegExp_IsIP_Add) = False Then
        AI_ConvertTextualIP_AddressToArray = False
        Exit Function
    End If

On Error GoTo Hell
    
    Dim i As Integer
    Dim aIP_AddressArray(3) As Byte
    Dim dTmp As Double
    Dim bIsPreviousOctetZero As Boolean
    bIsPreviousOctetZero = False
    
    For i = 0 To 3
        dTmp = Val(AI_RegExp_GetSubMatch(TextualIP_Address, cRegExp_IsIP_Add, False, 0, i))
        
        ' Sanity checks (is this really an address / mask?)
        Select Case ArgumentType
            ' IPv4 address, Wildcard mask
            Case 0, 2:
                If dTmp > 255 Then
                    AI_ConvertTextualIP_AddressToArray = False
                    Exit Function
                End If
            ' Subnet mask
            Case 1:
                If (dTmp <> 0 And bIsPreviousOctetZero = True) Or (Not privAI_IP_IsSubnetOctetSane(dTmp, False)) Then
                    AI_ConvertTextualIP_AddressToArray = False
                    Exit Function
                End If
                If dTmp = 0 Then bIsPreviousOctetZero = True
            Case Else
                AI_ConvertTextualIP_AddressToArray = False
                Exit Function
        End Select
        
        aIP_AddressArray(i) = dTmp
        'Debug.Print aIP_AddressArray(i)
    Next i
    
    AI_ConvertTextualIP_AddressToArray = aIP_AddressArray
        
    Exit Function

Hell:
    Debug.Print "Something went wrong in AI_ConvertTextualIP_AddressToArray: ", Err.Number, "  ", Err.Description, "  ", Err.Source
    AI_ConvertTextualIP_AddressToArray = False
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Performs a sanity check on an octet byte
'' Arguments:
''  SubnetOctet - a byte (one octet)
''  IsWildcard
''    False: this is a subnet mask (default)
''    True:  this is a wildcard mask
'' Output:    boolean
'' Alexander Ivashkin, 16 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Private Function privAI_IP_IsSubnetOctetSane(ByVal SubnetOctet As Byte, Optional ByVal IsWildcard = False) As Boolean
    ' Now I should probably have used some clever trick (an inline ASM, a SIMD instruction or at least iterating through an array, especially with GCC's -funroll-loops)
    ' But this works and is probably faster

    If IsWildcard Then
        ' Actually, if we've arrived here and not caused this function to fail by passing a wrong argument (not a byte),
        ' then this check is redundant. I will leave it though for clarity and safety. You never know how BASIC would work!
        If SubnetOctet < 0 Or SubnetOctet > 255 Then
            privAI_IP_IsSubnetOctetSane = False
            Exit Function
        End If
    Else
        Select Case SubnetOctet
            Case 0
            Case 255
            Case 254
            Case 252
            Case 248
            Case 240
            Case 224
            Case 192
            Case 128
            Case Else
                privAI_IP_IsSubnetOctetSane = False
                Exit Function
        End Select
    End If
    
    privAI_IP_IsSubnetOctetSane = True
    
End Function
