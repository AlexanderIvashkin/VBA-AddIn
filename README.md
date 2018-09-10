# VBA-AddIn
## VBA Addin with useful macros: regexps, networking-related (IP address/subnet mangling), VBA macro exporting &amp;c.

### Network-related functions

**Function AI_Is_IP_Address(ByVal StringToSearch As String, Optional CIDRMask As Boolean = False) As Boolean**

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

    Function AI_Get_IP_Address(ByVal StringToSearch As String) As String
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

**Function AI_Get_CIDR_PfxLen(ByVal StringToSearch As String) As String**

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

**Function AI_IP_CIDR_To_Mask(ByVal TextualCIDR As String) As String**
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Convert a textual CIDR into a subnet mask
    '' Arguments: textual representation of the CIDR notation.
    ''   Could be just the prefix length (with or without slash) or with the address
    ''   Examples: "/27" or "123.213.132.0/27" or "27"
    '' Returns:   textual representation of the subnet mask
    '' Alexander Ivashkin, 22 December 2017
    '' Updated version (to support mask without a slash): 16 January 2018
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

**Function AI_IP_NOT_Subnet(ByVal TextualSubnetMask As String) As String**

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' NOT's a subnet mask (flips 0 and 1s, useful for ACL matching)
    '' Arguments: textual representation of the subnet or wildcard mask
    '' Returns:   textual representation of the subnet or wildcard mask
    '' Alexander Ivashkin, 16 November 2017
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

**Function AI_IP_IsSubnet1InSubnet2(ByVal TextualSubnet1 As String, ByVal TextualSubnetMask1 As String, _
        ByVal IsSubnetMask1Wildcard As Boolean, ByVal TextualSubnet2 As String, ByVal TextualSubnetMask2 As String, ByVal IsSubnetMask2Wildcard As Boolean) As Boolean**
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Checks whether a subnet/mask will be within another subnet/mask
    '' Arguments: textual representations of the IP addresses / mask
    '' Returns:   boolean
    '' Alexander Ivashkin, 14-17 November 2017
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
**Function AI_IP_IsAddressInSubnet(ByVal TextualIP_Address As String, ByVal TextualSubnet_Number As String, ByVal TextualSubnetMask As String, ByVal IsSubnetMaskWildcard As Boolean) As Boolean**

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Checks whether an IP address will be within a subnet with a given mask
    '' Arguments: textual representation of the IP addresses / mask
    '' Returns:   boolean
    '' Alexander Ivashkin, 14 November 2017
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

**Function AI_IP_CalculateSubnet(ByVal TextualIP_Address As String, ByVal TextualSubnetMask As String) As String**

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Find a subnet out of textual IP address and subnet mask
    '' Returns a textual representation of subnet number
    '' Alexander Ivashkin, 14 November 2017
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
