# VBA-AddIn
## VBA Addin with useful macros: regexps, networking-related (IP address/subnet mangling), VBA macro exporting &amp;c.

### Network-related functions

**AI_Is_IP_Address**

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

**AI_Get_IP_Address**

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

**AI_Get_CIDR_PfxLen**

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

**AI_IP_CIDR_To_Mask**
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Convert a textual CIDR into a subnet mask
    '' Arguments: textual representation of the CIDR notation.
    ''   Could be just the prefix length (with or without slash) or with the address
    ''   Examples: "/27" or "123.213.132.0/27" or "27"
    '' Returns:   textual representation of the subnet mask
    '' Alexander Ivashkin, 22 December 2017
    '' Updated version (to support mask without a slash): 16 January 2018
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

**AI_IP_NOT_Subnet**

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' NOT's a subnet mask (flips 0 and 1s, useful for ACL matching)
    '' Arguments: textual representation of the subnet or wildcard mask
    '' Returns:   textual representation of the subnet or wildcard mask
    '' Alexander Ivashkin, 16 November 2017
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

**AI_IP_IsSubnet1InSubnet2**
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Checks whether a subnet/mask will be within another subnet/mask
    '' Arguments: textual representations of the IP addresses / mask
    '' Returns:   boolean
    '' Alexander Ivashkin, 14-17 November 2017
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
**AI_IP_IsAddressInSubnet**

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Checks whether an IP address will be within a subnet with a given mask
    '' Arguments: textual representation of the IP addresses / mask
    '' Returns:   boolean
    '' Alexander Ivashkin, 14 November 2017
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

**AI_IP_CalculateSubnet**

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '' Find a subnet out of textual IP address and subnet mask
    '' Returns a textual representation of subnet number
    '' Alexander Ivashkin, 14 November 2017
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

### Strings and regexps (g-strings.bas)

```
AI_MATCH_Regexp: Regexp version of MATCH (match a regexp against a range of strings)
AI_MATCH_Regexps:Regexp version of MATCH - another version: match a string against an array of regexps
AI_RegExp_IsMatch: Check if a regexp matches
AI_RegExp_GetSubMatch: Get a submatch from a regexp
```

**AI_ParseMagicSymbols: Replace magic symbols (placeholders) with dynamic data**

```
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Arguments: a string full of magic.
''
'' Placeholders consist of one symbol prepended with a %:
''    %d - current date
''    %t - current time
''    %u - username (user ID)
''    %n - full user name (usually name and surname)
''    %% - literal % (placeholder escape)
''    Using an unsupported magic symbol will treat the % literally, as if it had been escaped.
''    A single placeholder terminating the string will also be treated literally.
''    Magic symbols are case-sensitive.
''
'' Returns:   A string with no magic but with lots of beauty.
''
'' Examples:
'' "Today is %d" becomes "Today is 2018-01-26"
'' "Beautiful time: %%%t%%" yields "Beautiful time: %16:10:51%"
'' "There are %zero% magic symbols %here%.", true to its message, outputs "There are %zero% magic symbols %here%."
'' "%%% looks lovely %%%" would show "%% looks lovely %%" - one % for the escaped "%%" and the second one for the unused "%"!
''
'' Alexander Ivashkin, 26 January 2018
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
```

### Bitwise logical operations (bitwise.bas)

### Miscellaneous functions (misc.bas)

```
ICanHazSheet: check if a sheet exists (by name)
ClearAndResizeTable: clear and resize a named table on the active sheet then move the selection to the first cell of the table
AI_SendRangeAsHTML: Send a range as an HTML mail
AI_GetLastShitName: Get a name of the last sheet in the active workbook.
```
