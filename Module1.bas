Attribute VB_Name = "Module1"
Sub main()
    
    '// See if 192.168.2.10 exists in 192.168.1.0/255.255.255.0
    MsgBox IsInMask("192.168.1.0", "255.255.255.0", "192.168.2.10")
    
    '// Let's have a look at an IP address in binary anyway
    MsgBox ConvertToBinary("192.168.1.0")
End Sub


Function IsInMask(ByVal BaseIP As String, ByVal NetMask As String, ByVal CompareIP As String) As Boolean
    
    '// Function:  See if CompareIP is within the network range described
    '//            by BaseIP and NetMask, and return a boolean result.
    '//
    '// Inputs: BaseIP - The networks address.
    '//         NetMask - The network mask.
    '//         CompareIP - The IP address of the computer you are
    '//                  testing against the network address/mask
    
    Dim iIndex As Integer
    
    '// First we need to change each address to binary.
    BaseIP = ConvertToBinary(BaseIP)
    NetMask = ConvertToBinary(NetMask)
    CompareIP = ConvertToBinary(CompareIP)
    
    '// Now we need to compare significant bits which
    '// are indicated by 1's in the network mask.
    
    '// set the default
    IsInMask = True
    
    For iIndex = 1 To Len(NetMask)
        If Mid(NetMask, iIndex, 1) = "1" Then
            '// We need to compare
            If Mid(BaseIP, iIndex, 1) <> Mid(CompareIP, iIndex, 1) Then _
                IsInMask = False
        End If
    Next
    
End Function


Function ConvertToBinary(ByVal IPaddr As String) As String
    '// Convert an IP address into binary.
    Dim sIPaddr() As String
    
    Dim iIndex As Integer
    Dim iCount As Integer
    
    sIPaddr = Split(IPaddr, ".")
    
    For iIndex = 0 To UBound(sIPaddr)
        
        '// for each byte change to 8 character binary
        
        For iCount = 7 To 0 Step -1
            
            If CInt(sIPaddr(iIndex)) >= 2 ^ iCount Then
                ConvertToBinary = ConvertToBinary & "1"
                sIPaddr(iIndex) = CInt(sIPaddr(iIndex)) - 2 ^ iCount
            Else
                ConvertToBinary = ConvertToBinary & "0"
            End If
        
        Next
    Next
    
End Function


