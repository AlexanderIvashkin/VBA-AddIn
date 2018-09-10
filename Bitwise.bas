Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Bitwise logical operations
'' Alexander Ivashkin, 14 November 2017
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function AI_BitwiseAND(ByVal FirstOperand As Long, ByVal SecondOperand As Long) As Long
    AI_BitwiseAND = FirstOperand And SecondOperand
End Function

Function AI_BitwiseOR(ByVal FirstOperand, ByVal SecondOperand) As Long
    AI_BitwiseOR = FirstOperand Or SecondOperand
End Function

Function AI_BitwiseXOR(ByVal FirstOperand, ByVal SecondOperand) As Long
    AI_BitwiseXOR = FirstOperand Xor SecondOperand
End Function

Function AI_BitwiseNOT(ByVal FirstOperand) As Long
    AI_BitwiseNOT = Not FirstOperand
End Function
