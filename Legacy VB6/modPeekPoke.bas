Attribute VB_Name = "modPeekPoke"
Option Explicit

Public Const LPT1_32 = "&HD800"
Public Const LPT1_64 = "&H3FF8"

Public Declare Function Inp32 Lib "inpout32.dll" (ByVal port As Integer) As Integer
Public Declare Sub Out32 Lib "inpout32.dll" (ByVal port As Integer, ByVal Info As Integer)

Public Function CoinCheck() As Boolean
    On Error GoTo exitsub:
    
    Dim b As Integer
    b = 0

    BitWord(b, Bit1) = True
    BitWord(b, Bit2) = True
    BitWord(b, Bit3) = True
    BitWord(b, Bit4) = True
    BitWord(b, Bit5) = True
    BitWord(b, Bit6) = True
    BitWord(b, Bit7) = True
    BitWord(b, Bit8) = True

    Out32 CInt(Val(LPT1_64)), b
    
    b = Inp32(CInt(Val(LPT1_64)))

    If Not BitWord(b, Bit7) Then
        CoinCheck = True
    End If

   ' Debug.Print BitWord(b, Bit1) & " " & BitWord(b, Bit2) & " " & BitWord(b, Bit3) & " " & BitWord(b, Bit4) & " " & _
   '              BitWord(b, Bit5) & " " & BitWord(b, Bit6) & " " & BitWord(b, Bit7) & " " & BitWord(b, Bit8)
    
    Exit Function
exitsub:
    MyDebugPrint "CoinCheck() Error: " & Err.Description
Err.Clear
End Function

