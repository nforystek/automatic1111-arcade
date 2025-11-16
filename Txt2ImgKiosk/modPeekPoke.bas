Attribute VB_Name = "modPeekPoke"
Option Explicit

'###############################################################################
'#########  FOR THE COIN OPERATION, THIS REQUIRES A PARALLEL PORT ##############
'
' If you have a parallel port you can find the LPT1 address defined in hes below
' from the "Device Manager" by selecting your Parallel port's driver properties.
'
' By setting the value of B to all True, those data pins that are grounded (i.e.
' the normally open circut closes via the coin box switch), where I'm using pin7
' will then be False upon reading the port back into the B variable.
'
'##############################################################################

'Public Const LPT1 = "&HD800" 'this is what it is on my XP system
Public Const LPT1 = "&H3FF8" 'this is what it is on my Windows 11 system

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

    Out32 CInt(Val(LPT1)), b
    
    b = Inp32(CInt(Val(LPT1)))

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

