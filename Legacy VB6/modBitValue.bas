Attribute VB_Name = "modBitValue"
Option Explicit

'byte
Public Const Bit1 As Byte = &H1
Public Const Bit2 As Byte = &H2
Public Const Bit3 As Byte = &H4
Public Const Bit4 As Byte = &H8
Public Const Bit5 As Byte = &H10
Public Const Bit6 As Byte = &H20
Public Const Bit7 As Byte = &H40
Public Const Bit8 As Byte = &H80


'########################################################################################
'#### Property BitWord ##################################################################
'########################################################################################
'#### Gets or sets an int data type's (ithis) bit at bbit index (Bit01, ...) to value ###
'########################################################################################

Public Property Let BitWord(ByRef iThis As Integer, ByRef bBit As Integer, ByRef nValue As Boolean)
    If (iThis And bBit) And (Not nValue) Then
        iThis = iThis - bBit
    ElseIf (Not (iThis And bBit)) And nValue Then
        iThis = iThis Or bBit
    End If
End Property
Public Property Get BitWord(ByRef iThis As Integer, ByRef bBit As Integer) As Boolean
    BitWord = (iThis And bBit)
End Property

