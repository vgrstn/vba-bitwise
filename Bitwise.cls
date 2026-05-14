VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bitwise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Bitwise Operations Class."
'@IgnoreModule AssignedByValParameter, VariableNotAssigned, UnassignedVariableUsage, ProcedureNotUsed, UseMeaningfulName, MultipleDeclarations
'@PredeclaredId
'@Folder("Class")
'@ModuleDescription ("Bitwise Operations Class.")

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Bitwise Operations Class
' Enumeration: For Each Bit in object.BitArray(bits)
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.09.09
'
' The Bitwise Class is predeclared which creates a global default instance of
' the object with the same name making the object immediately accessible.
'
' The bits argument is used as variable name for a LongPtr bit pattern. A bit is
' a Boolean and an array of bits is a Boolean array. A bitstring is obviously a
' string consisting of 0's and 1's.
'
' All Methods are Functions returning a LongPtr except for the Count Function
' which returns a Long.
'
' Methods
' Count(bits)            Count the number of set bits in a bit pattern
' Flip(bits, No)         Flips a selected bit in a bit pattern
' Swap(bits, No1,No2)    Swaps two selected bits in a bit pattern
' Inverse(bits)          Inverses a bit pattern
' Reverse(bits)          Reverses a bit pattern
' Clear(bits, No)        Clears a bit in a bit pattern
' LClear(bits, No)       Clears the bit pattern left from a selected bit
' RClear(bits, No)       Clears the bit pattern right from a selected bit
' LShift(bits, amount)   Shifts a bit pattern to the left by an amount
' RShift(bits, amount)   Shifts a bit pattern to the right by an amount
' Shift(bits, amount)    Shifts a bit pattern by an amount
' Rotate(bits, amount)   Rotates a bit pattern by an amount
' Random()               Creates a random bit pattern
'
' Properties
' Bit(bits, No)          Returns or sets a selected bit in a bit pattern
' BitArray(bits)         Converts a bit pattern to or from an array of bits
' BitString(bits)        Converts a bit pattern to or from a bitstring
'
' A multiplication by n where n = 2 ^ i is equivalent to a left shift by i bit
' positions. For a left shift overflow by multiplication needs to be prevented.
' Likewise, a division by n where n = 2 ^ i is equivalent to a right shift by i
' positions. VB uses two's complements representation for signed integers. Hence,
' for negative integers, division doesn't correspond to a right shift. In the
' latter case the signbit needs to be masked before the division and subsequently
' set depending on whether the bit that is shifted to the signbit is set or not.
'
' Precalculated bitmasks are used for efficiency and are initialized at object
' instantiation. The Count, Reverse and BitString Function are relatively slow
' because for these routines require bit by bit processing. However, the use of
' precalculated byte lookup tables makes these functions 3-4 times faster.
'
' Method       time (ms)
' Count          0.00022 (with precalculated table)
' Flip           0.00007
' Inverse        0.00007
' Reverse        0.00033 (with precalculated table)
' Shift          0.00011
' Rotate         0.00014
' Random         0.00022
' Swap           0.00010
' BitString      0.00034 (with precalculated table)
'
' The standard module is typically 10-20% faster than the class implementation.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compiler directives
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The allocated memory can be minimized by switching off the StringTable lookup
' table which uses more memory than all the other tables and masks together.

#Const UseStringTable = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' VB uses two's complement representation for integers.
#If Win64 Then
Private Const SignBit As Long = 63
#Else
Private Const SignBit As Long = 31
#End If

' To be initialized by InitializeMasks
Private SignBitMask As LongPtr, BitMask16 As LongPtr

' BitMask(i) is a bit mask with the ith bit set.
' Allocated memory for BitMask() is 32 + 64 * 8 = 544 bytes (x64).
Private BitMask(0 To SignBit) As LongPtr

' MaxMask(i) is a bit mask with all jth bits set with j < i.
' Allocated memory for MaxMask() is 32 + 64 * 8 = 544 bytes (x64).
Private MaxMask(0 To SignBit) As LongPtr

' ReverseTable(n) is the reversed byte for byte n.
' Allocated memory for ReverseTable() is 32 + &H100 * 1 = 288 bytes (x64).
Private ReverseTable(0 To &HFF) As Byte

' CountTable(n) is the number of bits set in byte n.
' Allocated memory for CountTable() is 32 + &H100 * 1 = 288 bytes (x64).
Private CountTable(0 To &HFF) As Byte

#If UseStringTable Then
' StringTable(n) is the bitstring corresponding to byte n.
' Allocated memory for StringTable() is 32 + &H100 * 22 = 5664 bytes (x64).
Private StringTable(0 To &HFF) As String
#End If

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Class methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Class_Initialize()

    Randomize
    InitializeMasks
    InitializeTables

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Returns the number of set bits in a bit pattern.")
Public Function Count(ByVal bits As LongPtr) As Long

    ' Count and mask the SignBit if it is set.
    If bits < 0 Then
        Count = 1
        bits = bits And Not SignBitMask
    End If

    ' Sum the precalculated byte counts.
    Dim i As Long
    For i = 0 To SignBit \ 8
        Count = Count + CountTable(CLng(bits And &HFF))
        bits = bits \ &H100
    Next

End Function


'@Description("Flips a selected bit in a bit pattern and returns the outcome.")
Public Function Flip(ByVal bits As LongPtr, ByVal No As Long) As LongPtr

    Flip = bits Xor BitMask(No)

End Function


'@Description("Swaps two selected bits in a bit pattern and returns the outcome.")
Public Function Swap(ByVal bits As LongPtr, ByVal No1 As Long, ByVal No2 As Long) As LongPtr

    ' Save the mask for the two bits to be swapped.
    Dim Mask As LongPtr: Mask = BitMask(No1) Or BitMask(No2)

    ' Swap the two bits if they are different.
    Select Case bits And Mask
    Case 0, Mask
        Swap = bits
    Case Else
        Swap = bits Xor Mask
    End Select

End Function


'@Description("Inverses a bit pattern and returns the outcome.")
Public Function Inverse(ByVal bits As LongPtr) As LongPtr

    Inverse = Not bits

End Function


'@Description("Reverses a bit pattern and returns the outcome.")
Public Function Reverse(ByVal bits As LongPtr) As LongPtr

    ' Bit 0 is masked alongside SignBit to prevent multiplication overflow in the
    ' accumulation loop: without it ReverseTable(byte0) can reach 0xFF, pushing
    ' the 7-byte intermediate past 0x7FFFFFFFFFFFFF and overflowing on * &H100.
    Dim Mask As LongPtr: Mask = bits And (1 Or SignBitMask)

    ' Mask the 0th bit and the SignBit.
    bits = bits And Not (1 Or SignBitMask)

    ' Combine the precalculated reversed bytes.
    Dim i As Long
    For i = 0 To SignBit \ 8
        Reverse = &H100 * Reverse Or ReverseTable(CLng(bits And &HFF))
        bits = bits \ &H100
    Next

    ' Restore the two masked bits and swap them if they are different.
    Select Case Mask
    Case 1, SignBitMask
        Reverse = Reverse Or (Mask Xor (1 Or SignBitMask))
    Case Else
        Reverse = Reverse Or Mask
    End Select

End Function


'@Description("Clears a selected bit in a bit pattern and returns the outcome.")
Public Function Clear(ByVal bits As LongPtr, ByVal No As Long) As LongPtr

    Clear = bits And Not BitMask(No)

End Function


'@Description("Clears a bit pattern left from a selected bit and returns the outcome.")
Public Function LClear(ByVal bits As LongPtr, ByVal No As Long) As LongPtr

    LClear = bits And (BitMask(No) Or MaxMask(No))

End Function


'@Description("Clears a bit pattern right from a selected bit and returns the outcome.")
Public Function RClear(ByVal bits As LongPtr, ByVal No As Long) As LongPtr

    RClear = bits And Not MaxMask(No)

End Function


'@Description("Shifts a bit pattern to the left by an amount and returns the outcome.")
Public Function LShift(ByVal bits As LongPtr, ByVal amount As Byte) As LongPtr

    If amount > SignBit Then Exit Function

    If amount = 0 Then
        LShift = bits
    ElseIf (bits And BitMask(SignBit - amount)) = 0 Then
        ' Clear the overflow bits before multiplying.
        LShift = (bits And MaxMask(SignBit - amount)) * BitMask(amount)
    Else
        ' Clear the overflow bits before multiplying and set the signbit.
        LShift = (bits And MaxMask(SignBit - amount)) * BitMask(amount) Or SignBitMask
    End If

End Function


'@Description("Shifts a bit pattern to the right by an amount and returns the outcome.")
Public Function RShift(ByVal bits As LongPtr, ByVal amount As Byte) As LongPtr
' RShiftBits is effectively a right logical shift, it fills the vacated bits with 0.

    If amount > SignBit Then Exit Function

    If amount = 0 Then
        RShift = bits
    ElseIf bits < 0 Then
        ' Clear the signbit before dividing and set the shifted signbit.
        RShift = (bits And Not SignBitMask) \ BitMask(amount) Or BitMask(SignBit - amount)
    Else
        RShift = bits \ BitMask(amount)
    End If

End Function


'@Description("Shifts a bit pattern by an amount and returns the outcome.")
Public Function Shift(ByVal bits As LongPtr, ByVal amount As Long) As LongPtr
' Performs a left (amount < 0) or a right (amount > 0) logical shift.
' The LShift/RShift codes are repeated to avoid extra stack frames.

    If amount < 0 And amount >= -SignBit Then
        If (bits And BitMask(SignBit + amount)) = 0 Then
            Shift = (bits And MaxMask(SignBit + amount)) * BitMask(-amount)
        Else
            Shift = (bits And MaxMask(SignBit + amount)) * BitMask(-amount) Or SignBitMask
        End If
    ElseIf amount > 0 And amount <= SignBit Then
        If bits < 0 Then
            Shift = (bits And Not SignBitMask) \ BitMask(amount) Or BitMask(SignBit - amount)
        Else
            Shift = bits \ BitMask(amount)
        End If
    ElseIf amount = 0 Then
        Shift = bits
    End If

End Function


'@Description("Rotates a bit pattern by an amount and returns the outcome.")
Public Function Rotate(ByVal bits As LongPtr, ByVal amount As Long) As LongPtr
' Circulates left or right for a negative or a positive amount respectively.
' The LShift/RShift codes are repeated to avoid extra stack frames.

    ' Normalize the rotation to the -SignBit to SignBit domain.
    amount = amount Mod (SignBit + 1)
    ' First set the amount for the right shift.
    If amount < 0 Then
        amount = SignBit + 1 + amount
    ElseIf amount = 0 Then
        Rotate = bits
        Exit Function
    End If

    ' Calculate the right shift.
    Dim RPart As LongPtr
    If bits < 0 Then
        RPart = (bits And Not SignBitMask) \ BitMask(amount) Or BitMask(SignBit - amount)
    Else
        RPart = bits \ BitMask(amount)
    End If

    ' Calculate the left shift.
    Dim LPart As LongPtr
    If (bits And BitMask(amount - 1)) = 0 Then
        LPart = (bits And MaxMask(amount - 1)) * BitMask(SignBit + 1 - amount)
    Else
        LPart = (bits And MaxMask(amount - 1)) * BitMask(SignBit + 1 - amount) Or SignBitMask
    End If

    ' Combine the left and right shifts to obtain the rotation.
    Rotate = LPart Or RPart

End Function


'@Description("Returns a random bit pattern.")
Public Function Random() As LongPtr
' The VBA.Rnd function is used to construct a pseudo-random LongPtr.

    ' Generate the first random integer.
    Const Max16 As Long = &H10000
    Random = Int(Max16 * Rnd)

#If Win64 Then
    ' Generate two more random integers and combine.
    Random = Max16 * (Max16 * Random Or Int(Max16 * Rnd)) Or Int(Max16 * Rnd)
#End If

    ' Generate the last random integer and combine to get the result.
    If (Random And BitMask16) = 0 Then
        Random = Max16 * Random Or Int(Max16 * Rnd)
    Else
        ' Mask the bit that ends up as signbit before multiplying.
        Random = Max16 * (Random And Not BitMask16) Or SignBitMask Or Int(Max16 * Rnd)
    End If

End Function


'@DefaultMember
'@Description("Returns or sets a selected bit in a bit pattern.")
'@Ignore ParameterCanBeByVal
Public Property Get Bit(ByRef bits As LongPtr, ByVal No As Long) As Boolean

    Bit = ((bits And BitMask(No)) <> 0)

End Property
Public Property Let Bit(ByRef bits As LongPtr, ByVal No As Long, ByVal RHS As Boolean)

    If RHS Then
        bits = bits Or BitMask(No)
    Else
        bits = bits And Not BitMask(No)
    End If

End Property


'@Description "Converts a bit pattern to or from an array of bits."
'@Ignore ParameterCanBeByVal
Public Property Get BitArray(ByRef bits As LongPtr) As Boolean()

    Dim BitArr(0 To SignBit) As Boolean
    Dim i As Long
    For i = 0 To SignBit
        BitArr(i) = bits And BitMask(i)
    Next

    BitArray = BitArr

End Property
Public Property Let BitArray(ByRef bits As LongPtr, ByRef RHS() As Boolean)
' A RHS array shorter than SignBit + 1 it is in effect padded with 0's.

    ' (Not Not RHS) = 0 is a standard trick.
    If (Not Not RHS) = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "Array is empty"
    If UBound(RHS) < LBound(RHS) Then Err.Raise vbErrorInvalidProcedureCall, , "Array has no elements"
    If UBound(RHS) - LBound(RHS) > SignBit Then Err.Raise vbErrorInvalidProcedureCall, , "Array is too long"

    bits = 0
    Dim Lbnd As Long: Lbnd = LBound(RHS)
    Dim i As Long
    For i = Lbnd To UBound(RHS)
        If RHS(i) Then
            bits = bits Or BitMask(i - Lbnd)
        End If
    Next

End Property


#If UseStringTable Then
'@Description "Converts a bit pattern to or from a bitstring."
'@Ignore ParameterCanBeByVal
Public Property Get BitString(ByRef bits As LongPtr) As String

#If Win64 Then
    Const NullStr As String * 64 = vbNullString
#Else
    Const NullStr As String * 32 = vbNullString
#End If
    BitString = NullStr

    ' Mask the SignBit and copy to a local variable.
    Dim Mask As LongPtr: Mask = bits And Not SignBitMask

    ' Combine the precalculated bitstrings for the individual bytes.
    Dim i As Long
    For i = SignBit + 1 - 7 To 1 Step -8
        Mid$(BitString, i) = StringTable(CLng(Mask And &HFF))
        Mask = Mask \ &H100
    Next

    ' Set the SignBit in the bitstring if it is set.
    If bits < 0 Then
        Mid$(BitString, 1) = "1"
    End If

End Property
#Else
'@Description "Converts a bit pattern to or from a bitstring."
'@Ignore ParameterCanBeByVal
Public Property Get BitString(ByRef bits As LongPtr) As String

#If Win64 Then
    Const ZeroStr As String * 64 = "00000000000000000000000000000000" & _
                                   "00000000000000000000000000000000"
#Else
    Const ZeroStr As String * 32 = "00000000000000000000000000000000"
#End If
    BitString = ZeroStr

    Dim i As Long
    For i = SignBit To 0 Step -1
        If bits And BitMask(i) Then
            Mid$(BitString, SignBit + 1 - i) = "1"
        End If
    Next

End Property
#End If
Public Property Let BitString(ByRef bits As LongPtr, ByVal RHS As String)
' A RHS string shorter than SignBit + 1 it is in effect left padded with 0's.

    If Len(RHS) = 0 Then Err.Raise vbErrorInvalidProcedureCall, , "String is empty"
    If Len(RHS) > SignBit + 1 Then Err.Raise vbErrorInvalidProcedureCall, , "String is too long"

    bits = 0
    Dim i As Long
    For i = 1 To Len(RHS)
        Select Case Mid$(RHS, i, 1)
        Case "1"
            bits = bits Or BitMask(Len(RHS) - i)
        Case Is <> "0"
            Err.Raise vbErrorInvalidProcedureCall, , "String is invalid"
        End Select
    Next

End Property


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Initializes the bit masks.")
Private Sub InitializeMasks()

    ' After the loop, "bits" will equal 2^(SignBit) - 1
    Dim bits As LongPtr
    Dim i As Long
    For i = 0 To SignBit - 1
        MaxMask(i) = bits
        BitMask(i) = bits + 1
        bits = 2 * bits + 1
    Next

    ' Now bits = 2^SignBit - 1 for sure (all lower bits set)
    MaxMask(SignBit) = bits
    BitMask(SignBit) = Not bits

    ' Set masks.
    SignBitMask = BitMask(SignBit)
    BitMask16 = BitMask(SignBit - 16)

End Sub


'@Description("Initializes the byte tables.")
Private Sub InitializeTables()

    ReverseTable(0) = 0
    CountTable(0) = 0
    Dim bits As Byte, Reverse As Byte, Count As Byte
    Dim n As Long, i As Long
    For n = 1 To 254
        If CountTable(n) = 0 Then
            bits = n
            Reverse = 0
            Count = 0
            For i = 0 To 7
                If (bits And 1) = 1 Then
                    Reverse = 2 * Reverse Or 1
                    Count = Count + 1
                Else
                    Reverse = 2 * Reverse
                End If
                bits = bits \ 2
            Next
            ReverseTable(n) = Reverse
            ReverseTable(Reverse) = n
            CountTable(n) = Count
            CountTable(Reverse) = Count
        End If
    Next
    ReverseTable(255) = 255
    CountTable(255) = 8

#If UseStringTable Then
    Const ZeroStr As String * 8 = "00000000"
    StringTable(0) = ZeroStr
    Dim BitStr As String * 8
    For n = 1 To 254
        If Len(StringTable(n)) = 0 Then
            bits = n
            BitStr = ZeroStr
            For i = 8 To 1 Step -1
                If (bits And 1) = 1 Then
                    Mid$(BitStr, i) = "1"
                End If
                bits = bits \ 2
            Next
            StringTable(n) = BitStr
            StringTable(ReverseTable(n)) = StrReverse$(BitStr)
        End If
    Next
    StringTable(255) = "11111111"
#End If

End Sub
