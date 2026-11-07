Attribute VB_Name = "BitwiseLib"
'@IgnoreModule AssignedByValParameter, VariableNotAssigned, UnassignedVariableUsage, ProcedureNotUsed, UseMeaningfulName, MultipleDeclarations
'@Folder("Module")
'@ModuleDescription("Bitwise Operations Module.")

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Bitwise Operations Module
' Enumeration: For Each Bit in BitArray(bits)
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.07
'
' The bits argument is used as variable name for a LongPtr bit pattern. A bit is
' a Boolean and an array of bits is a Boolean array. A bitstring is obviously a
' string consisting of 0's and 1's.
'
' All Methods are Functions returning a LongPtr except for the Count Function
' which returns a Long and the GetBit Function returning a Boolean.
'
' Methods
' CountBits(bits)           Returns the number of set bits in a bit pattern
' GetBit(bits, No)          Returns a bit from a bit pattern
' SetBit(bits, No)          Sets a selected bit in a bit pattern
' FlipBit(bits, No)         Flips a selected bit in a bit pattern
' SwapBits(bits, No1,No2)   Swaps two selected bits in a bit pattern
' InvertBits(bits)          Inverts a bit pattern
' ReverseBits(bits)         Reverses a bit pattern
' ClearBit(bits, No)        Clears a selected bit in a bit pattern
' LClearBits(bits, No)      Clears a bits pattern left from a selected bit
' RClearBits(bits, No)      Clears a bits pattern right from a selected bit
' LShiftBits(bits, amount)  Shifts a bit pattern to the left by an amount
' RShiftBits(bits, amount)  Shifts a bit pattern to the right by an amount
' ShiftBits(bits, amount)   Shifts a bit pattern by an amount
' RotateBits(bits, amount)  Rotates a bit pattern by an amount
' RandomBits()              Creates a random bit pattern
'
' Properties
' BitArray(bits)            Converts a bit pattern to or from an array of bits
' BitString(bits)           Converts a bit pattern to or from a bitstring
'
' A multiplication by n where n = 2 ^ i is equivalent to a left shift by i bit
' positions. For a left shift overflow by multiplication needs to be prevented.
' Likewise, a division by n where n = 2 ^ i is equivalent to a right shift by i
' positions. VB uses two's complements representation for signed integers. Hence,
' for negative integers, division doesn't correspond to a right shift. In the
' latter case the signbit needs to be masked before the division and subsequently
' set depending on whether the bit that is shifted to the signbit is set or not.
'
' Precalculated bitmasks are used for efficiency and initialized on first use.
' The Count, Reverse and BitString Function are relatively slow because these
' routines require bit by bit processing. However, the use of precalculated
' byte lookup tables makes these functions 3-4 times faster.
'
' Time per call in µs (RunBenchmarks in BenchmarkBitwise, 6,400,000 calls per
' method, loop overhead subtracted):
'
' Method       Module     Class
' Count        0.20       0.22      (byte lookup table)
' Flip         0.04       0.07
' Invert       0.02       0.06
' Reverse      0.24       0.25      (byte lookup table)
' Shift        0.08       0.10
' Rotate       0.11       0.13
' Random       0.19       0.21
' Swap         0.07       0.09
' BitString    0.30       0.34      (string lookup table)
'
' A class call costs about 0.02-0.03 µs more than a module call. For the fastest
' methods (Flip, Invert, Swap, Shift) that is 35-150%, for the others 5-20%.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Compiler directives
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The allocated memory can be minimized by switching off the StringTable lookup
' table which uses more memory than all the other tables and masks together.

#Const UseStringTable = True

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' VB uses two's complement representation for numbers.
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
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Returns the number of set bits in a bit pattern.")
Public Function CountBits(ByVal bits As LongPtr) As Long

    If CountTable(&HFF) = 0 Then InitializeTables
    
    ' Count and mask the SignBit if it is set.
    If bits < 0 Then
        CountBits = 1
        bits = bits And Not SignBitMask
    End If
    
    ' Sum the precalculated byte counts.
    Dim i As Long
    For i = 0 To SignBit \ 8
        CountBits = CountBits + CountTable(CLng(bits And &HFF))
        bits = bits \ &H100
    Next
    
End Function


'@Description("Returns a selected bit from a bit pattern.")
Public Function GetBit(ByVal bits As LongPtr, ByVal No As Long) As Boolean

    If BitMask(0) = 0 Then InitializeMasks
    
    GetBit = ((bits And BitMask(No)) <> 0)

End Function


'@Description("Sets a selected bit in a bit pattern and returns the outcome.")
Public Function SetBit(ByVal bits As LongPtr, ByVal No As Long) As LongPtr
    
    If BitMask(0) = 0 Then InitializeMasks
    
    SetBit = bits Or BitMask(No)

End Function


'@Description("Flips a selected bit in a bit pattern and returns the outcome.")
Public Function FlipBit(ByVal bits As LongPtr, ByVal No As Long) As LongPtr

    If BitMask(0) = 0 Then InitializeMasks
    
    FlipBit = bits Xor BitMask(No)

End Function


'@Description("Swaps two selected bits in a bit pattern and returns the outcome.")
Public Function SwapBits(ByVal bits As LongPtr, ByVal No1 As Long, ByVal No2 As Long) As LongPtr

    If BitMask(0) = 0 Then InitializeMasks
    
    ' Save the mask for the two bits to be swapped.
    Dim Mask As LongPtr: Mask = BitMask(No1) Or BitMask(No2)
   
   ' Swap the two bits if they are different.
    Select Case bits And Mask
    Case 0, Mask
        SwapBits = bits
    Case Else
        SwapBits = bits Xor Mask
    End Select
    
End Function


'@Description("Inverts a bit pattern and returns the outcome.")
Public Function InvertBits(ByVal bits As LongPtr) As LongPtr

    InvertBits = Not bits
    
End Function


'@Description("Reverses a bit pattern and returns the outcome.")
Public Function ReverseBits(ByVal bits As LongPtr) As LongPtr

    If ReverseTable(&HFF) = 0 Then InitializeTables
    
    ' Save the bit pattern for the 0th bit and the SignBit.
    Dim Mask As LongPtr: Mask = bits And (1 Or SignBitMask)
    
    ' Mask the 0th bit and the SignBit.
    bits = bits And Not (1 Or SignBitMask)
    
    ' Combine the precalculated reversed bytes.
    Dim i As Long
    For i = 0 To SignBit \ 8
        ReverseBits = &H100 * ReverseBits Or ReverseTable(CLng(bits And &HFF))
        bits = bits \ &H100
    Next
    
    ' Restore the two masked bits and swap them if they are different.
    Select Case Mask
    Case 1, SignBitMask
        ReverseBits = ReverseBits Or (Mask Xor (1 Or SignBitMask))
    Case Else
        ReverseBits = ReverseBits Or Mask
    End Select
    
End Function


'@Description("Clears a selected bit in a bit pattern and returns the outcome.")
Public Function ClearBit(ByVal bits As LongPtr, ByVal No As Long) As LongPtr
    
    If BitMask(0) = 0 Then InitializeMasks
        
    ClearBit = bits And Not BitMask(No)

End Function


'@Description("Clears a bit pattern left from a selected bit and returns the outcome.")
Public Function LClearBits(ByVal bits As LongPtr, ByVal No As Long) As LongPtr

    If BitMask(0) = 0 Then InitializeMasks
    
    LClearBits = bits And (BitMask(No) Or MaxMask(No))
    
End Function


'@Description("Clears a bit pattern right from a selected bit and returns the outcome.")
Public Function RClearBits(ByVal bits As LongPtr, ByVal No As Long) As LongPtr

    If BitMask(0) = 0 Then InitializeMasks
    
    RClearBits = bits And Not MaxMask(No)
    
End Function


'@Description("Shifts a bit pattern to the left by an amount and returns the outcome.")
Public Function LShiftBits(ByVal bits As LongPtr, ByVal amount As Byte) As LongPtr
    
    If amount > SignBit Then Exit Function
    
    If BitMask(0) = 0 Then InitializeMasks
    
    If amount = 0 Then
        LShiftBits = bits
    ElseIf (bits And BitMask(SignBit - amount)) = 0 Then
        ' Clear the overflow bits before multiplying.
        LShiftBits = (bits And MaxMask(SignBit - amount)) * BitMask(amount)
    Else
        ' Clear the overflow bits before multiplying and set the signbit.
        LShiftBits = (bits And MaxMask(SignBit - amount)) * BitMask(amount) Or SignBitMask
    End If
    
End Function


'@Description("Shifts a bit pattern to the right by an amount and returns the outcome.")
Public Function RShiftBits(ByVal bits As LongPtr, ByVal amount As Byte) As LongPtr
' RShiftBits is effectively a right logical shift, it fills the vacated bits with 0.
    
    If amount > SignBit Then Exit Function
        
    If BitMask(0) = 0 Then InitializeMasks
        
    If amount = 0 Then
        RShiftBits = bits
    ElseIf bits < 0 Then
        ' Clear the signbit before dividing and set the shifted signbit.
        RShiftBits = (bits And Not SignBitMask) \ BitMask(amount) Or BitMask(SignBit - amount)
    Else
        RShiftBits = bits \ BitMask(amount)
    End If
 
End Function


'@Description("Shifts a bit pattern by an amount and returns the outcome.")
Public Function ShiftBits(ByVal bits As LongPtr, ByVal amount As Long) As LongPtr
' Performs a left (amount < 0) or a right (amount > 0) logical shift.
' The LShift/RShift codes are repeated to avoid extra stack frames.
    
    If BitMask(0) = 0 Then InitializeMasks
        
    If amount < 0 And amount >= -SignBit Then
        If (bits And BitMask(SignBit + amount)) = 0 Then
            ShiftBits = (bits And MaxMask(SignBit + amount)) * BitMask(-amount)
        Else
            ShiftBits = (bits And MaxMask(SignBit + amount)) * BitMask(-amount) Or SignBitMask
        End If
    ElseIf amount > 0 And amount <= SignBit Then
        If bits < 0 Then
            ShiftBits = (bits And Not SignBitMask) \ BitMask(amount) Or BitMask(SignBit - amount)
        Else
            ShiftBits = bits \ BitMask(amount)
        End If
    ElseIf amount = 0 Then
        ShiftBits = bits
    End If
    
End Function


'@Description("Rotates a bit pattern by an amount and returns the outcome.")
Public Function RotateBits(ByVal bits As LongPtr, ByVal amount As Long) As LongPtr
' Performs a left or right rotation for a negative or a positive amount respectively.
' The LShift/RShift codes are repeated to avoid extra stack frames.

    If BitMask(0) = 0 Then InitializeMasks
        
    ' Normalize the rotation to the -SignBit to SignBit domain.
    amount = amount Mod (SignBit + 1)
    
    ' First set the amount for the right shift.
    If amount < 0 Then
        amount = SignBit + 1 + amount
    ElseIf amount = 0 Then
        RotateBits = bits
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
    RotateBits = LPart Or RPart
        
End Function


'@Description("Returns a random bit pattern.")
Public Function RandomBits() As LongPtr
' The VBA.Rnd function is used to construct a pseudo-random LongPtr.
    
    If BitMask(0) = 0 Then InitializeMasks: Randomize
    
    ' Generate the first random integer.
    Const Max16 As Long = &H10000
    RandomBits = Int(Max16 * Rnd)

#If Win64 Then
    ' Generate two more random integers and combine.
    RandomBits = Max16 * (Max16 * RandomBits Or Int(Max16 * Rnd)) Or Int(Max16 * Rnd)
#End If
    
    ' Generate the last random integer and combine to get the result.
    If (RandomBits And BitMask16) = 0 Then
        RandomBits = Max16 * RandomBits Or Int(Max16 * Rnd)
    Else
        ' Mask the bit that ends up as signbit before multiplying.
        RandomBits = Max16 * (RandomBits And Not BitMask16) Or SignBitMask Or Int(Max16 * Rnd)
    End If
    
End Function


'@Description("Converts a bit pattern to or from an array of bits.")
'@Ignore ParameterCanBeByVal
Public Property Get BitArray(ByRef bits As LongPtr) As Boolean()

    If BitMask(0) = 0 Then InitializeMasks
            
    Dim BitArr(0 To SignBit) As Boolean
    Dim i As Long
    For i = 0 To SignBit
        BitArr(i) = bits And BitMask(i)
    Next
    
    BitArray = BitArr

End Property
Public Property Let BitArray(ByRef bits As LongPtr, ByRef RHS() As Boolean)
' An RHS array shorter than SignBit + 1 is in effect padded with 0's.
'
' VBA crashes the host when an array that the caller received ByRef is passed
' on as the RHS of an array Property Let. Copy it to a local array first:
'     Dim Values() As Boolean: Values = bitValues
'     BitArray(bits) = Values

    On Error GoTo Fail
    Dim Lo As Long: Lo = LBound(RHS)
    Dim Up As Long: Up = UBound(RHS)
    On Error GoTo 0
    
    If Up < Lo Then Err.Raise vbErrorInvalidProcedureCall, , "Array has no elements"
    If Up - Lo > SignBit Then Err.Raise vbErrorInvalidProcedureCall, , "Array is too long"

    If BitMask(0) = 0 Then InitializeMasks
    
    bits = 0
    Dim i As Long
    For i = Lo To Up
        If RHS(i) Then
            bits = bits Or BitMask(i - Lo)
        End If
    Next
    Exit Property

Fail:
    Err.Raise vbErrorInvalidProcedureCall, , "Array is empty"
    
End Property


#If UseStringTable Then
'@Description("Converts a bit pattern to or from a bitstring.")
'@Ignore ParameterCanBeByVal
Public Property Get BitString(ByRef bits As LongPtr) As String
    
#If Win64 Then
    Const NullStr As String * 64 = vbNullString
#Else
    Const NullStr As String * 32 = vbNullString
#End If
    BitString = NullStr
 
    If Len(StringTable(0)) = 0 Then InitializeTables
    
    ' Mask the SignBit and copy to a local variable (ByRef bits!)
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
'@Description("Converts a bit pattern to or from a bitstring.")
'@Ignore ParameterCanBeByVal
Public Property Get BitString(ByRef bits As LongPtr) As String

    If BitMask(0) = 0 Then InitializeMasks
    
#If Win64 Then
    Const ZeroStr As String * 64 = "0000000000000000000000000000000000000000000000000000000000000000"
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
    
    If BitMask(0) = 0 Then InitializeMasks

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
    
    If BitMask(0) = 0 Then InitializeMasks
        
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
