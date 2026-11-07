Attribute VB_Name = "BitwiseTestHelpers"
'@IgnoreModule UseMeaningfulName
'@Folder("Tests")
'@ModuleDescription("Reference bit operations on bitstrings for the unit tests of BitwiseLib and Bitwise.")
Option Explicit
Option Private Module

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Reference operations for the Bitwise unit tests
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.07
'
' The tests compare every result with a reference that works on bitstrings. The
' reference uses only Hex$ and string functions, so it does not depend on the
' code under test. Bit 0 is the last character of a bitstring, the sign bit the
' first.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The number of bits in a LongPtr and the pattern with only the sign bit set.
#If Win64 Then
Public Const PatternSize As Long = 64
Private Const SignPattern As LongLong = &H8000000000000000^
#Else
Public Const PatternSize As Long = 32
Private Const SignPattern As Long = &H80000000
#End If

' The number of random patterns added to the fixed samples.
Private Const RandomSamples As Long = 40


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Returns the test patterns: edge cases and random patterns.")
Public Function SamplePatterns() As Variant

    ' 0101...01 over the full width, and its invert 1010...10.
    Dim Alternating As LongPtr
    Dim Position As Long
    For Position = 1 To PatternSize \ 2
        Alternating = 4 * Alternating + 1
    Next

    Dim Fixed As Variant
    Fixed = Array(0, 1, 2, 3, -1, -2, 255, 256, 65535, 123456, -123456, _
                  SignPattern, SignPattern + 1, Not SignPattern, Alternating, Not Alternating)

    Dim Samples() As LongPtr
    ReDim Samples(0 To UBound(Fixed) + RandomSamples)
    Dim Index As Long
    For Index = 0 To UBound(Fixed)
        Samples(Index) = Fixed(Index)
    Next
    For Index = UBound(Fixed) + 1 To UBound(Samples)
        Samples(Index) = RandomBits
    Next

    SamplePatterns = Samples

End Function


'@Description("Returns the bitstring of a bit pattern, made with Hex$ only.")
Public Function ReferenceString(ByVal bitPattern As LongPtr) As String

    Const HexDigits As String = "0123456789ABCDEF"
    Const Nibbles As String = "0000000100100011010001010110011110001001101010111100110111101111"

    Dim HexText As String
    HexText = Right$(String$(PatternSize \ 4, "0") & Hex$(bitPattern), PatternSize \ 4)

    Dim Position As Long
    For Position = 1 To Len(HexText)
        Dim Digit As Long
        Digit = InStr(HexDigits, Mid$(HexText, Position, 1)) - 1
        ReferenceString = ReferenceString & Mid$(Nibbles, 4 * Digit + 1, 4)
    Next

End Function


'@Description("Returns the character (0 or 1) of a bit in a bitstring.")
Public Function RefGetBit(ByVal bitText As String, ByVal bitNo As Long) As String

    RefGetBit = Mid$(bitText, PatternSize - bitNo, 1)

End Function


'@Description("Returns a bitstring with a bit replaced by a character (0 or 1).")
Public Function RefPutBit( _
    ByVal bitText As String, _
    ByVal bitNo As Long, _
    ByVal bitValue As String _
) As String

    Mid$(bitText, PatternSize - bitNo, 1) = bitValue
    RefPutBit = bitText

End Function


'@Description("Returns a bitstring with a bit flipped.")
Public Function RefFlipBit(ByVal bitText As String, ByVal bitNo As Long) As String

    RefFlipBit = RefPutBit(bitText, bitNo, IIf(RefGetBit(bitText, bitNo) = "1", "0", "1"))

End Function


'@Description("Returns a bitstring with two bits swapped.")
Public Function RefSwapBits( _
    ByVal bitText As String, _
    ByVal bitNo1 As Long, _
    ByVal bitNo2 As Long _
) As String

    Dim First As String: First = RefGetBit(bitText, bitNo1)
    RefSwapBits = RefPutBit(RefPutBit(bitText, bitNo1, RefGetBit(bitText, bitNo2)), bitNo2, First)

End Function


'@Description("Returns the number of 1's in a bitstring.")
Public Function RefCountBits(ByVal bitText As String) As Long

    RefCountBits = Len(bitText) - Len(Replace(bitText, "1", vbNullString))

End Function


'@Description("Returns a bitstring with every bit inverted.")
Public Function RefInvertBits(ByVal bitText As String) As String

    RefInvertBits = Replace(Replace(Replace(bitText, "0", "x"), "1", "0"), "x", "1")

End Function


'@Description("Returns a bitstring with the bits left from a bit cleared.")
Public Function RefLClearBits(ByVal bitText As String, ByVal bitNo As Long) As String

    RefLClearBits = String$(PatternSize - bitNo - 1, "0") & Right$(bitText, bitNo + 1)

End Function


'@Description("Returns a bitstring with the bits right from a bit cleared.")
Public Function RefRClearBits(ByVal bitText As String, ByVal bitNo As Long) As String

    RefRClearBits = Left$(bitText, PatternSize - bitNo) & String$(bitNo, "0")

End Function


'@Description("Returns a bitstring shifted left (shiftAmount < 0) or right (shiftAmount > 0), filled with 0's.")
Public Function RefShiftBits(ByVal bitText As String, ByVal shiftAmount As Long) As String

    If Abs(shiftAmount) >= PatternSize Then
        RefShiftBits = String$(PatternSize, "0")
    ElseIf shiftAmount < 0 Then
        RefShiftBits = Mid$(bitText, 1 - shiftAmount) & String$(-shiftAmount, "0")
    Else
        RefShiftBits = String$(shiftAmount, "0") & Left$(bitText, PatternSize - shiftAmount)
    End If

End Function


'@Description("Returns a bitstring rotated left (shiftAmount < 0) or right (shiftAmount > 0).")
Public Function RefRotateBits(ByVal bitText As String, ByVal shiftAmount As Long) As String

    ' A left rotation equals a right rotation by the complement.
    Dim RightAmount As Long: RightAmount = shiftAmount Mod PatternSize
    If RightAmount < 0 Then RightAmount = RightAmount + PatternSize

    RefRotateBits = Right$(bitText, RightAmount) & Left$(bitText, PatternSize - RightAmount)

End Function


'@Description("Adds a failure to a count and keeps the text of the first one.")
Public Sub NoteResult( _
    ByVal actualText As String, _
    ByVal expectedText As String, _
    ByVal caseText As String, _
    ByRef failureCount As Long, _
    ByRef failureText As String _
)

    If actualText <> expectedText Then
        failureCount = failureCount + 1
        If failureCount = 1 Then
            failureText = caseText & ": expected " & expectedText & ", actual " & actualText
        End If
    End If

End Sub
