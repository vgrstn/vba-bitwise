Attribute VB_Name = "TestBitwise"
'@IgnoreModule VariableNotAssigned, UnassignedVariableUsage, LineLabelNotUsed, VariableNotUsed
'@TestModule
'@Folder("Tests")
'@ModuleDescription("Rubberduck unit tests for the Bitwise class: every result is checked against the bitstring reference in BitwiseTestHelpers.")
Option Explicit
Option Private Module

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Rubberduck unit tests for the Bitwise class
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.07
'
' Each test runs over the sample patterns (edge cases and random patterns) and
' every bit position or amount, and reports the number of failures and the first
' one. The timings are in BenchmarkBitwise, not here. Every call is qualified
' with Bitwise, the predeclared instance; unqualified names would call BitwiseLib.

#Const LateBinding = False

#If LateBinding Then
Private Assert As Object
#Else
Private Assert As AssertClass
#End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Test module set-up
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@ModuleInitialize
Private Sub ModuleInitialize()
#If LateBinding Then
    Set Assert = CreateObject("Rubberduck.AssertClass")
#Else
    Set Assert = New AssertClass
#End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Helpers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Returns the error description BitString Let raises for a string (empty when none).")
Private Function StringLetError(ByVal bitText As String) As String

    Dim Target As LongPtr
    On Error Resume Next
    Bitwise.BitString(Target) = bitText
    StringLetError = Err.Description
    On Error GoTo 0

End Function


'@Description("Returns the error description BitArray Let raises for an array (empty when none).")
Private Function ArrayLetError(ByRef bitValues() As Boolean) As String

    ' VBA crashes the host when a ByRef array parameter is passed on as the
    ' value of an array Property Let, so the Let gets a local copy.
    Dim Values() As Boolean
    Dim Target As LongPtr
    On Error Resume Next
    Values = bitValues
    Bitwise.BitArray(Target) = Values
    ArrayLetError = Err.Description
    On Error GoTo 0

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: single bits
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Bits")
Private Sub BitGet_AllPositions_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Expected As String: Expected = ReferenceString(Sample)
        Dim Pattern As LongPtr: Pattern = Sample
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            NoteResult IIf(Bitwise.Bit(Pattern, Position), "1", "0"), _
                       RefGetBit(Expected, Position), _
                       "Bit(" & Expected & ", " & Position & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Bits")
Private Sub BitLet_AllPositions_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            Dim Pattern As LongPtr
            Pattern = Sample
            Bitwise.Bit(Pattern, Position) = True
            NoteResult ReferenceString(Pattern), _
                       RefPutBit(Original, Position, "1"), _
                       "Bit(" & Original & ", " & Position & ") = True", _
                       Failures, _
                       FirstFailure
            Pattern = Sample
            Bitwise.Bit(Pattern, Position) = False
            NoteResult ReferenceString(Pattern), _
                       RefPutBit(Original, Position, "0"), _
                       "Bit(" & Original & ", " & Position & ") = False", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Bits")
Private Sub Bit_DefaultMember_GetsAndSets()
    On Error GoTo TestFail

    ' Bit is the default member: Bitwise(x, No) is Bitwise.Bit(x, No).
    Dim Pattern As LongPtr
    Bitwise.Bit(Pattern, 2) = True

    Assert.IsTrue Pattern = 4, "Bitwise(x, 2) = True gives " & Pattern
    Assert.IsTrue Bitwise.Bit(Pattern, 2), "Bitwise(4, 2)"
    Assert.IsFalse Bitwise.Bit(Pattern, 1), "Bitwise(4, 1)"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Bits")
Private Sub Clear_AllPositions_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            NoteResult ReferenceString(Bitwise.Clear(Sample, Position)), _
                       RefPutBit(Original, Position, "0"), _
                       "Clear(" & Original & ", " & Position & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Bits")
Private Sub Flip_AllPositions_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            NoteResult ReferenceString(Bitwise.Flip(Sample, Position)), _
                       RefFlipBit(Original, Position), _
                       "Flip(" & Original & ", " & Position & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Bits")
Private Sub Swap_AllPairs_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            Dim Other As Long
            For Other = 0 To PatternSize - 1
                NoteResult ReferenceString(Bitwise.Swap(Sample, Position, Other)), _
                           RefSwapBits(Original, Position, Other), _
                           "Swap(" & Original & ", " & Position & ", " & Other & ")", _
                           Failures, _
                           FirstFailure
            Next
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: whole patterns
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Patterns")
Private Sub Count_Samples_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        NoteResult CStr(Bitwise.Count(Sample)), _
                   CStr(RefCountBits(Original)), _
                   "Count(" & Original & ")", _
                   Failures, _
                   FirstFailure
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Patterns")
Private Sub Invert_Samples_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        NoteResult ReferenceString(Bitwise.Invert(Sample)), _
                   RefInvertBits(Original), _
                   "Invert(" & Original & ")", _
                   Failures, _
                   FirstFailure
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Patterns")
Private Sub Reverse_Samples_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        NoteResult ReferenceString(Bitwise.Reverse(Sample)), _
                   StrReverse$(Original), _
                   "Reverse(" & Original & ")", _
                   Failures, _
                   FirstFailure
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Patterns")
Private Sub LClear_AllPositions_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            NoteResult ReferenceString(Bitwise.LClear(Sample, Position)), _
                       RefLClearBits(Original, Position), _
                       "LClear(" & Original & ", " & Position & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Patterns")
Private Sub RClear_AllPositions_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            NoteResult ReferenceString(Bitwise.RClear(Sample, Position)), _
                       RefRClearBits(Original, Position), _
                       "RClear(" & Original & ", " & Position & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: shifts and rotations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Shifts")
Private Sub LShift_AllAmounts_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        ' Beyond the pattern size every bit is shifted out.
        Dim Places As Long
        For Places = 0 To PatternSize + 1
            NoteResult ReferenceString(Bitwise.LShift(Sample, Places)), _
                       RefShiftBits(Original, -Places), _
                       "LShift(" & Original & ", " & Places & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Shifts")
Private Sub RShift_AllAmounts_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Places As Long
        For Places = 0 To PatternSize + 1
            NoteResult ReferenceString(Bitwise.RShift(Sample, Places)), _
                       RefShiftBits(Original, Places), _
                       "RShift(" & Original & ", " & Places & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Shifts")
Private Sub Shift_AllAmounts_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Places As Long
        For Places = -PatternSize - 1 To PatternSize + 1
            NoteResult ReferenceString(Bitwise.Shift(Sample, Places)), _
                       RefShiftBits(Original, Places), _
                       "Shift(" & Original & ", " & Places & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Shifts")
Private Sub Rotate_AllAmounts_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        ' Amounts beyond the pattern size wrap around.
        Dim Places As Long
        For Places = -2 * PatternSize To 2 * PatternSize
            NoteResult ReferenceString(Bitwise.Rotate(Sample, Places)), _
                       RefRotateBits(Original, Places), _
                       "Rotate(" & Original & ", " & Places & ")", _
                       Failures, _
                       FirstFailure
        Next
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: random patterns
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("Random")
Private Sub Random_ManyDraws_AreBalanced()
    On Error GoTo TestFail

    ' About half of all bits are set, and every bit is set and cleared at least
    ' once. With 2000 draws a false failure is practically impossible.
    Const Draws As Long = 2000
    Dim SetCount(0 To PatternSize - 1) As Long
    Dim Ones As Long
    Dim Draw As Long
    For Draw = 1 To Draws
        Dim Drawn As String: Drawn = ReferenceString(Bitwise.Random)
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            If RefGetBit(Drawn, Position) = "1" Then
                SetCount(Position) = SetCount(Position) + 1
                Ones = Ones + 1
            End If
        Next
    Next

    Dim Half As Double: Half = Draws * PatternSize / 2
    Assert.IsTrue Abs(Ones - Half) < 0.05 * Half, Ones & " of " & 2 * Half & " bits set"
    For Position = 0 To PatternSize - 1
        Assert.IsTrue SetCount(Position) > 0 And SetCount(Position) < Draws, "bit " & Position & " set " & SetCount(Position) & " times"
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: BitString
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("BitString")
Private Sub BitStringGet_Samples_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Pattern As LongPtr: Pattern = Sample
        NoteResult Bitwise.BitString(Pattern), _
                   ReferenceString(Pattern), _
                   "BitString(" & Hex$(Pattern) & ")", _
                   Failures, _
                   FirstFailure
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("BitString")
Private Sub BitStringLet_Samples_RoundTrips()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Pattern As LongPtr
        Bitwise.BitString(Pattern) = ReferenceString(Sample)
        NoteResult Hex$(Pattern), _
                   Hex$(Sample), _
                   "BitString(x) = " & ReferenceString(Sample), _
                   Failures, _
                   FirstFailure
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("BitString")
Private Sub BitStringLet_ShortString_IsLeftPadded()
    On Error GoTo TestFail

    Dim Pattern As LongPtr
    Bitwise.BitString(Pattern) = "10101"

    Assert.IsTrue Pattern = 21, "BitString(x) = ""10101"" gives " & Pattern

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("BitString")
Private Sub BitStringLet_BadStrings_Raise()
    On Error GoTo TestFail

    Assert.AreEqual "String is empty", StringLetError(vbNullString)
    Assert.AreEqual "String is too long", StringLetError(String$(PatternSize + 1, "0"))
    Assert.AreEqual "String is invalid", StringLetError("10201")
    Assert.AreEqual vbNullString, StringLetError(String$(PatternSize, "1"))

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Tests: BitArray
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@TestMethod("BitArray")
Private Sub BitArrayGet_Samples_MatchesReference()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Pattern As LongPtr: Pattern = Sample
        Dim Values() As Boolean: Values = Bitwise.BitArray(Pattern)
        Dim Actual As String: Actual = vbNullString
        Dim Position As Long
        For Position = UBound(Values) To LBound(Values) Step -1
            Actual = Actual & IIf(Values(Position), "1", "0")
        Next
        NoteResult Actual, _
                   ReferenceString(Pattern), _
                   "BitArray(" & Hex$(Pattern) & ")", _
                   Failures, _
                   FirstFailure
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("BitArray")
Private Sub BitArrayLet_Samples_RoundTrips()
    On Error GoTo TestFail

    Dim Failures As Long
    Dim FirstFailure As String
    Dim Sample As Variant
    For Each Sample In SamplePatterns()
        Dim Original As String: Original = ReferenceString(Sample)
        Dim Values() As Boolean
        ReDim Values(0 To PatternSize - 1)
        Dim Position As Long
        For Position = 0 To PatternSize - 1
            Values(Position) = (RefGetBit(Original, Position) = "1")
        Next
        Dim Pattern As LongPtr
        Bitwise.BitArray(Pattern) = Values
        NoteResult ReferenceString(Pattern), _
                   Original, _
                   "BitArray(x) = " & Original, _
                   Failures, _
                   FirstFailure
    Next

    Assert.IsTrue Failures = 0, Failures & " failures, first " & FirstFailure

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("BitArray")
Private Sub BitArrayLet_ShortArray_IsPadded()
    On Error GoTo TestFail

    ' The lower bound does not matter: the first element is bit 0.
    Dim Values() As Boolean
    ReDim Values(1 To 3)
    Values(1) = True
    Values(3) = True
    Dim Pattern As LongPtr
    Bitwise.BitArray(Pattern) = Values

    Assert.IsTrue Pattern = 5, "BitArray(x) = (True, False, True) gives " & Pattern

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("BitArray")
Private Sub BitArrayLet_BadArrays_Raise()
    On Error GoTo TestFail

    Dim TooLong() As Boolean
    ReDim TooLong(0 To PatternSize)
    Assert.AreEqual "Array is too long", ArrayLetError(TooLong)

    Dim Undimensioned() As Boolean
    Assert.AreEqual "Array is empty", ArrayLetError(Undimensioned)

    Dim FullLength() As Boolean
    ReDim FullLength(1 To PatternSize)
    Assert.AreEqual vbNullString, ArrayLetError(FullLength)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
