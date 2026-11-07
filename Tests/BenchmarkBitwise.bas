Attribute VB_Name = "BenchmarkBitwise"
'@IgnoreModule FunctionReturnValueDiscarded, VariableNotUsed, AssignmentNotUsed
'@Folder("Tests")
'@ModuleDescription("Benchmarks for BitwiseLib and the Bitwise class; the results go to the Immediate window.")
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Benchmarks for BitwiseLib and the Bitwise class
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2026.11.07
'
' Run from the Immediate window:
'     RunBenchmarks             100,000 patterns x 64 calls per method
'     RunBenchmarks 10000       a quick run
'
' The time is in ms per call, after subtracting the time of the same loop with
' only an array read, as in the method table in the header of BitwiseLib. The
' benchmarks are not Rubberduck tests, so the Test Explorer stays fast.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#If Win64 Then
Private Const SignBit As Long = 63
#Else
Private Const SignBit As Long = 31
#End If

Private Const DefaultPatternCount As Long = 100000

' The timed methods; mtBaseline is the loop with only an array read.
Private Enum MethodTimed
    mtBaseline
    mtCount
    mtFlip
    mtInvert
    mtReverse
    mtShift
    mtRotate
    mtRandom
    mtSwap
    mtBitString
    mtLast = mtBitString
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Times every method of BitwiseLib and the Bitwise class and prints ms per call.")
'@Ignore ProcedureNotUsed
Public Sub RunBenchmarks(Optional ByVal patternCount As Long = DefaultPatternCount)

    Dim Inputs() As LongPtr
    ReDim Inputs(1 To patternCount)
    Dim Index As Long
    For Index = 1 To patternCount
        Inputs(Index) = RandomBits
    Next

    Dim Calls As Double: Calls = CDbl(patternCount) * (SignBit + 1)
    Dim ModuleBase As Double: ModuleBase = ModuleTime(mtBaseline, Inputs)
    Dim ClassBase As Double: ClassBase = ClassTime(mtBaseline, Inputs)

    Debug.Print
    Debug.Print "Bitwise benchmark, " & Format$(Calls, "#,##0") & " calls per method"
    Debug.Print "Method", "Module (µs)", "Class (µs)"
    Dim Method As MethodTimed
    For Method = mtCount To mtLast
        Dim ModuleCall As Double
        ModuleCall = (ModuleTime(Method, Inputs) - ModuleBase) * 1000000 / Calls
        Dim ClassCall As Double
        ClassCall = (ClassTime(Method, Inputs) - ClassBase) * 1000000 / Calls
        Debug.Print MethodName(Method), Format$(ModuleCall, "0.000"), Format$(ClassCall, "0.000")
    Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description("Returns the name of a timed method.")
Private Function MethodName(ByVal methodNo As MethodTimed) As String

    MethodName = Split("Baseline Count Flip Invert" & _
                       " Reverse Shift Rotate Random Swap BitString")(methodNo)

End Function


'@Description("Returns the seconds a BitwiseLib method takes for every pattern and bit position.")
Private Function ModuleTime(ByVal methodNo As MethodTimed, ByRef patterns() As LongPtr) As Double
' One loop per method: a Select Case inside the loop would be timed too.

    Dim Result As LongPtr
    Dim Counted As Long
    Dim Converted As String
    Dim Index As Long
    Dim Position As Long

    Stopwatch.Reset
    Stopwatch.Start
    Select Case methodNo
    Case mtBaseline
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = patterns(Index)
            Next
        Next
    Case mtCount
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Counted = CountBits(patterns(Index))
            Next
        Next
    Case mtFlip
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = FlipBit(patterns(Index), Position)
            Next
        Next
    Case mtInvert
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = BitwiseLib.InvertBits(patterns(Index))
            Next
        Next
    Case mtReverse
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = ReverseBits(patterns(Index))
            Next
        Next
    Case mtShift
        ' Amounts from -SignBit to SignBit in steps of 2.
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = ShiftBits(patterns(Index), 2 * Position - SignBit)
            Next
        Next
    Case mtRotate
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = RotateBits(patterns(Index), 2 * Position - SignBit)
            Next
        Next
    Case mtRandom
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = RandomBits
            Next
        Next
    Case mtSwap
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = SwapBits(patterns(Index), Position, SignBit - Position)
            Next
        Next
    Case mtBitString
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Converted = BitString(patterns(Index))
            Next
        Next
    End Select
    ModuleTime = Stopwatch.Halt

End Function


'@Description("Returns the seconds a Bitwise class method takes for every pattern and bit position.")
Private Function ClassTime(ByVal methodNo As MethodTimed, ByRef patterns() As LongPtr) As Double
' The same loops as ModuleTime, with the methods of the predeclared Bitwise instance.

    Dim Result As LongPtr
    Dim Counted As Long
    Dim Converted As String
    Dim Index As Long
    Dim Position As Long

    Stopwatch.Reset
    Stopwatch.Start
    Select Case methodNo
    Case mtBaseline
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = patterns(Index)
            Next
        Next
    Case mtCount
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Counted = Bitwise.Count(patterns(Index))
            Next
        Next
    Case mtFlip
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = Bitwise.Flip(patterns(Index), Position)
            Next
        Next
    Case mtInvert
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = Bitwise.Invert(patterns(Index))
            Next
        Next
    Case mtReverse
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = Bitwise.Reverse(patterns(Index))
            Next
        Next
    Case mtShift
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = Bitwise.Shift(patterns(Index), 2 * Position - SignBit)
            Next
        Next
    Case mtRotate
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = Bitwise.Rotate(patterns(Index), 2 * Position - SignBit)
            Next
        Next
    Case mtRandom
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = Bitwise.Random
            Next
        Next
    Case mtSwap
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Result = Bitwise.Swap(patterns(Index), Position, SignBit - Position)
            Next
        Next
    Case mtBitString
        For Index = LBound(patterns) To UBound(patterns)
            For Position = 0 To SignBit
                Converted = Bitwise.BitString(patterns(Index))
            Next
        Next
    End Select
    ClassTime = Stopwatch.Halt

End Function
