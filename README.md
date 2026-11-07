# vba-bitwise
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

Bitwise operations on `LongPtr` (x86 and x64), as a standard module and as a predeclared class.

- **`BitwiseLib`** — standard module: call `CountBits(bits)`, `RotateBits(bits, 3)`, … directly
- **`Bitwise`** — predeclared class: call `Bitwise.Count(bits)`, `Bitwise.Rotate(bits, 3)`, … without `New`

Both offer the same operations; the module is slightly faster, the class keeps the names short and grouped.

---

## 📦 Features

- Full set of bitwise operations on `LongPtr` (32 bits on x86, 64 bits on x64)
- Count, get, set, flip, swap, clear, invert, reverse, shift, rotate and random patterns
- Conversion to and from a `Boolean` array and a `"0"`/`"1"` bitstring, with `For Each Bit In BitArray(bits)`
- Byte lookup tables make `Count`, `Reverse` and `BitString` 3–4× faster than bit-by-bit processing
- Optional string table (`#Const UseStringTable`) for a 4× faster `BitString`
- Pure VBA, no references, no dependencies; Rubberduck annotations
- Unit tests for every operation, checked against an independent reference

---

## 📁 Files

| File | Description |
|---|---|
| `BitwiseLib.bas` | Standard module, with [Rubberduck](https://rubberduckvba.com/) annotations |
| `BitwiseLib_WithAttributes.bas` | The same module with the VB attributes baked in — no Rubberduck needed |
| `Bitwise.cls` | Predeclared class, with Rubberduck annotations |
| `Bitwise_WithAttributes.cls` | The same class with the VB attributes baked in — no Rubberduck needed |
| `Tests/TestBitwiseOps.bas` | Rubberduck unit tests for `BitwiseLib` |
| `Tests/TestBitwise.bas` | Rubberduck unit tests for `Bitwise` |
| `Tests/BitwiseTestHelpers.bas` | Reference operations on bitstrings used by both test modules |
| `Tests/BenchmarkBitwise.bas` | Benchmark of the module and the class (needs [Stopwatch](https://github.com/vgrstn/VBA-Stopwatch)) |

Import the `_WithAttributes` files if you are not using Rubberduck. Without Rubberduck the plain `.cls` file loses
its default member (`Bit`) and the member descriptions; the `.bas` file only loses the descriptions.

---

## ⚙️ Public Interface

`bits` is a `LongPtr` bit pattern; bit 0 is the least significant bit, bit 63 (x64) or 31 (x86) the sign bit.

### Methods

| `BitwiseLib` (module)        | `Bitwise` (class)         | Returns   | Description |
|------------------------------|---------------------------|-----------|-------------|
| `CountBits(bits)`            | `Count(bits)`             | `Long`    | Number of set bits. |
| `GetBit(bits, No)`           | `Bit(bits, No)`           | `Boolean` | Bit `No`. |
| `SetBit(bits, No)`           | `Bit(bits, No) = True`    | `LongPtr` | Sets bit `No`. |
| `ClearBit(bits, No)`         | `Clear(bits, No)`         | `LongPtr` | Clears bit `No`. |
| `FlipBit(bits, No)`          | `Flip(bits, No)`          | `LongPtr` | Flips bit `No`. |
| `SwapBits(bits, No1, No2)`   | `Swap(bits, No1, No2)`    | `LongPtr` | Swaps bits `No1` and `No2`. |
| `InvertBits(bits)`           | `Invert(bits)`            | `LongPtr` | Inverts every bit (`Not bits`). |
| `ReverseBits(bits)`          | `Reverse(bits)`           | `LongPtr` | Reverses the bit order. |
| `LClearBits(bits, No)`       | `LClear(bits, No)`        | `LongPtr` | Clears the bits to the **left** of bit `No`. |
| `RClearBits(bits, No)`       | `RClear(bits, No)`        | `LongPtr` | Clears the bits to the **right** of bit `No`. |
| `LShiftBits(bits, amount)`   | `LShift(bits, amount)`    | `LongPtr` | Logical left shift. |
| `RShiftBits(bits, amount)`   | `RShift(bits, amount)`    | `LongPtr` | Logical right shift (fills with 0's). |
| `ShiftBits(bits, amount)`    | `Shift(bits, amount)`     | `LongPtr` | Left if `amount < 0`, right if `amount > 0`; 0 when every bit is shifted out. |
| `RotateBits(bits, amount)`   | `Rotate(bits, amount)`    | `LongPtr` | Left if `amount < 0`, right if `amount > 0`; any amount wraps around. |
| `RandomBits()`               | `Random()`                | `LongPtr` | Pseudo-random bit pattern (`VBA.Rnd`). |

### Properties

| Member | Type | Description |
|---|---|---|
| `Bit(bits, No)` *(class, default member)* | `Boolean` | Gets or sets bit `No`; `Bitwise(bits, No)` works too. |
| `BitArray(bits)` | `Boolean()` | Converts a bit pattern to or from a `Boolean` array (element 0 = bit 0). |
| `BitString(bits)` | `String` | Converts a bit pattern to or from a `"0"`/`"1"` string (first character = sign bit). |

The setters (`Bit`, `BitArray` and `BitString` Let) change `bits` itself, which is passed `ByRef`.
A shorter array or string is padded with 0's.

### Errors

All errors are `vbErrorInvalidProcedureCall` (5):

| Property | Description |
|---|---|
| `BitArray` Let | `Array is empty` (not dimensioned), `Array is too long` |
| `BitString` Let | `String is empty`, `String is too long`, `String is invalid` (a character other than 0 or 1) |

---

## 🚀 Quick Start

### Module

```vb
Dim bits As LongPtr

bits = SetBit(bits, 3)
Debug.Print GetBit(bits, 3)      ' True
bits = ShiftBits(bits, -2)       ' left by 2: bit 3 -> bit 5
bits = RotateBits(bits, 1)       ' right by 1: bit 5 -> bit 4
Debug.Print CountBits(bits)      ' 1

BitString(bits) = "1111"
Debug.Print bits                 ' 15

Dim Bit As Variant
For Each Bit In BitArray(bits)
    Debug.Print Bit;             ' True True True True False False ...
Next
```

### Class

```vb
Dim bits As LongPtr

Bitwise.Bit(bits, 3) = True      ' or: Bitwise(bits, 3) = True
Debug.Print Bitwise(bits, 3)     ' True
bits = Bitwise.LShift(bits, 2)   ' bit 3 -> bit 5
bits = Bitwise.Rotate(bits, 1)   ' bit 5 -> bit 4
Debug.Print Bitwise.Count(bits)  ' 1
Debug.Print Bitwise.BitString(bits)
```

---

## ⏱️ Performance

Time per call in µs, Windows x64, `RunBenchmarks` in `Tests/BenchmarkBitwise.bas`
(6,400,000 calls per method, loop overhead subtracted):

| Method      | Module | Class | Notes |
|-------------|-------:|------:|-------|
| `Count`     | 0.20 | 0.22 | byte lookup table |
| `Flip`      | 0.04 | 0.07 | |
| `Invert`    | 0.02 | 0.06 | |
| `Reverse`   | 0.24 | 0.25 | byte lookup table |
| `Shift`     | 0.08 | 0.10 | |
| `Rotate`    | 0.11 | 0.13 | |
| `Random`    | 0.19 | 0.21 | |
| `Swap`      | 0.07 | 0.09 | |
| `BitString` | 0.30 | 0.34 | string lookup table (1.25 µs without) |

A class call costs about 0.02–0.03 µs more than a module call. For the fastest methods
(Flip, Invert, Swap, Shift) that is 35–150%, for the others 5–20%.

---

## 🔧 Compiler directive

```vb
#Const UseStringTable = True
```

The string table holds the bitstring of every byte value. It makes `BitString` about 4× faster but uses more
memory than all other tables and masks together. Set it to `False` to save memory.

| Table / mask     | Size (x64) |
|------------------|-----------:|
| `BitMask()`      | 544 bytes |
| `MaxMask()`      | 544 bytes |
| `ReverseTable()` | 288 bytes |
| `CountTable()`   | 288 bytes |
| `StringTable()`  | 5,664 bytes |

Total about 1.7 KB without the string table, 7.3 KB with it.

---

## 🧠 How it works

- **Masks** — `BitMask(i)` has only bit `i` set, `MaxMask(i)` all bits below `i`. Every single-bit operation is
  one `And`, `Or` or `Xor` with a mask.
- **Shifts** — a left shift by `i` is a multiplication by `2 ^ i`, a right shift a division. VBA uses two's
  complement, so the bits that would overflow are cleared before multiplying, and for a negative pattern the sign
  bit is masked before dividing and put back at its shifted position. A rotation combines a right and a left shift.
- **Lookup tables** — `Count`, `Reverse` and `BitString` process a pattern byte by byte with precalculated
  256-entry tables instead of bit by bit.

---

## 📝 Implementation notes

- The module initializes its masks and tables on first use; the class does it in `Class_Initialize`.
- `RandomBits` / `Random` build a pattern from 16-bit pieces of `VBA.Rnd`. `Randomize` is called once: by the
  module on the first call of any of its functions, by the class in `Class_Initialize`.
- **VBA crash:** VBA crashes the host application when an array that a procedure received `ByRef` is passed on
  as the value of an array `Property Let`. This is a VBA bug, not specific to `BitArray`. Copy the array to a
  local array first:

  ```vb
  Private Sub SetPattern(ByRef bits As LongPtr, ByRef bitValues() As Boolean)
      Dim Values() As Boolean: Values = bitValues
      BitArray(bits) = Values          ' BitArray(bits) = bitValues would crash
  End Sub
  ```

---

## 🧪 Tests

The `Tests` folder holds 47 [Rubberduck](https://rubberduckvba.com/) unit tests (23 for the module, 24 for the
class). Each test runs over edge cases (0, ±1, sign bit, maximum, alternating patterns) and 40 random patterns,
for every bit position or shift amount, and compares the result with a reference in `BitwiseTestHelpers` that
works on strings with `Hex$` only, independent of the code under test.

- The test modules use early binding: they need a reference to the Rubberduck type library. Set
  `#Const LateBinding = True` to run them without it.
- `BenchmarkBitwise` is a normal module, so it does not slow down the Test Explorer. Run it from the Immediate
  window with `RunBenchmarks` (or `RunBenchmarks 10000` for a quick run). It needs the
  [Stopwatch class](https://github.com/vgrstn/VBA-Stopwatch).

---

## 📄 License

MIT © 2025–2026 Vincent van Geerestein
