# vba-bitwise
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
![Platform](https://img.shields.io/badge/Platform-VBA%20(Excel%2C%20Access%2C%20Word%2C%20Outlook%2C%20PowerPoint)-blue)
![Architecture](https://img.shields.io/badge/Architecture-x86%20%7C%20x64-lightgrey)
![Rubberduck](https://img.shields.io/badge/Rubberduck-Ready-orange)

VBA Class for bitwise operations using `LongPtr` (x86 and x64)

A predeclared **Bitwise Operations Class** for VBA with:
- Full set of bitwise operations on `LongPtr` (32-bit on x86, 64-bit on x64)
- Byte-level lookup tables for fast `Count`, `Reverse`, and `BitString`
- Shift, rotate, flip, swap, clear, and random operations
- Zero dependencies (pure VBA)

---

## 📦 Features

- **Predeclared** — accessible as `Bitwise.Method(...)` without instantiation
- **x86 / x64** — uses `LongPtr` throughout; all operations work on 32 or 64 bits
- **Lookup tables** — `Count`, `Reverse`, and `BitString` use precalculated byte tables (3–4× faster than bit-by-bit)
- **Optional `StringTable`** — set `#Const UseStringTable = True` to trade ~5 KB memory for faster `BitString`
- Pure VBA, no external references, Rubberduck-friendly annotations

---

## ⚙️ Public Interface

### Methods

| Member                      | Returns    | Description |
|-----------------------------|------------|-------------|
| `Count(bits)`               | `Long`     | Number of set bits in the pattern. |
| `Flip(bits, No)`            | `LongPtr`  | Flips bit `No`. |
| `Swap(bits, No1, No2)`      | `LongPtr`  | Swaps bits `No1` and `No2`. |
| `Inverse(bits)`             | `LongPtr`  | Inverts all bits (`Not bits`). |
| `Reverse(bits)`             | `LongPtr`  | Reverses the bit order. |
| `Clear(bits, No)`           | `LongPtr`  | Clears bit `No`. |
| `LClear(bits, No)`          | `LongPtr`  | Clears all bits to the **left** of bit `No`. |
| `RClear(bits, No)`          | `LongPtr`  | Clears all bits to the **right** of bit `No`. |
| `LShift(bits, amount)`      | `LongPtr`  | Logical left shift by `amount`. |
| `RShift(bits, amount)`      | `LongPtr`  | Logical right shift by `amount` (fills with 0). |
| `Shift(bits, amount)`       | `LongPtr`  | Left if `amount < 0`, right if `amount > 0`. |
| `Rotate(bits, amount)`      | `LongPtr`  | Left if `amount < 0`, right if `amount > 0`. |
| `Random()`                  | `LongPtr`  | Pseudo-random bit pattern. |

### Properties

| Member                      | Type        | Description |
|-----------------------------|-------------|-------------|
| `Bit(bits, No)` *(Default)* | `Boolean`   | Gets or sets bit `No` in `bits` (passed `ByRef`). |
| `BitArray(bits)`            | `Boolean()` | Converts between a bit pattern and a `Boolean` array. |
| `BitString(bits)`           | `String`    | Converts between a bit pattern and a `"0"`/`"1"` string. |

**Error behavior**
- `BitArray` raises `vbErrorInvalidProcedureCall (5)` for empty, zero-length, or oversized arrays.
- `BitString` raises `vbErrorInvalidProcedureCall (5)` for empty, oversized, or non-binary strings.

---

## 🚀 Quick Start

```vb
Dim bits As LongPtr

' Set and read individual bits
Bitwise.Bit(bits, 3) = True
Debug.Print Bitwise.Bit(bits, 3)    ' -> True
Debug.Print Bitwise.BitString(bits) ' -> "...00001000"

' Shift and rotate
bits = Bitwise.LShift(bits, 2)      ' bit 3 -> bit 5
bits = Bitwise.Rotate(bits, 1)      ' rotate right by 1

' Count set bits
Debug.Print Bitwise.Count(bits)     ' -> 1

' Convert to/from string
Bitwise.BitString(bits) = "00001111"
Debug.Print Bitwise.Count(bits)     ' -> 4
```

---

## ⏱️ Performance

Timings (ms) per operation, measured on Windows x64:

| Method      | Time (ms) | Notes                          |
|-------------|-----------|--------------------------------|
| `Count`     | 0.00022   | Byte lookup table              |
| `Flip`      | 0.00007   |                                |
| `Inverse`   | 0.00007   |                                |
| `Swap`      | 0.00010   |                                |
| `Shift`     | 0.00011   |                                |
| `Rotate`    | 0.00014   |                                |
| `Random`    | 0.00022   |                                |
| `Reverse`   | 0.00033   | Byte lookup table              |
| `BitString` | 0.00034   | Byte lookup table              |

A standard module implementation is typically 10–20% faster than the class.

---

## 🧠 Memory Usage

| Table / Mask     | Size (x64)  |
|-----------------|-------------|
| `BitMask()`     | 544 bytes   |
| `MaxMask()`     | 544 bytes   |
| `ReverseTable()`| 288 bytes   |
| `CountTable()`  | 288 bytes   |
| `StringTable()` | 5,664 bytes *(optional, `#Const UseStringTable = True`)* |

Total without `StringTable`: ~1.7 KB. With: ~7.3 KB.

---

## 📄 License

MIT © 2025 Vincent van Geerestein
