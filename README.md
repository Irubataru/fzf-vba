fzf-vba
=======

fzf-vba is an implementation of [fzf][fzf-gh] a general-purpose fuzzy finder.

![fzf-vba example preview](https://raw.githubusercontent.com/Irubataru/img/master/fzf-vba/fzf-vba-preview-animated.gif)

It contains both functions to be used by VBA for VBA as well as some functions
that can be used as functions directly in Excel.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
  - [`FzfAlgorithm`](#fzfalgorithm)
  - [`FzfUtilities`](#fzfutilities)
  - [`FzfWorkbookMethods`](#fzfworkbookmethods)
- [License](#license)


## Installation

The installation varies depending on which features you need, but basically it
all boils down to importing the files in [fzf-vba](fzf-vba) in your VBA project.

  * **Minimal**: Import [fzf-vba/FzfAlgorithm.cls](fzf-vba/FzfAlgorithm.cls) and
      [fzf-vba/FzfResult.cls](fzf-vba/FzfResult.cls).
  * **VBA only**: Also import [fzf-vba/FzfUtilities.cls](fzf-vba/FzfUtilities.cls).
  * **Worksheet functions**: Also import [fzf-vba/FzfWorkbookMethods.bas](fzf-vba/FzfWorkbookMethods.bas).

Optionally if you have the [Rubberduck VBE add-in][rubberduck-gh] you can also
import the unit tests in [Tests](Tests).

There is an [Excel-DNA/IntelliSense][intellisense-gh] csv sheet in [bin](bin)
for those who use that add-in.

Finally there is an example in the [example](example) folder (where the preview
is taken from), more information on the [wiki][wiki-example].

## Usage

### `FzfAlgorithm`

```vb
Public Function FuzzyMatchV1( _
        ByVal Text As String, _
        ByVal Pattern As String, _
        Optional ByVal WithPositions As Boolean = False, _
        Optional ByVal WithNormalize As Boolean = True) As FzfResult
```

The base of the entire library, it does a fuzzy match on `Text` using the
`Pattern` and returns information about the match. If `WithPositions` is true it
will also return the positions of the characters in the match. If
`WithNormalize` is true it will run `Text` through the `Normalize` function
before matching.

The return value is basically a struct with the following items

```vb
Type FzfResult
    Score As Long
    StartIndex As Long
    EndIndex As Long
    Positions() As Long
End Type
```

If `Score = 0` that means that there was no match, otherwise `Score` will
generally be larger than 0. The algorithm is the V1 algorithm of the
[fzf][fzf-gh] library which searches for the first match that has the shortest
substring. Due to various character position bonuses this is not guaranteed to
return the match that has the highest score. The V2 version which basically is a
Smith-Waterman algorithm might be implemented in the future.

The `Positions` return variable is implemented as a `Variant` and will be
`Empty` if `WithPositions` is false.

The function only accepts ASCII characters, or characters that are ASCII after a
pass through the `Normalize` function. It is also only does a case insensitive
fuzzy match.

**Future improvements**:
- [ ] Implement the V2 algorithm.
- [ ] Support non-ASCII characters.
- [ ] Support a case sensitive version.

---

```vb
Public Function Normalize(ByVal Value As String) As String
```

Returns a normalized version of the string where character accents etc have been
stripped from the letters.

*Example*: `Normalize("Ḥɇƚɭø") => "Hello"`

---

### `FzfUtilities`

```vb
Public Function SortAndFilter( _
        ByVal Texts As Variant, _
        ByVal Pattern As String, _
        Optional ByVal ScoreThreshold As Long = 0, _
        Optional OnlyTopN As Long = -1) As String()
```

Takes a single string or an array of string and applies the fuzzy match on it.
Filter out all matches that have a score equal to or lower than
`ScoreThreshold`, finally sorts the result and returns an array of strings
sorted by score. If `OnlyTopN` is specified (and not `=-1` ) it will only
include the `N` results with the highest score. If there are fewer matches than
`N` then all matches will be returned but the array will not be padded. The
result array is 0-indexed.

---

```vb
Public Function SortAndFilterStrict( _
        ByRef Texts() As String, _
        ByVal Pattern As String, _
        Optional ByVal ScoreThreshold As Long = 0, _
        Optional OnlyTopN As Long = -1) As String()
```

The same as the `SortAndFilter` function however with stricter typing on the
arguments (specifically `Texts`). If you already have a 0-indexed array of
strings then this is slightly cheaper to call.
 
---

### `FzfWorkbookMethods`

```vb
Public Function Fzf( _
        ByVal Value As Variant, _
        ByVal Pattern As Variant, _
        Optional ByVal OnlyTopN As Long = 1, _
        Optional ByVal ScoreThreshold As Long = 0, _
        Optional ByVal TransposeResult As Boolean = False) As Variant
```

Applies `SortAndFilter` to `Value` and returns the result. The return value can
be either a value (if there is only one result) or an array. The array is a
column unless `TransposeResult` is true.

`Value` can be either a value or a range that can contain multiple cells (but
not multiple areas). If it is a matrix (`Rows>1` and `Columns>1`) then it will
take all values (flatten the matrix).

Pattern can also be a cell or a value but has to be a single cell range.
 
---

```vb
Public Function FzfScore(ByVal Value As Variant, ByVal Pattern As Variant) As Variant
```

Returns the score of fuzzy matching `Value` with `Pattern`. Both have to
represent a single string value (either as value or a single cell range).

---

```vb
Public Function FzfStartIndex(ByVal Value As Variant, ByVal Pattern As Variant) As Variant
```

Returns the start index of fuzzy matching `Value` with `Pattern`. Both have to
represent a single string value (either as value or a single cell range).

---

```vb
Public Function FzfEndIndex(ByVal Value As Variant, ByVal Pattern As Variant) As Variant
```

Returns the end index of fuzzy matching `Value` with `Pattern`. Both have to
represent a single string value (either as value or a single cell range).

---

```vb
Public Function FzfNormalize(ByVal Value As Variant) As String
```

Returns the normalized text string after apllying `FzfAlgorithm.Normalize` to
it. `Value` has to represent a single string value (either as value or a single
cell range).


## License

MIT

Copyright (c) 2020 Jonas R. Glesaaen

fzf copyright (c) 2013-2020 Junegunn Choi


[fzf-gh]: https://github.com/junegunn/fzf
[rubberduck-gh]: https://github.com/rubberduck-vba/Rubberduck
[intellisense-gh]: https://github.com/Excel-DNA/IntelliSense
[wiki-example]: https://github.com/Irubataru/fzf-vba/wiki/Example-usage
