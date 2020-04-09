VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FzfAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Namespace class; do not create an instance."
' =====================================================================================================================
' fzf-vba v0.1
' Copyright (c) 2020 Jonas R. Glesaaen (jonas@glesaaen.com)
'
' Implementation of fzf for VBA.
'
' Author: Jonas R. Glesaaen (jonas@glesaaen.com)
' License: MIT
'
' fzf author: Junegunn Choi
' fzf repo:   https://github.com/junegunn/fzf
' =====================================================================================================================

'@Folder("fzf-vba")
'@PredeclaredId
'@ModuleDescription("Namespace class; do not create an instance.")
Option Explicit

' Weights used to calculate the score
Private Const ScoreMatch As Long = 16
Private Const ScoreGapStart As Long = -3
Private Const ScoreGapExtention As Long = -1

' We prefer matches at the beginning of a word, but the bonus should not be too
' great to prevent the longer acronym matches from always winning over shorter
' fuzzy matches. The bonus point here was specifically chosen that the bonus is
' cancelled when the gap between the acronyms grows over 8 characters, which is
' approximately the average length of the words found in web2 dictionary and my
' file system.
Private Const BonusBoundary As Long = ScoreMatch \ 2

' Although bonus point for non-word characters is non-contextual, we need it for
' computing bonus points for consecutive chunks starting with a non-word
' character.
Private Const BonusNonWord As Long = ScoreMatch \ 2

' Edge-triggered bonus for matches in camelCase words. Compared to
' word-boundary case, they don't accompany single-character gaps
' (e.g. FooBar vs. foo-bar), so we deduct bonus point accordingly.
Private Const BonusCamel123 As Long = BonusBoundary + ScoreGapExtention

' Minimum bonus point given to characters in consecutive chunks. Note that
' bonus points for consecutive matches shouldn't have needed if we used fixed
' match score as in the original algorithm.
Private Const BonusConsecutive  As Long = -(ScoreGapStart + ScoreGapExtention)

' The first character in the typed pattern usually has more significance than
' the rest so it's important that it appears at special positions where bonus
' points are given. e.g. "to-go" vs. "ongoing" on "og" or on "ogo". The amount
' of the extra bonus should be limited so that the gap penalty is still
' respected.
Private Const BonusFirstCharMultiplier = 2

' ASCII constants
Private Const AsciiCode_Lower_A As Long = 97
Private Const AsciiCode_Lower_Z As Long = 122
Private Const AsciiCode_Upper_A As Long = 65
Private Const AsciiCode_Upper_Z As Long = 90
Private Const AsciiCode_0 As Long = 48
Private Const AsciiCode_9 As Long = 57

Private Const AsciiCaseSkip As Long = 32

Private Enum CharacterClass
    charNonWord
    charLower
    charUpper
    charNumber
End Enum

Private Type ScoreReturnType
    Score As Long
    Positions As Variant
End Type

' ---------------------------------------------------------------------------------------------------------------------
' --- Public methods --------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Fuzzy match a string with a pattern and return it's score.")
Public Function FuzzyMatchV1( _
        ByVal Text As String, _
        ByVal Pattern As String, _
        Optional ByVal WithPositions As Boolean = False, _
        Optional ByVal WithNormalize As Boolean = True) As FzfResult
Attribute FuzzyMatchV1.VB_Description = "Fuzzy match a string with a pattern and return it's score."

' The simple fuzzy match algorithm searches for the first match by first doing a
' forward sweep to find all the characters in the pattern (ignoring case)
' followed by a backwards sweep to see if there exists a shorter match.
'
' Example:  Text: "a____b___ab______c", Pattern: "abc"
'
' Forward sweep:  "a____b___ab______c"
'                  ^start           ^end
'
' Backward sweep: "a____b___ab______c"
'                           ^start  ^end
'
' This is not guaranteed to find the match that has the highest score which is
' why the original fzf library has two algorithms with the correct one using a
' modified Smith-Waterman algorithm to accomplish this. Implementing the second
' algorithm is a possible future enhancement if the need is there.
'
' There is one difference in the return value from that of the original fzf
' implementation and that is that the end index will be inclusive (i.e. one
' less than the original) due to the fact that VBA generally does inclusive
' loops).
'
' The WithNormalize argument determines whether or not the string should be
' ASCII normalized before matching.

    If VBA.Len(Pattern) = 0 Then
        Set FuzzyMatchV1 = FzfResult(0, 0, 0, Empty)
        Exit Function
    End If
    
    If VBA.Len(Text) = 0 Then
        Set FuzzyMatchV1 = FzfResult(-1, -1, 0, Empty)
        Exit Function
    End If
    
    If WithNormalize Then
        Text = Normalize(Text)
    End If
    
    If Not IsAscii(Text) Then
        Err.Raise _
            Number:=13, _
            Source:=VBA.TypeName(Me) & ".FuzzyMatchV1", _
            Description:="Text parameter is not an ASCII string."
    ElseIf Not IsAscii(Pattern) Then
        Err.Raise _
            Number:=13, _
            Source:=VBA.TypeName(Me) & ".FuzzyMatchV1", _
            Description:="Pattern parameter is not an ASCII string."
    End If
    
    ' First store the arrays as ascii tables to optimise search speed
    Dim TextByteArray() As Byte
    TextByteArray = ToLCaseByteArray(Text)
    
    Dim PatternByteArray() As Byte
    PatternByteArray = ToLCaseByteArray(Pattern)
    
    ' Step 0: Check if the string matches
    If AsciiFuzzyIndex(TextByteArray, PatternByteArray) < 0 Then
        Set FuzzyMatchV1 = FzfResult(-1, -1, 0, Empty)
        Exit Function
    End If
    
    Dim pIdx As Long: pIdx = 0
    Dim startIdx As Long: startIdx = -1
    Dim endIdx As Long: endIdx = -1
    
    ' Step 1: Loop forward to find the first fuzzy match
    Dim idx As Long
    For idx = 0 To UBound(TextByteArray)
    
        If TextByteArray(idx) = PatternByteArray(pIdx) Then
            
            If startIdx < 0 Then
                startIdx = idx
            End If
            
            If pIdx = UBound(PatternByteArray) Then
                endIdx = idx
                Exit For
            End If
            
            pIdx = pIdx + 1
            
        End If
    
    Next idx
    
    ' Step 2: loop back to see if we find a shorter fuzzy match
    If startIdx >= 0 And endIdx >= 0 Then
    
        pIdx = pIdx - 1
        
        If pIdx >= 0 Then
            For idx = endIdx - 1 To startIdx Step -1
            
                If TextByteArray(idx) = PatternByteArray(pIdx) Then
                
                    If pIdx = 0 Then
                        startIdx = idx
                        Exit For
                    End If
                    
                    pIdx = pIdx - 1
                
                End If
            
            Next idx
        End If
        
        ' Calculate the score of the match and exit
        Dim Score As ScoreReturnType
        Score = CalculateScore(ToByteArray(Text), PatternByteArray, startIdx, endIdx, WithPositions)
        
        Set FuzzyMatchV1 = FzfResult(startIdx, endIdx, Score.Score, Score.Positions)
    
    Else
        Set FuzzyMatchV1 = FzfResult(-1, -1, 0, Empty)
    End If

End Function

'@Description("Normalize latin script letters.")
Public Function Normalize(ByVal Value As String) As String
Attribute Normalize.VB_Description = "Normalize latin script letters."

    Dim ByteTable() As Byte
    ByteTable = Value
    
    Dim i As Long
    For i = 0 To UBound(ByteTable) Step 2
    
        Dim Compressed As Byte
        Compressed = NormalizeCharacter(ByteTable(i), ByteTable(i + 1))
        
        If Compressed <> 0 Then
            ByteTable(i) = Compressed
            ByteTable(i + 1) = 0
        End If
    
    Next i
    
    Normalize = ByteTable

End Function

' ---------------------------------------------------------------------------------------------------------------------
' --- Private methods -------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Calculate the score of a match of a substring determined by two indices.")
Private Function CalculateScore( _
        ByRef Text() As Byte, ByRef Pattern() As Byte, _
        ByVal startIdx As Long, ByVal endIdx As Long, _
        Optional ByVal WithPositions As Boolean = False) As ScoreReturnType
Attribute CalculateScore.VB_Description = "Calculate the score of a match of a substring determined by two indices."
    
    'The weights are specified by private constants at the top of the file.
    
    Dim pIdx As Long: pIdx = 0
    Dim Score As Long: Score = 0
    Dim InGap As Boolean: InGap = False
    Dim Consecutive As Long: Consecutive = 0
    Dim FirstBonus As Long: FirstBonus = 0
    
    Dim Positions As Variant
    If WithPositions Then
           ReDim Positions(0 To UBound(Pattern)) As Long
    End If
    
    Dim PrevClass As CharacterClass
    If startIdx > 0 Then
        PrevClass = CharClassOfAsciiByte(Text(startIdx - 1))
    Else
        PrevClass = charNonWord
    End If
    
    Dim idx As Long
    For idx = startIdx To endIdx
    
        Dim CurrentClass As CharacterClass
        CurrentClass = CharClassOfAsciiByte(Text(idx))
    
        ' Case insensitive match
        If ToLower(Text(idx)) = Pattern(pIdx) Then
        
            If WithPositions Then
                ' VBA strings are 1-indexed
                Positions(pIdx) = idx + 1
            End If
        
            Score = Score + ScoreMatch
            
            Dim Bonus As Long
            Bonus = BonusFor(PrevClass, CurrentClass)
            
            If Consecutive = 0 Then
                FirstBonus = Bonus
            Else
            
                If Bonus = BonusBoundary Then
                    FirstBonus = Bonus
                End If
            
                Bonus = Max(Max(Bonus, FirstBonus), BonusConsecutive)
            
            End If
            
            If pIdx = 0 Then
                Score = Score + Bonus * BonusFirstCharMultiplier
            Else
                Score = Score + Bonus
            End If
            
            InGap = False
            Consecutive = Consecutive + 1
            pIdx = pIdx + 1
            
        Else
            
            ' Gap penalties
            If InGap Then
                Score = Score + ScoreGapExtention
            Else
                Score = Score + ScoreGapStart
            End If
            
            InGap = True
            Consecutive = 0
            FirstBonus = 0
            
        End If
        
        PrevClass = CurrentClass
    
    Next idx
    
    CalculateScore.Score = Score
    CalculateScore.Positions = Positions

End Function

'@Description("First index of a fuzzy match, -1 if it doesn't match.")
Private Function AsciiFuzzyIndex(ByRef Text() As Byte, ByRef Pattern() As Byte) As Long
Attribute AsciiFuzzyIndex.VB_Description = "First index of a fuzzy match, -1 if it doesn't match."

    Dim firstIdx As Long: firstIdx = 0
    Dim nextIdx As Long: nextIdx = 0
    
    Dim pIdx As Long
    For pIdx = 0 To UBound(Pattern)

        nextIdx = TrySkip(Text, Pattern(pIdx), nextIdx)
        
        If nextIdx < 0 Then
            AsciiFuzzyIndex = -1
            Exit Function
        End If
        
        If pIdx = 0 Then
            firstIdx = nextIdx
        End If
        
        nextIdx = nextIdx + 1
    
    Next pIdx
    
    AsciiFuzzyIndex = firstIdx

End Function

'@Description("Search for and return the index of a the next character starting at an index, -1 if it doesn't find it.")
Private Function TrySkip(ByRef Text() As Byte, ByVal toFind As Byte, ByVal startIdx As Long) As Long
Attribute TrySkip.VB_Description = "Search for and return the index of a the next character starting at an index, -1 if it doesn't find it."

    Dim idx As Long
    For idx = startIdx To UBound(Text)
    
        If Text(idx) = toFind Then
            TrySkip = idx
            Exit Function
        End If
    
    Next idx
    
    TrySkip = -1

End Function

'@Description("Returns the character class of an ASCII Byte.")
Private Function CharClassOfAsciiByte(ByVal Char As Byte) As CharacterClass
Attribute CharClassOfAsciiByte.VB_Description = "Returns the character class of an ASCII Byte."

    Select Case True
        Case Char >= AsciiCode_Lower_A And Char <= AsciiCode_Lower_Z
            CharClassOfAsciiByte = charLower
        Case Char >= AsciiCode_Upper_A And Char <= AsciiCode_Upper_Z
            CharClassOfAsciiByte = charUpper
        Case Char >= AsciiCode_0 And Char <= AsciiCode_9
            CharClassOfAsciiByte = charNumber
        Case Else
            CharClassOfAsciiByte = charNonWord
    End Select

End Function

'@Description("Score bonus depending on the current and preceeding character classes.")
Private Function BonusFor(ByVal PrevClass As CharacterClass, ByVal CurrentClass As CharacterClass) As Long
Attribute BonusFor.VB_Description = "Score bonus depending on the current and preceeding character classes."

    Select Case True
        Case PrevClass = charNonWord And CurrentClass <> charNonWord
            BonusFor = BonusBoundary
            
        Case (PrevClass = charLower And CurrentClass = charUpper) Or _
             (PrevClass <> charNumber And CurrentClass = charNumber)
            BonusFor = BonusCamel123
            
        Case CurrentClass = charNonWord
            BonusFor = BonusNonWord
            
        Case Else
            BonusFor = 0
            
    End Select

End Function

' ---------------------------------------------------------------------------------------------------------------------
' --- ASCII Utilities -------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Checks whether a string is an ASCII string.")
Private Function IsAscii(ByVal Str As String) As Boolean
Attribute IsAscii.VB_Description = "Checks whether a string is an ASCII string."

    Dim AsByteArray() As Byte
    AsByteArray = Str

    IsAscii = False

    Dim i As Long
    For i = 1 To UBound(AsByteArray) Step 2
        If AsByteArray(i) <> 0 Then Exit Function
    Next i

    IsAscii = True

End Function

'@Description("Convert an ASCII string to a Byte array.")
Private Function ToByteArray(ByVal Str As String) As Byte()
Attribute ToByteArray.VB_Description = "Convert an ASCII string to a Byte array."

    Dim AsByteArray() As Byte
    AsByteArray = Str

    Dim Result() As Byte
    ReDim Result(0 To VBA.Len(Str) - 1) As Byte
    
    Dim i As Long
    For i = 0 To UBound(AsByteArray) - 1 Step 2
        Result(i \ 2) = AsByteArray(i)
    Next i
    
    ToByteArray = Result
    
End Function

'@Description("Convert an ASCII string to a Byte array and converts it to lower case.")
Private Function ToLCaseByteArray(ByVal Str As String) As Byte()
Attribute ToLCaseByteArray.VB_Description = "Convert an ASCII string to a Byte array and converts it to lower case."

    Dim AsByteArray() As Byte
    AsByteArray = Str

    Dim Result() As Byte
    ReDim Result(0 To VBA.Len(Str) - 1) As Byte
    
    Dim i As Long
    For i = LBound(AsByteArray) To UBound(AsByteArray) - 1 Step 2
        Result(i \ 2) = ToLower(AsByteArray(i))
    Next i
    
    ToLCaseByteArray = Result

End Function

'@Description("Lower case of an ASCII byte code.")
Private Function ToLower(ByVal Char As Byte) As Byte
Attribute ToLower.VB_Description = "Lower case of an ASCII byte code."

    If Char >= AsciiCode_Upper_A And Char <= AsciiCode_Upper_Z Then
        ToLower = Char + AsciiCaseSkip
    Else
        ToLower = Char
    End If

End Function

'@Description("Normalize a latin script letter descibed by two bytes.")
Private Function NormalizeCharacter(ByVal FirstByte As Byte, ByVal SecondByte As Byte) As Byte
Attribute NormalizeCharacter.VB_Description = "Normalize a latin script letter descibed by two bytes."

    Dim Character As Long
    Character = FirstByte + 256 * SecondByte
    
    Select Case Character
        Case 225 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 259 'WITH BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 462 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 226 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 228 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 551 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 7841 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 513 'WITH DOUBLE GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 224 'WITH GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 7843 'WITH HOOK ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 515 'WITH INVERTED BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 257 'WITH MACRON, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 261 'WITH OGONEK, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 7834 'WITH RIGHT HALF RING, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 229 'WITH RING ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 7681 'WITH RING BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 227 'WITH TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 867 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 97
        Case 592 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 97
        Case 7683 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 98
        Case 7685 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 98
        Case 595 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 98
        Case 7687 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 98
        Case 384 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 98
        Case 387 'WITH TOPBAR, LATIN SMALL LETTER
            NormalizeCharacter = 98
        Case 263 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 269 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 231 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 265 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 597 'WITH CURL, LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 267 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 392 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 572 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 872 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 99
        Case 663 ', LATIN LETTER STRETCHED
            NormalizeCharacter = 99
        Case 8580 ', LATIN SMALL LETTER REVERSED
            NormalizeCharacter = 99
        Case 271 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 7697 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 7699 'WITH CIRCUMFLEX BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 545 'WITH CURL, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 7691 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 7693 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 599 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 7695 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 273 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 598 'WITH TAIL, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 396 'WITH TOPBAR, LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 873 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 100
        Case 233 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 277 'WITH BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 283 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 553 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 7705 'WITH CIRCUMFLEX BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 234 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 235 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 279 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 7865 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 517 'WITH DOUBLE GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 232 'WITH GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 7867 'WITH HOOK ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 605 'WITH HOOK, LATIN SMALL LETTER REVERSED OPEN
            NormalizeCharacter = 101
        Case 519 'WITH INVERTED BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 275 'WITH MACRON, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 281 'WITH OGONEK, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 583 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 7707 'WITH TILDE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 7869 'WITH TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 868 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 101
        Case 666 ', LATIN SMALL LETTER CLOSED OPEN
            NormalizeCharacter = 101
        Case 606 ', LATIN SMALL LETTER CLOSED REVERSED OPEN
            NormalizeCharacter = 101
        Case 603 ', LATIN SMALL LETTER OPEN
            NormalizeCharacter = 101
        Case 600 ', LATIN SMALL LETTER REVERSED
            NormalizeCharacter = 101
        Case 604 ', LATIN SMALL LETTER REVERSED OPEN
            NormalizeCharacter = 101
        Case 477 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 101
        Case 7432 ', LATIN SMALL LETTER TURNED OPEN
            NormalizeCharacter = 101
        Case 7711 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 102
        Case 402 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 102
        Case 501 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 287 'WITH BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 487 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 291 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 285 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 289 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 608 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 7713 'WITH MACRON, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 485 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 103
        Case 609 ', LATIN SMALL LETTER SCRIPT
            NormalizeCharacter = 103
        Case 7723 'WITH BREVE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 543 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 7721 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 293 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 7719 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 7715 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 7717 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 686 'WITH FISHHOOK, LATIN SMALL LETTER TURNED
            NormalizeCharacter = 104
        Case 614 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 7830 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 295 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 874 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 104
        Case 613 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 104
        Case 8341 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 104
        Case 237 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 301 'WITH BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 464 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 238 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 239 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 7883 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 521 'WITH DOUBLE GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 236 'WITH GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 7881 'WITH HOOK ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 523 'WITH INVERTED BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 299 'WITH MACRON, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 303 'WITH OGONEK, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 616 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 7725 'WITH TILDE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 297 'WITH TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 869 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 305 ', LATIN SMALL LETTER DOTLESS
            NormalizeCharacter = 105
        Case 7433 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 105
        Case 7522 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 105
        Case 8305 ', SUPERSCRIPT LATIN SMALL LETTER
            NormalizeCharacter = 105
        Case 496 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 106
        Case 309 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 106
        Case 669 'WITH CROSSED-TAIL, LATIN SMALL LETTER
            NormalizeCharacter = 106
        Case 585 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 106
        Case 607 'WITH STROKE, LATIN SMALL LETTER DOTLESS
            NormalizeCharacter = 106
        Case 567 ', LATIN SMALL LETTER DOTLESS
            NormalizeCharacter = 106
        Case 7729 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 107
        Case 489 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 107
        Case 311 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 107
        Case 7731 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 107
        Case 409 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 107
        Case 7733 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 107
        Case 670 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 107
        Case 8342 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 107
        Case 314 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 410 'WITH BAR, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 620 'WITH BELT, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 318 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 316 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 7741 'WITH CIRCUMFLEX BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 564 'WITH CURL, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 7735 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 7739 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 320 'WITH MIDDLE DOT, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 619 'WITH MIDDLE TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 621 'WITH RETROFLEX HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 322 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 108
        Case 8343 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 108
        Case 7743 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 109
        Case 7745 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 109
        Case 7747 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 109
        Case 625 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 109
        Case 624 'WITH LONG LEG, LATIN SMALL LETTER TURNED
            NormalizeCharacter = 109
        Case 875 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 109
        Case 7455 ', LATIN SMALL LETTER SIDEWAYS TURNED
            NormalizeCharacter = 109
        Case 623 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 109
        Case 8344 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 109
        Case 324 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 328 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 326 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 7755 'WITH CIRCUMFLEX BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 565 'WITH CURL, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 7749 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 7751 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 505 'WITH GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 626 'WITH LEFT HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 7753 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 414 'WITH LONG RIGHT LEG, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 627 'WITH RETROFLEX HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 241 'WITH TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 110
        Case 8345 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 110
        Case 243 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 335 'WITH BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 466 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 244 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 246 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 559 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 7885 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 337 'WITH DOUBLE ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 525 'WITH DOUBLE GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 242 'WITH GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 7887 'WITH HOOK ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 417 'WITH HORN, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 527 'WITH INVERTED BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 333 'WITH MACRON, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 491 'WITH OGONEK, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 248 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 7443 'WITH STROKE, LATIN SMALL LETTER SIDEWAYS
            NormalizeCharacter = 111
        Case 245 'WITH TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 870 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 111
        Case 629 ', LATIN SMALL LETTER BARRED
            NormalizeCharacter = 111
        Case 7447 ', LATIN SMALL LETTER BOTTOM HALF
            NormalizeCharacter = 111
        Case 596 ', LATIN SMALL LETTER OPEN
            NormalizeCharacter = 111
        Case 7441 ', LATIN SMALL LETTER SIDEWAYS
            NormalizeCharacter = 111
        Case 7442 ', LATIN SMALL LETTER SIDEWAYS OPEN
            NormalizeCharacter = 111
        Case 7446 ', LATIN SMALL LETTER TOP HALF
            NormalizeCharacter = 111
        Case 7765 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 112
        Case 7767 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 112
        Case 421 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 112
        Case 8346 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 112
        Case 587 'WITH HOOK TAIL, LATIN SMALL LETTER
            NormalizeCharacter = 113
        Case 672 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 113
        Case 341 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 345 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 343 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 7769 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 7771 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 529 'WITH DOUBLE GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 638 'WITH FISHHOOK, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 639 'WITH FISHHOOK, LATIN SMALL LETTER REVERSED
            NormalizeCharacter = 114
        Case 635 'WITH HOOK, LATIN SMALL LETTER TURNED
            NormalizeCharacter = 114
        Case 531 'WITH INVERTED BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 7775 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 636 'WITH LONG LEG, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 634 'WITH LONG LEG, LATIN SMALL LETTER TURNED
            NormalizeCharacter = 114
        Case 589 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 637 'WITH TAIL, LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 876 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 114
        Case 633 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 114
        Case 7523 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 114
        Case 347 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 353 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 351 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 349 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 537 'WITH COMMA BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 7777 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 7835 'WITH DOT ABOVE, LATIN SMALL LETTER LONG
            NormalizeCharacter = 115
        Case 7779 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 642 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 575 'WITH SWASH TAIL, LATIN SMALL LETTER
            NormalizeCharacter = 115
        Case 383 ', LATIN SMALL LETTER LONG
            NormalizeCharacter = 115
        Case 223 ', LATIN SMALL LETTER SHARP
            NormalizeCharacter = 115
        Case 8347 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 115
        Case 357 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 355 'WITH CEDILLA, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 7793 'WITH CIRCUMFLEX BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 539 'WITH COMMA BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 566 'WITH CURL, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 7831 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 7787 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 7789 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 429 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 7791 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 427 'WITH PALATAL HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 648 'WITH RETROFLEX HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 359 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 877 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 116
        Case 647 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 116
        Case 8348 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 116
        Case 649 'BAR, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 250 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 365 'WITH BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 468 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 7799 'WITH CIRCUMFLEX BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 251 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 7795 'WITH DIAERESIS BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 252 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 7909 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 369 'WITH DOUBLE ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 533 'WITH DOUBLE GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 249 'WITH GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 7911 'WITH HOOK ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 432 'WITH HORN, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 535 'WITH INVERTED BREVE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 363 'WITH MACRON, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 371 'WITH OGONEK, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 367 'WITH RING ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 7797 'WITH TILDE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 361 'WITH TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 871 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 117
        Case 7453 ', LATIN SMALL LETTER SIDEWAYS
            NormalizeCharacter = 117
        Case 7454 ', LATIN SMALL LETTER SIDEWAYS DIAERESIZED
            NormalizeCharacter = 117
        Case 7524 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 117
        Case 7807 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 118
        Case 651 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 118
        Case 7805 'WITH TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 118
        Case 878 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 118
        Case 652 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 118
        Case 7525 ', LATIN SUBSCRIPT SMALL LETTER
            NormalizeCharacter = 118
        Case 7811 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 119
        Case 373 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 119
        Case 7813 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 119
        Case 7815 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 119
        Case 7817 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 119
        Case 7809 'WITH GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 119
        Case 7832 'WITH RING ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 119
        Case 653 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 119
        Case 7821 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 120
        Case 7819 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 120
        Case 879 ', COMBINING LATIN SMALL LETTER
            NormalizeCharacter = 120
        Case 253 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 375 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 255 'WITH DIAERESIS, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 7823 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 7925 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 7923 'WITH GRAVE, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 7927 'WITH HOOK ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 436 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 563 'WITH MACRON, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 7833 'WITH RING ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 591 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 7929 'WITH TILDE, LATIN SMALL LETTER
            NormalizeCharacter = 121
        Case 654 ', LATIN SMALL LETTER TURNED
            NormalizeCharacter = 121
        Case 378 'WITH ACUTE, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 382 'WITH CARON, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 7825 'WITH CIRCUMFLEX, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 657 'WITH CURL, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 380 'WITH DOT ABOVE, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 7827 'WITH DOT BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 549 'WITH HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 7829 'WITH LINE BELOW, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 656 'WITH RETROFLEX HOOK, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 438 'WITH STROKE, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 576 'WITH SWASH TAIL, LATIN SMALL LETTER
            NormalizeCharacter = 122
        Case 593 ', latin small letter script
            NormalizeCharacter = 97
        Case 193 'WITH ACUTE, LATIN CAPITAL LETTER
            NormalizeCharacter = 65
        Case 194 'WITH CIRCUMFLEX, LATIN CAPITAL LETTER
            NormalizeCharacter = 65
        Case 196 'WITH DIAERESIS, LATIN CAPITAL LETTER
            NormalizeCharacter = 65
        Case 192 'WITH GRAVE, LATIN CAPITAL LETTER
            NormalizeCharacter = 65
        Case 197 'WITH RING ABOVE, LATIN CAPITAL LETTER
            NormalizeCharacter = 65
        Case 570 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 65
        Case 195 'WITH TILDE, LATIN CAPITAL LETTER
            NormalizeCharacter = 65
        Case 7424 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 65
        Case 385 'WITH HOOK, LATIN CAPITAL LETTER
            NormalizeCharacter = 66
        Case 579 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 66
        Case 665 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 66
        Case 7427 ', LATIN LETTER SMALL CAPITAL BARRED
            NormalizeCharacter = 66
        Case 199 'WITH CEDILLA, LATIN CAPITAL LETTER
            NormalizeCharacter = 67
        Case 571 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 67
        Case 7428 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 67
        Case 394 'WITH HOOK, LATIN CAPITAL LETTER
            NormalizeCharacter = 68
        Case 393 ', LATIN CAPITAL LETTER AFRICAN
            NormalizeCharacter = 68
        Case 7429 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 68
        Case 201 'WITH ACUTE, LATIN CAPITAL LETTER
            NormalizeCharacter = 69
        Case 202 'WITH CIRCUMFLEX, LATIN CAPITAL LETTER
            NormalizeCharacter = 69
        Case 203 'WITH DIAERESIS, LATIN CAPITAL LETTER
            NormalizeCharacter = 69
        Case 200 'WITH GRAVE, LATIN CAPITAL LETTER
            NormalizeCharacter = 69
        Case 582 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 69
        Case 400 ', LATIN CAPITAL LETTER OPEN
            NormalizeCharacter = 69
        Case 398 ', LATIN CAPITAL LETTER REVERSED
            NormalizeCharacter = 69
        Case 7431 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 69
        Case 403 'WITH HOOK, LATIN CAPITAL LETTER
            NormalizeCharacter = 71
        Case 667 'WITH HOOK, LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 71
        Case 610 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 71
        Case 668 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 72
        Case 205 'WITH ACUTE, LATIN CAPITAL LETTER
            NormalizeCharacter = 73
        Case 206 'WITH CIRCUMFLEX, LATIN CAPITAL LETTER
            NormalizeCharacter = 73
        Case 207 'WITH DIAERESIS, LATIN CAPITAL LETTER
            NormalizeCharacter = 73
        Case 304 'WITH DOT ABOVE, LATIN CAPITAL LETTER
            NormalizeCharacter = 73
        Case 204 'WITH GRAVE, LATIN CAPITAL LETTER
            NormalizeCharacter = 73
        Case 407 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 73
        Case 618 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 73
        Case 584 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 74
        Case 7434 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 74
        Case 7435 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 75
        Case 573 'WITH BAR, LATIN CAPITAL LETTER
            NormalizeCharacter = 76
        Case 7436 'WITH STROKE, LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 76
        Case 671 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 76
        Case 412 ', LATIN CAPITAL LETTER TURNED
            NormalizeCharacter = 77
        Case 7437 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 77
        Case 413 'WITH LEFT HOOK, LATIN CAPITAL LETTER
            NormalizeCharacter = 78
        Case 544 'WITH LONG RIGHT LEG, LATIN CAPITAL LETTER
            NormalizeCharacter = 78
        Case 209 'WITH TILDE, LATIN CAPITAL LETTER
            NormalizeCharacter = 78
        Case 628 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 78
        Case 7438 ', LATIN LETTER SMALL CAPITAL REVERSED
            NormalizeCharacter = 78
        Case 211 'WITH ACUTE, LATIN CAPITAL LETTER
            NormalizeCharacter = 79
        Case 212 'WITH CIRCUMFLEX, LATIN CAPITAL LETTER
            NormalizeCharacter = 79
        Case 214 'WITH DIAERESIS, LATIN CAPITAL LETTER
            NormalizeCharacter = 79
        Case 210 'WITH GRAVE, LATIN CAPITAL LETTER
            NormalizeCharacter = 79
        Case 415 'WITH MIDDLE TILDE, LATIN CAPITAL LETTER
            NormalizeCharacter = 79
        Case 216 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 79
        Case 213 'WITH TILDE, LATIN CAPITAL LETTER
            NormalizeCharacter = 79
        Case 390 ', LATIN CAPITAL LETTER OPEN
            NormalizeCharacter = 79
        Case 7439 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 79
        Case 7440 ', LATIN LETTER SMALL CAPITAL OPEN
            NormalizeCharacter = 79
        Case 7448 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 80
        Case 586 'WITH HOOK TAIL, LATIN CAPITAL LETTER SMALL
            NormalizeCharacter = 81
        Case 588 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 82
        Case 640 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 82
        Case 641 ', LATIN LETTER SMALL CAPITAL INVERTED
            NormalizeCharacter = 82
        Case 7449 ', LATIN LETTER SMALL CAPITAL REVERSED
            NormalizeCharacter = 82
        Case 7450 ', LATIN LETTER SMALL CAPITAL TURNED
            NormalizeCharacter = 82
        Case 574 'WITH DIAGONAL STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 84
        Case 430 'WITH RETROFLEX HOOK, LATIN CAPITAL LETTER
            NormalizeCharacter = 84
        Case 7451 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 84
        Case 580 'BAR, LATIN CAPITAL LETTER
            NormalizeCharacter = 85
        Case 218 'WITH ACUTE, LATIN CAPITAL LETTER
            NormalizeCharacter = 85
        Case 219 'WITH CIRCUMFLEX, LATIN CAPITAL LETTER
            NormalizeCharacter = 85
        Case 220 'WITH DIAERESIS, LATIN CAPITAL LETTER
            NormalizeCharacter = 85
        Case 217 'WITH GRAVE, LATIN CAPITAL LETTER
            NormalizeCharacter = 85
        Case 7452 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 85
        Case 434 'WITH HOOK, LATIN CAPITAL LETTER
            NormalizeCharacter = 86
        Case 581 ', LATIN CAPITAL LETTER TURNED
            NormalizeCharacter = 86
        Case 7456 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 86
        Case 7457 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 87
        Case 221 'WITH ACUTE, LATIN CAPITAL LETTER
            NormalizeCharacter = 89
        Case 376 'WITH DIAERESIS, LATIN CAPITAL LETTER
            NormalizeCharacter = 89
        Case 590 'WITH STROKE, LATIN CAPITAL LETTER
            NormalizeCharacter = 89
        Case 655 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 89
        Case 7458 ', LATIN LETTER SMALL CAPITAL
            NormalizeCharacter = 90
        Case Else
            NormalizeCharacter = 0
    End Select

End Function

' ---------------------------------------------------------------------------------------------------------------------
' --- Private methods -------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Max of two integers.")
Private Function Max(ByVal x As Long, ByVal y As Long) As Long
Attribute Max.VB_Description = "Max of two integers."

    If x < y Then
        Max = y
    Else
        Max = x
    End If

End Function

'@Description("Initialization protection.")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Initialization protection."
    If Not Me Is FzfAlgorithm Then
        Err.Raise _
            Number:=445, _
            Source:=TypeName(Me) & ".Initialize", _
            Description:="Cannot create an instance of a namespace."
    End If
End Sub
