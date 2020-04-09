Attribute VB_Name = "TestFzf"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests.fzf-vba")

Private Const ScoreMatch As Long = 16
Private Const ScoreGapStart As Long = -3
Private Const ScoreGapExtention As Long = -1
Private Const BonusBoundary As Long = ScoreMatch \ 2
Private Const BonusNonWord As Long = ScoreMatch \ 2
Private Const BonusCamel123 As Long = BonusBoundary + ScoreGapExtention
Private Const BonusConsecutive  As Long = -(ScoreGapStart + ScoreGapExtention)
Private Const BonusFirstCharMultiplier = 2

Private Assert As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
End Sub

'@TestMethod("fzf-vba")
Private Sub TestFuzzyMatch()
    On Error GoTo TestFail
    
    AssertMatch _
        2, 8, 3 * ScoreMatch + BonusCamel123 + ScoreGapStart + 3 * ScoreGapExtention, _
        "fooBarbaz1", "oBZ"
        
    AssertMatch _
        0, 8, ScoreMatch * 3 + BonusBoundary * BonusFirstCharMultiplier + BonusBoundary * 2 + 2 * ScoreGapStart + 4 * ScoreGapExtention, _
        "foo bar baz", "fbb"
        
    AssertMatch _
        9, 12, ScoreMatch * 4 + BonusCamel123 + BonusConsecutive * 2, _
        "/AutomatorDocument.icns", "rdoc"
        
    AssertMatch _
        6, 9, ScoreMatch * 4 + BonusBoundary * BonusFirstCharMultiplier + BonusBoundary * 3, _
        "/man1/zshcompctl.1", "zshc"
        
    AssertMatch _
        8, 12, ScoreMatch * 4 + BonusBoundary * BonusFirstCharMultiplier + BonusBoundary * 3 + ScoreGapStart, _
        "/.oh-my-zsh/cache", "zshc"
        
    AssertMatch _
        3, 9, ScoreMatch * 5 + BonusConsecutive * 3 + ScoreGapStart + ScoreGapExtention, _
        "ab0123 456", "12356"
        
    AssertMatch _
        3, 9, ScoreMatch * 5 + BonusCamel123 * BonusFirstCharMultiplier + BonusCamel123 * 2 + BonusConsecutive + ScoreGapStart + ScoreGapExtention, _
        "abc123 456", "12356"
        
    AssertMatch _
        0, 8, ScoreMatch * 3 + BonusBoundary * BonusFirstCharMultiplier + BonusBoundary * 2 + 2 * ScoreGapStart + 4 * ScoreGapExtention, _
        "foo/bar/baz", "fbb"
        
    AssertMatch _
        0, 6, ScoreMatch * 3 + BonusBoundary * BonusFirstCharMultiplier + BonusCamel123 * 2 + 2 * ScoreGapStart + 2 * ScoreGapExtention, _
        "fooBarBaz", "fbb"
        
    AssertMatch _
        0, 7, ScoreMatch * 3 + BonusBoundary * BonusFirstCharMultiplier + BonusBoundary + ScoreGapStart * 2 + ScoreGapExtention * 3, _
        "foo barbaz", "fbb"
        
    AssertMatch _
        0, 3, ScoreMatch * 4 + BonusBoundary * BonusFirstCharMultiplier + BonusBoundary * 3, _
        "fooBar Baz", "foob"
        
    AssertMatch _
        1, 5, ScoreMatch * 5 + BonusCamel123 * BonusFirstCharMultiplier + BonusCamel123 * 2 + BonusNonWord + BonusBoundary, _
        "xFoo-Bar Baz", "foo-b"
        
    AssertMatch _
        2, 5, ScoreMatch * 4 + BonusBoundary * 3, _
        "foo-bar", "o-ba"
    
    ' Edge case for repeating characters
    AssertMatch _
        0, 3, ScoreMatch * 2 + BonusBoundary * BonusFirstCharMultiplier + ScoreGapStart + ScoreGapExtention, _
        "barbar", "bb"
        
    ' Single letter matches
    AssertMatch _
        0, 0, ScoreMatch + BonusBoundary * BonusFirstCharMultiplier, _
        "foobar", "f"
    
    AssertMatch _
        3, 3, ScoreMatch, _
        "foobar", "b"

    ' Non-match
    AssertMatch -1, -1, 0, "fooBarbaz", "fzb"
    AssertMatch -1, -1, 0, "Foo Bar Baz", "bbb"
    AssertMatch -1, -1, 0, "fooBarbaz", "fooBarbazz"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("fzf-vba")
Private Sub TestEmptyText()
    On Error GoTo TestFail
    
    AssertMatch -1, -1, 0, "", "fb"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("fzf-vba")
Private Sub TestEmptyPattern()
    On Error GoTo TestFail
    
    AssertMatch 0, 0, 0, "foobar", ""

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("fzf-vba")
Private Sub TestNormalize()
    On Error GoTo TestFail
    
    AssertMatch 0, 6, 89, "Só Danço Samba", "sodc"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub AssertMatch(ByVal StartIndex As Long, ByVal EndIndex As Long, ByVal Score As Long, ByVal Text As String, ByVal Pattern As String)

    Dim Expected As FzfResult
    Set Expected = FzfResult(StartIndex, EndIndex, Score, Empty)

    Dim Result As FzfResult
    Set Result = FzfAlgorithm.FuzzyMatchV1(Text, Pattern)

    If (Expected.StartIndex <> Result.StartIndex) Or _
       (Expected.EndIndex <> Result.EndIndex) Or _
       (Expected.Score <> Result.Score) Then
       
        Assert.Fail _
            "Expected: " & SerializeResult(Expected) & ", " & _
            "Actual: " & SerializeResult(Result) & ", " & _
            "matching """ & Pattern & """ in """ & Text & """."
       
    Else
        Assert.Succeed
    End If

End Sub

Private Function SerializeResult(ByVal Result As FzfResult) As String
    
    SerializeResult = "{" & Result.StartIndex & ", " & Result.EndIndex & ", " & Result.Score & "}"
    
End Function
