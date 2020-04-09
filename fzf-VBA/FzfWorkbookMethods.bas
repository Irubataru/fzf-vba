Attribute VB_Name = "FzfWorkbookMethods"
Attribute VB_Description = "Workbook methods for the fzf-vba library."
'@Folder("fzf-vba")
'@ModuleDescription("Workbook methods for the fzf-vba library.")

' Module (public): FzfWorkbookMethods
' -----------------------------------
' Functions that can be used in range formulas to do fuzzy matching.

' ---------------------------------------------------------------------------------------------------------------------
' --- Public methods --------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Fuzzy find on a string or range and return the results.")
Public Function Fzf( _
        ByVal Value As Variant, _
        ByVal Pattern As Variant, _
        Optional ByVal OnlyTopN As Long = 1, _
        Optional ByVal ScoreThreshold As Long = 0, _
        Optional ByVal TransposeResult As Boolean = False) As Variant
Attribute Fzf.VB_Description = "Runs a fuzzy find on the input text with a pattern."

    On Error GoTo CleanExitOnError

    ' First parse the input arguments
    Dim Texts() As String
    Texts = ParseStringArray(Value)
    Pattern = ParseString(Pattern)
    
    ' Then run it through a filter
    Dim FilteredTexts() As String
    FilteredTexts = FzfUtilities.SortAndFilterStrict(Texts, Pattern, ScoreThreshold:=ScoreThreshold, OnlyTopN:=OnlyTopN)
    
    If ArrayIsEmpty(FilteredTexts) Then Exit Function
    
    ' Finally construct the result
    If OnlyTopN = 1 Then
        Fzf = FilteredTexts(0)
    Else
        
        Dim ReturnValue() As Variant
        ReDim ReturnValue(1 To UBound(FilteredTexts) + 1, 1 To 1) As Variant
        
        Dim Index As Long
        For Index = 0 To UBound(FilteredTexts)
            ReturnValue(Index + 1, 1) = FilteredTexts(Index)
        Next Index
        
        If TransposeResult Then
            Fzf = Transpose(ReturnValue)
        Else
            Fzf = ReturnValue
        End If
        
    End If
    
Finally:
    Exit Function
    
CleanExitOnError:
    If Err.Number = 13 Then
        Fzf = "#VALUE!"
    Else
        Fzf = "#N/A?"
    End If
    
    Resume Finally

End Function

'@Description("Return the fzf score searching for pattern in value.")
Public Function FzfScore(ByVal Value As Variant, ByVal Pattern As Variant) As Variant
Attribute FzfScore.VB_Description = "Return the fzf score searching for pattern in value."

    On Error GoTo CleanExitOnError
    FzfScore = FzfMatchExcelValues(Value, Pattern)
    
Finally:
    Exit Function
    
CleanExitOnError:
    If Err.Number = 13 Then
        FzfScore = "#VALUE!"
    Else
        FzfScore = "#N/A?"
    End If
    
    Resume Finally

End Function

'@Description("Return the fzf start index searching for pattern in value.")
Public Function FzfStartIndex(ByVal Value As Variant, ByVal Pattern As Variant) As Variant
Attribute FzfStartIndex.VB_Description = "Return the fzf start index searching for pattern in value."

    On Error GoTo CleanExitOnError
    FzfStartIndex = FzfMatchExcelValues(Value, Pattern).StartIndex + 1
    
Finally:
    Exit Function
    
CleanExitOnError:
    If Err.Number = 13 Then
        FzfStartIndex = "#VALUE!"
    Else
        FzfStartIndex = "#N/A?"
    End If
    
    Resume Finally

End Function

'@Description("Return the fzf end index searching for pattern in value.")
Public Function FzfEndIndex(ByVal Value As Variant, ByVal Pattern As Variant) As Variant
Attribute FzfEndIndex.VB_Description = "Return the fzf end index searching for pattern in value."

    On Error GoTo CleanExitOnError
    FzfEndIndex = FzfMatchExcelValues(Value, Pattern).EndIndex + 1
    
Finally:
    Exit Function
    
CleanExitOnError:
    If Err.Number = 13 Then
        FzfEndIndex = "#VALUE!"
    Else
        FzfEndIndex = "#N/A?"
    End If
    
    Resume Finally

End Function

'@Description("Normalize latin script letters.")
Public Function FzfNormalize(ByVal Value As Variant) As String
Attribute FzfNormalize.VB_Description = "Normalize latin script letters."

    On Error GoTo CleanExitOnError
    FzfNormalize = FzfAlgorithm.Normalize(ParseString(Value))
    
Finally:
    Exit Function
    
CleanExitOnError:
    If Err.Number = 13 Then
        FzfNormalize = "#VALUE!"
    Else
        FzfNormalize = "#N/A?"
    End If
    
    Resume Finally

End Function

' ---------------------------------------------------------------------------------------------------------------------
' --- Private methods -------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Parse an array of strings from a variant that can also be a Range.")
Private Function ParseStringArray(ByVal Value As Variant) As String()
Attribute ParseStringArray.VB_Description = "Parse an array of strings from a variant that can also be a Range."

    If VBA.IsObject(Value) Then
    
        If VBA.TypeName(Value) <> "Range" Then
            Err.Raise _
                Number:=13, _
                Source:="FzfWorkbookMethods.ParseStringArray", _
                Description:="Unknown argument type """ & VBA.TypeName(Value) & """."
        End If
        
        Dim Cells As Range
        Set Cells = Value
        
        Value = Cells.Value

    End If
    
    Dim Result() As String
        
    If VBA.IsArray(Value) Then
    
        ' Parse a matrix
        If IsMatrix(Value) Then
        
            Dim Length As Long
            Length = (UBound(Value, 1) - LBound(Value, 1) + 1) * (UBound(Value, 2) - LBound(Value, 2) + 1)
        
            ReDim Result(0 To Length - 1) As String
            
            Dim SuperIndex As Long: SuperIndex = 0
        
            Dim i As Long
            For i = LBound(Value, 1) To UBound(Value, 1)
            
                Dim j As Long
                For j = LBound(Value, 2) To UBound(Value, 2)
                    Result(SuperIndex) = Value(i, j)
                    SuperIndex = SuperIndex + 1
                Next j
            
            Next i
        
        ' Parse an array
        Else
        
            Length = (UBound(Value) - LBound(Value) + 1)
        
            ReDim Result(0 To Length - 1) As String
        
            For i = LBound(Value) To UBound(Value)
                Result(i - LBound(Value)) = Value(i)
            Next i
        
        End If
    
    ' Parse a value
    Else
        ReDim Result(0 To 0) As String
        Result(0) = Value
    End If
    
    ParseStringArray = Result

End Function

'@Description("Parse a string from a variant that can also be a Range.")
Private Function ParseString(ByVal Value As Variant) As String
Attribute ParseString.VB_Description = "Parse a string from a variant that can also be a Range."

    If VBA.IsObject(Value) Then
    
        If VBA.TypeName(Value) <> "Range" Then
            Err.Raise _
                Number:=13, _
                Source:="FzfWorkbookMethods.ParseString", _
                Description:="Unknown argument type """ & VBA.TypeName(Value) & """."
        End If
        
        Dim Cells As Range
        Set Cells = Value
        
        Value = Cells.Value

    End If
        
    If VBA.IsArray(Value) Then
        Err.Raise _
            Number:=13, _
            Source:="FzfWorkbookMethods.ParseString", _
            Description:="Expected a single valued argument, got an array."
    End If
    
    ParseString = Value

End Function

'@Description("Parse text strings from Excel formula variants and call the fuzzy algorithm.")
Private Function FzfMatchExcelValues(ByVal Value As Variant, ByVal Pattern As Variant) As FzfResult
Attribute FzfMatchExcelValues.VB_Description = "Parse text strings from Excel formula variants and call the fuzzy algorithm."
    
    Set FzfMatchExcelValues = FzfAlgorithm.FuzzyMatchV1(ParseString(Value), ParseString(Pattern))
    
End Function

'@Description("Check if an array is empty.")
Private Function ArrayIsEmpty(ByVal Items As Variant) As Boolean
Attribute ArrayIsEmpty.VB_Description = "Check if an array is empty."

    On Error GoTo CleanExit

    Dim Dummy As Long
    Dummy = LBound(Items)
    
    ArrayIsEmpty = False

Finally:
    Exit Function

CleanExit:
    ArrayIsEmpty = True
    Resume Finally

End Function

'@Description("Check if a variant is a matrix.")
Private Function IsMatrix(ByVal Value As Variant) As Boolean
Attribute IsMatrix.VB_Description = "Check if a variant is a matrix."

    ' Didn't find a better solution than to check it by error handling, really
    ' wish there was a better way, but this was all I could find.

    On Error GoTo CleanExit
    
    Dim Dummy As Long
    Dummy = LBound(Value, 2)
    
    IsMatrix = True
    
Finally:
    Exit Function
    
CleanExit:
    IsMatrix = False
    Resume Finally

End Function

'@Description("Transpose a matrix.")
Private Function Transpose(ByVal Matrix As Variant) As Variant()
Attribute Transpose.VB_Description = "Transpose a matrix."

    Dim Result() As Variant
    ReDim Result(LBound(Matrix, 2) To UBound(Matrix, 2), LBound(Matrix, 1) To UBound(Matrix, 1))
    
    Dim i As Long
    For i = LBound(Matrix, 1) To UBound(Matrix, 1)
    
        Dim j As Long
        For j = LBound(Matrix, 2) To UBound(Matrix, 2)
            Result(j, i) = Matrix(i, j)
        Next j
    
    Next i

    Transpose = Result

End Function
