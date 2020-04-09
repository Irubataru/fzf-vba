VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FzfUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Namespace class; do not create an instance."
'@Folder("fzf-vba")
'@PredeclaredId
'@ModuleDescription("Namespace class; do not create an instance.")
Option Explicit

' Namespace: FzfUtilities
' -----------------------
' Utilities to manipulate arrays of strings that should be fuzzy matched,
' filtered, and sorted based on some criteria.

Private Type ScoreTextPair
    Score As Long
    Text As String
End Type
' ---------------------------------------------------------------------------------------------------------------------
' --- Public methods --------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Score an array of texts, filter them based on a score threshold and sort them from highest to lowest score.")
Public Function SortAndFilter( _
        ByVal Texts As Variant, _
        ByVal Pattern As String, _
        Optional ByVal ScoreThreshold As Long = 0, _
        Optional OnlyTopN As Long = -1) As String()
Attribute SortAndFilter.VB_Description = "Score an array of texts, filter them based on a score threshold and sort them from highest to lowest score."

    ' If the user requests 0 results there is nothing to do
    If OnlyTopN = 0 Then Exit Function

    SortAndFilter = SortAndFilterStrict(ConvertTextArgument(Texts), Pattern, ScoreThreshold, OnlyTopN)

End Function

'@Description("SortAndFilter but with stricter typing.")
Public Function SortAndFilterStrict( _
        ByRef Texts() As String, _
        ByVal Pattern As String, _
        Optional ByVal ScoreThreshold As Long = 0, _
        Optional OnlyTopN As Long = -1) As String()
Attribute SortAndFilterStrict.VB_Description = "SortAndFilter but with stricter typing."

    ' If the user requests 0 results there is nothing to do
    If OnlyTopN = 0 Then Exit Function

    Dim Results() As ScoreTextPair
    ReDim Results(0 To UBound(Texts) - LBound(Texts)) As ScoreTextPair
    
    Dim Count As Long: Count = 0
    
    ' Step 1: Loop through, calculate score and filter out the ones with a too low score
    Dim Index As Long
    For Index = LBound(Texts) To UBound(Texts)
    
        Dim Result As FzfResult
        Set Result = FzfAlgorithm.FuzzyMatchV1(Texts(Index), Pattern)
        
        If Result.Score > ScoreThreshold Then
            Results(Count).Score = Result.Score
            Results(Count).Text = Texts(Index)
            Count = Count + 1
        End If

    Next Index
    
    If Count = 0 Then Exit Function
    
    ReDim Preserve Results(0 To Count - 1) As ScoreTextPair
    
    ' Step 2: Sort the scores and return the texts
    Dim ReturnCount As Long
    If OnlyTopN > 0 And OnlyTopN < Count Then
        ReturnCount = OnlyTopN
    Else
        ReturnCount = Count
    End If
    
    Dim ReturnValues() As String
    ReDim ReturnValues(0 To ReturnCount - 1) As String
    
    ' If we only want the result with the highest score we do not have to sort
    If OnlyTopN = 1 Then
        
        Dim HighestScore As Long: HighestScore = Results(0).Score
        ReturnValues(0) = Results(0).Text
        
        For Index = 1 To Count - 1
            
            If Results(Index).Score > HighestScore Then
                HighestScore = Results(Index).Score
                ReturnValues(0) = Results(Index).Text
            End If
            
        Next Index
        
    ' Otherwise sort the array as usual and extract the top N results
    Else
    
        Results = SortByScore(Results)
        
        For Index = 0 To ReturnCount - 1
            ReturnValues(Index) = Results(Index).Text
        Next Index
    
    End If
    
    SortAndFilterStrict = ReturnValues

End Function

' ---------------------------------------------------------------------------------------------------------------------
' --- Private methods -------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

' Merge sort methods
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Sort the results based on score.")
Private Function SortByScore(ByRef Items() As ScoreTextPair) As ScoreTextPair()
Attribute SortByScore.VB_Description = "Sort the results."

    ' This is an implementation of mergesort that manipulates an index array
    ' instead of the array itself. At the end it reorders the items based on
    ' the index array.
    
    ' If there is only 1 element it is already sorted
    If UBound(Items) = 0 Then
        SortByScore = Items
        Exit Function
    End If

    ' Step 0: Create the index array
    Dim Indices() As Long
    ReDim Indices(0 To UBound(Items)) As Long
    
    Dim Index As Long
    For Index = 0 To UBound(Items)
        Indices(Index) = Index
    Next Index

    ' Step 1: Merge sort algorithm
    Dim WorkArray() As Long
    ReDim WorkArray(0 To UBound(Items)) As Long
    
    MergeSort_Copy Indices, WorkArray
    MergeSort_TopDownSplitMerge WorkArray, Indices, 0, UBound(Items), Items
    
    ' Step 2: Reorder Items based on the index positions
    Dim Result() As ScoreTextPair
    ReDim Result(0 To UBound(Items)) As ScoreTextPair
    
    For Index = 0 To UBound(Items)
        Result(Index) = Items(Indices(Index))
    Next Index
    
    SortByScore = Result

End Function

Private Sub MergeSort_TopDownSplitMerge( _
        ByRef WorkArrayA() As Long, _
        ByRef WorkArrayB() As Long, _
        ByVal IndexBegin As Long, _
        ByVal IndexEnd As Long, _
        ByRef Items() As ScoreTextPair)

    ' If the partition only has 1 item it is sorted
    If IndexEnd <= IndexBegin Then Exit Sub
    
    Dim IndexMiddle As Long: IndexMiddle = (IndexBegin + IndexEnd) \ 2
    MergeSort_TopDownSplitMerge WorkArrayB, WorkArrayA, IndexBegin, IndexMiddle, Items
    MergeSort_TopDownSplitMerge WorkArrayB, WorkArrayA, IndexMiddle + 1, IndexEnd, Items
    MergeSort_TopDownMerge WorkArrayA, WorkArrayB, IndexBegin, IndexMiddle + 1, IndexEnd, Items
    
End Sub

Private Sub MergeSort_TopDownMerge( _
        ByRef Indices() As Long, _
        ByRef Destination() As Long, _
        ByVal IndexBegin As Long, _
        ByVal IndexMiddle As Long, _
        ByVal IndexEnd As Long, _
        ByRef Items() As ScoreTextPair)

    Dim i As Long: i = IndexBegin
    Dim j As Long: j = IndexMiddle
    
    Dim k As Long
    For k = IndexBegin To IndexEnd
    
        ' Due to VBA's lack of short-circuiting we need to do this check in stages
        Dim NextFromLower As Boolean
        NextFromLower = i < IndexMiddle
        If NextFromLower Then
            If j <= IndexEnd Then
                NextFromLower = Items(Indices(j)).Score < Items(Indices(i)).Score
            End If
        End If
    
        If NextFromLower Then
            Destination(k) = Indices(i)
            i = i + 1
        Else
            Destination(k) = Indices(j)
            j = j + 1
        End If
    Next k

End Sub

Private Sub MergeSort_Copy(ByRef Items() As Long, ByRef Destination() As Long)

    Dim i As Long
    For i = LBound(Items) To UBound(Items)
        Destination(i) = Items(i)
    Next i

End Sub

' Utility methods
' ---------------------------------------------------------------------------------------------------------------------

'@Description("Lenient conversion of the Texts argument of the SortAndFilter function.")
Private Function ConvertTextArgument(ByVal Texts As Variant) As String()
Attribute ConvertTextArgument.VB_Description = "Lenient conversion of the Texts argument of the SortAndFilter function."

    If Not VBA.VarType(Texts) = vbString And Not VBA.IsArray(Texts) Then
        Err.Raise _
            Number:=13, _
            Source:=VBA.TypeName(Me) & ".SortAndFilter", _
            Description:="Texts argument has to either be a string or an array of strings."
    End If
    
    Dim Result() As String
    If Not VBA.IsArray(Texts) Then
        Result = TextToArray(Texts)
    Else
    
        ReDim Result(0 To UBound(Texts) - LBound(Texts)) As String
        
        Dim Index As Long
        For Index = 0 To UBound(Result)
            Result(Index) = Texts(LBound(Texts) + Index)
        Next Index
        
    End If
    
    ConvertTextArgument = Result

End Function

'@Description("Convert an item to a 1-long array.")
Private Function TextToArray(ByVal Value As String) As String()
Attribute TextToArray.VB_Description = "Convert an item to a 1-long array."

    Dim Result() As String
    ReDim Result(0 To 0) As String
    
    Result(0) = Value
    TextToArray = Result

End Function

'@Description("Initialization protection.")
Private Sub Class_Initialize()
Attribute Class_Initialize.VB_Description = "Initialization protection."
    If Not Me Is FzfUtilities Then
        Err.Raise _
            Number:=445, _
            Source:=TypeName(Me) & ".Initialize", _
            Description:="Cannot create an instance of a namespace."
    End If
End Sub
