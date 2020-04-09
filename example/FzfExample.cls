VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FzfExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("example")
'@PredeclaredId
Option Explicit

' Example to to showcase one of the things you could do with the library.
' In this example we have a table with the elements we want to filter, and a
' pattern to apply (also in a table). Finally there is a textbox in which the
' results are printed. They are printed according to their score and the
' characters used in the match is highlighted.

Public Sub RefreshOnChange(ByVal Target As Range)

    Dim ValuesTable As ListObject
    Set ValuesTable = FzfExampleSheet.ListObjects("ValuesTable")
    
    Dim PatternTable As ListObject
    Set PatternTable = FzfExampleSheet.ListObjects("PatternTable")
    
    ' If we aren't changing the items nor the pattern then we aren't changing
    ' anything that has an effect.
    If Application.Intersect(ValuesTable.DataBodyRange, Target) Is Nothing And _
       Application.Intersect(PatternTable.DataBodyRange, Target) Is Nothing Then Exit Sub
    
    ' First collect the values in the table
    Dim Values() As String
    ReDim Values(0 To ValuesTable.DataBodyRange.Rows.Count - 1) As String
    
    Dim Row As Long
    For Row = 0 To UBound(Values)
        Values(Row) = ValuesTable.DataBodyRange.Cells(Row + 1, 1)
    Next Row
    
    Dim Pattern As String
    Pattern = PatternTable.DataBodyRange.Cells(1, 1)
    
    ' Then sort and filter the items
    Dim Results() As String
    Results = FzfUtilities.SortAndFilterStrict(Values, Pattern)
    
    Dim TextBox As Shape
    Set TextBox = FzfExampleSheet.Shapes("ResultTextbox")
    
    If ArrayIsEmpty(Results) Then
        TextBox.TextFrame2.TextRange.Characters.Text = VBA.vbNullString
        Exit Sub
    End If
    
    ' Then create the text string we are going to fill in the text box
    Dim Index As Long
    For Index = LBound(Results) To UBound(Results)
    
        Dim Text As String
        If Index = LBound(Results) Then
            Text = Results(Index)
        Else
            Text = Text & VBA.vbNewLine & Results(Index)
        End If
        
    Next Index
    
    With TextBox.TextFrame2.TextRange.Characters
        .Text = Text
        .Font.Bold = msoFalse
        .Font.Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
    End With
    
    ' Finally highlight the character used for the match.
    ' This has to be a second step because it will reset the formatting every
    ' time we reset the contents.
    Dim Offset As Long: Offset = 0
    For Index = LBound(Results) To UBound(Results)
    
        Dim Match As FzfResult
        Set Match = FzfAlgorithm.FuzzyMatchV1(Results(Index), Pattern, True)
        
        Dim pIdx As Long
        For pIdx = 0 To UBound(Match.Positions)
        
            Dim charOffset As Long
            charOffset = Offset + Match.Positions(pIdx)
        
            With TextBox.TextFrame2.TextRange.Characters(charOffset, 1).Font
                .Bold = msoTrue
                .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent2
            End With
        
        Next pIdx
        
        Offset = Offset + VBA.Len(Results(Index)) + 1
        
    Next Index

End Sub

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
