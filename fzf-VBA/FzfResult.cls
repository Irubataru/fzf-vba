VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FzfResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Default instance factory; avoid altering default instance state."
'@Folder("fzf-vba")
'@PredeclaredId
'@ModuleDescription("Default instance factory; avoid altering default instance state.")
Option Explicit

' Structure: FzfResult
' --------------------
' Simple structure with the return variables for the fzf algorithm.

Private Type Members
    StartIndex As Long
    EndIndex As Long
    Score As Long
    Positions As Variant
    Text As String
End Type

Private this As Members

' ---------------------------------------------------------------------------------------------------------------------
' --- Factory method --------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

'@DefaultMember
'@Description("Factory method.")
Public Function Create( _
        ByVal StartIndex As Long, _
        ByVal EndIndex As Long, _
        ByVal Score As Long, _
        Optional ByVal Positions As Variant = Empty, _
        Optional ByVal Text As String = VBA.vbNullString) As FzfResult
Attribute Create.VB_Description = "Factory method."
Attribute Create.VB_UserMemId = 0

    With New FzfResult
    
        .StartIndex = StartIndex
        .EndIndex = EndIndex
        .Score = Score
        .Positions = Positions
        .Text = Text
        
        Set Create = .Self
        
    End With

End Function

Public Function Self() As FzfResult
    Set Self = Me
End Function

' ---------------------------------------------------------------------------------------------------------------------
' --- Properties ------------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------

' --- StartIndex
Public Property Get StartIndex() As Long
    StartIndex = this.StartIndex
End Property

Public Property Let StartIndex(ByVal Value As Long)
    this.StartIndex = Value
End Property

' --- EndIndex
Public Property Get EndIndex() As Long
    EndIndex = this.EndIndex
End Property

Public Property Let EndIndex(ByVal Value As Long)
    this.EndIndex = Value
End Property

' --- Score
Public Property Get Score() As Long
    Score = this.Score
End Property

Public Property Let Score(ByVal Value As Long)
    this.Score = Value
End Property

' --- Positions
Public Property Get Positions() As Variant
    Positions = this.Positions
End Property

Public Property Let Positions(ByVal Value As Variant)
    this.Positions = Value
End Property

' --- Text
Public Property Get Text() As String
    Text = this.Text
End Property

Public Property Let Text(ByVal Value As String)
    this.Text = Value
End Property
