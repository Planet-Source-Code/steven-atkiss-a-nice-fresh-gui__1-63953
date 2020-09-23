Attribute VB_Name = "Module1"
Option Explicit
Public Type LabelText
    Tag As String 'The Numeric ID Of The Next Page Using A Label
    Caption As String 'The Labels Caption
    Top As Long 'Top Position Of Label
    Left As Long 'Left Position Of Label
    FontSize As Single '
    Fade As Single 'Rate In Which The Itam fades In
    Next As Boolean 'Display Next Button
    Previous As Boolean 'Display Previous Button
    NextTag As Single 'The Numeric ID Of The Next Page When Using Next Button
    NextFrom As Single 'Display The Next Page, Only Fading Controls From Index
    PrevFrom As Single 'Display The Last Page, Only Fading Controls From Index
    PrevTag As Single 'The Numeric ID Of The Previous Page When Using Previous Button
    Tip As String
End Type
            
            'Allowed Tag Information
            'Numeric For Page ID
            'Input - Display An Input Box When Clicked On
            'End - End The Program
            'NoSelect - Dont Underline Or Trigger Mouse Events
            'ComboTitle - Display A Combo Box Populated With Titles
            
Public Type LabelInfo
    LabelInfo() As LabelText
End Type

Public Labels() As LabelInfo

