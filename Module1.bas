Attribute VB_Name = "Module1"
Option Explicit
'****************************************************************************************
'Ellipsis Text
'Dependencies: None
'Author(s): Matthew Hood Email: DragonWeyrDev@Yahoo.com
'****************************************************************************************
'****************************************************************************************
'Private Data Types
'****************************************************************************************
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'****************************************************************************************
'Public Constants/Variables
'****************************************************************************************
Public Const DT_PATH_ELLIPSIS As Long = &H4000
Public Const DT_END_ELLIPSIS As Long = &H8000
Public Const DT_WORD_ELLIPSIS As Long = &H40000
'****************************************************************************************
'API Declarations
'****************************************************************************************
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'****************************************************************************************
'Public Routines
'****************************************************************************************
'Truncates long strings with ellipses.
Public Function CEllipses(hdc As Long, Width As Long, ByVal Text As String, EllipsesType As Long) As String
On Error Resume Next
    Const DT_CALCRECT = &H400
    Const DT_MODIFYSTRING = &H10000
    Dim r As RECT

    r.Right = Width / Screen.TwipsPerPixelX

    Call DrawText(hdc, Text, -1, r, DT_CALCRECT Or DT_MODIFYSTRING Or EllipsesType)
   
   CEllipses = Text

End Function
