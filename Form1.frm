VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ellipses Demo"
   ClientHeight    =   825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   825
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOriginal 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Text            =   "C:\Program Files\Microsoft Visual Studio\VB98\Projects\Ellipses\Project1.vbp"
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtTruncated 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "C:\Program Files\Microsoft Visual Studio\VB98\Projects\Ellipses\Project1.vbp"
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "TruncatedText:"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Original Text:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Error Resume Next

    txtOriginal.Text = App.Path & "\" & App.EXEName & ".vbp"
    
End Sub
Private Sub Form_Resize()
On Error Resume Next

    txtOriginal.Width = Me.ScaleWidth - txtOriginal.Left
    txtTruncated.Width = Me.ScaleWidth - txtTruncated.Left
    
    Call txtOriginal_Change

End Sub

Private Sub txtOriginal_Change()
On Error Resume Next
    Const cTBWidth As Single = 0.97 'This is to adjust for the textbox boardwidth.
    Const cCBWidth As Integer = 2400 'This is to adjust for then Controlbox width and icon on the titlebar.
    Dim lhDc As Long 'Device context handle
    Dim sOriginal As String 'Original Text String
    Dim lTBWidth As Long 'Width of Control.
    Dim lCBWidth As Long 'Scalewidth - Controlbox width.
    Dim sCaption As String 'Caption Bar Text String.
    
    'Get the form's hDC property.
    'You can also use a picturebox's hDC property.
    'As long as this is a valid handle, it doesn't matter which.
    lhDc = Me.hdc

    'Get additional params.
    sOriginal = txtOriginal.Text
    lTBWidth = txtTruncated.Width * cTBWidth
    lCBWidth = Me.Width - cCBWidth

    'Replace with truncated text.
    txtTruncated.Text = CEllipses(lhDc, lTBWidth, sOriginal, DT_PATH_ELLIPSIS)

    'Place the same path on the titlebar.
    sCaption = "Ellipses Demo: " & sOriginal
    Me.Caption = CEllipses(lhDc, lCBWidth, sCaption, DT_PATH_ELLIPSIS)

End Sub
