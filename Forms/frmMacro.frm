VERSION 5.00
Begin VB.Form frmMacro 
   Caption         =   "frmMacro"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCheck 
      Caption         =   "&Check"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox txtCode 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'The API
 'For other Versions of VB change vba6.dll to version eg vba5.dll
 Private Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, _
    ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long

'The Powerfull Function
Public Function ExecuteLine(psCode As String, Optional pbCheckOnly As Boolean) As Boolean
    'User the strPointer to the Code you want to execute for pStringToExec
    'The pbCheckOnly Flag will execute code on TRUE and Validate it on FALSE
    'The Function EbExecuteLine will return TRUE on success and FALSE on Failure
   ExecuteLine = EbExecuteLine(StrPtr(psCode), 0&, 0&, Abs(pbCheckOnly)) = 0
End Function


Private Sub cmdCheck_Click()
    'Set the pbCheckOnly as TRUE to Validate the Code
    'The MsgBox will return the Validation Status
    MsgBox ExecuteLine(txtCode.Text, True)
End Sub

Private Sub cmdRun_Click()
    'Set the pbCheckOnly to False to execute the code
    ExecuteLine txtCode.Text, False
End Sub

Private Sub Form_Load()
    'Just some sample code to test
    'You could actually try anything here
    txtCode.Text = "frmMacro.BackColor = vbred:MsgBox " & Chr$(34) & "Just think what you could do next!" & Chr$(34)
End Sub
