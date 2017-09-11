VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "RGSS msgbox Generator"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4365
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   4365
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame2 
      Caption         =   "조건문"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   4650
      Width           =   4095
      Begin VB.OptionButton optNormal 
         Caption         =   "미사용"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optIf 
         Caption         =   "if 문"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton optCase 
         Caption         =   "case 문"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "스크립트를 클립보드에 복사"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   4095
   End
   Begin VB.TextBox txtScript 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   6120
      Width           =   4095
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "테스트"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   1935
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "스크립트 생성"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   11
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "옵션"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   4095
      Begin MSComctlLib.ImageList ImageList 
         Left            =   3360
         Top             =   2160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":0C5E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":18B0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":2502
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMain.frx":3154
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageCombo imgCmbIcon 
         Height          =   570
         Left            =   1320
         TabIndex        =   21
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1005
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Text            =   "ImageCombo1"
         ImageList       =   "ImageList"
      End
      Begin VB.ComboBox cmbFocus 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         Style           =   2  '드롭다운 목록
         TabIndex        =   3
         Top             =   1380
         Width           =   2655
      End
      Begin VB.ComboBox cmbButton 
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         Style           =   2  '드롭다운 목록
         TabIndex        =   2
         Top             =   330
         Width           =   2655
      End
      Begin VB.CheckBox chkRight 
         Caption         =   "글자 오른쪽 정렬"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox chkR2L 
         Caption         =   "좌우 뒤집기"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CheckBox chkTopMost 
         Caption         =   "항상 위로 표시"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "버튼 포커스"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "아이콘 종류"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   17
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "버튼 종류"
         BeginProperty Font 
            Name            =   "맑은 고딕"
            Size            =   9
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.TextBox txtContent 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  '수직
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "제목"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "내용"
      BeginProperty Font 
         Name            =   "맑은 고딕"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''
'title  RGSS MessageBox Code Generator

'date   12/22/2016
'author jubin-park
'.....................................

Option Explicit
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wtype As Long) As Long

Dim dataButton As Integer
Dim dataIcon As Integer
Dim dataButtonFocus As Integer
Dim dataRight As Long
Dim dataR2L As Long
Dim dataTopMost As Long

Dim dataRGSSButton As String
Dim dataRGSSIcon As String
Dim dataRGSSButtonFocus As String
Dim dataRGSSRight As String
Dim dataRGSSR2L As String
Dim dataRGSSTopMost As String

Private Sub cmdCopy_Click()
    If Len(txtScript.Text) > 0 Then
        Clipboard.Clear
        Clipboard.SetText (txtScript.Text)
        MsgBox "클립보드에 복사하였습니다." & Chr(13) & "스크립트 에디터에 Ctrl+V 로 붙여넣으세요.", vbInformation, Me.Caption
    End If
End Sub

Private Function createNormal() As String

    Dim str, strContent As String

    strContent = txtContent.Text
    strContent = Replace(strContent, """", "\""")
    strContent = Replace(strContent, "'", "\'")
    strContent = Replace(strContent, Chr(13), "\n")
    strContent = Replace(strContent, Chr(10), "\n")
    
    str = "msgbox(""" & strContent & """)"
    
    If (dataButton Or dataIcon Or dataButtonFocus Or dataRight Or dataR2L Or dataTopMost) > 0 Then
        str = str + " { ["
        If dataButton > 0 And Not dataRGSSButton = "" Then
            str = str + dataRGSSButton + " | "
        End If
        If dataIcon > 0 And Not dataRGSSIcon = "" Then
            str = str + dataRGSSIcon + " | "
        End If
        If dataButtonFocus > 0 And Not dataRGSSButtonFocus = "" Then
            str = str + dataRGSSButtonFocus + " | "
        End If
        If dataRight > 0 And Not dataRGSSRight = "" Then
            str = str + dataRGSSRight + " | "
        End If
        If dataR2L > 0 And Not dataRGSSR2L = "" Then
            str = str + dataRGSSR2L + " | "
        End If
        If dataTopMost > 0 And Not dataRGSSTopMost = "" Then
            str = str + dataRGSSTopMost + " | "
        End If
        If Len(txtTitle.Text) > 0 Then
            str = str + ", """ + txtTitle.Text + """"
        End If
        str = str + "] }"
        If Len(txtTitle.Text) > 0 Then
            str = Replace(str, " | , """ + txtTitle.Text + """] }", ", """ + txtTitle.Text + """] }")
        Else
            str = Replace(str, " | ] }", "] }")
        End If
    Else
        If Len(txtTitle.Text) > 0 Then
            str = str + " { [""" + txtTitle.Text + """] }"
        End If
    End If
    
    createNormal = str

End Function

Private Function createIf() As String
    
    Dim str As String
    
    str = "if " + createNormal
    
    Select Case dataButton
    Case 0 'MB::IDOK = 1
        str = str + " == " + "MB::IDOK"
        str = str + Chr(10) + "  # [확인]"
        str = str + Chr(10) + "  "
    Case 1 'MB::IDOK = 1        MB::IDCANCEL = 2
        str = str + " == " + "MB::IDOK"
        str = str + Chr(10) + "  # [확인]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDCANCEL"
        str = str + Chr(10) + "  # [취소]"
        str = str + Chr(10) + "  "
    Case 2 'MB::IDABORT = 3     MB::IDRETRY = 4     MB::IDIGNORE = 5
        str = str + " == " + "MB::IDABORT"
        str = str + Chr(10) + "  # [중단(A)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDRETRY"
        str = str + Chr(10) + "  # [다시 시도(R)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDIGNORE"
        str = str + Chr(10) + "  # [무시(I)]"
        str = str + Chr(10) + "  "
    Case 3 'MB::IDYES = 6       MB::IDNO = 7        MB::IDCANCEL = 2
        str = str + " == " + "MB::IDYES"
        str = str + Chr(10) + "  # [예(Y)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDNO"
        str = str + Chr(10) + "  # [아니오(N)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDCANCEL"
        str = str + Chr(10) + "  # [취소]"
        str = str + Chr(10) + "  "
    Case 4 'MB::IDYES = 6       MB::IDNO = 7
        str = str + " == " + "MB::IDYES"
        str = str + Chr(10) + "  # [예]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDNO"
        str = str + Chr(10) + "  # [아니오]"
        str = str + Chr(10) + "  "
    Case 5 'MB::IDRETRY = 4     MB::IDCANCEL = 2
        str = str + " == " + "MB::IDRETRY"
        str = str + Chr(10) + "  # [다시 시도(R)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDCANCEL"
        str = str + Chr(10) + "  # [취소]"
        str = str + Chr(10) + "  "
    Case 6 'MB::IDCANCEL = 2    MB::IDTRYAGAIN = 10     MB::IDCONTINUE = 11
        str = str + " == " + "MB::IDCANCEL"
        str = str + Chr(10) + "  # [취소]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDTRYAGAIN"
        str = str + Chr(10) + "  # [다시 시도(T)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::IDCONTINUE"
        str = str + Chr(10) + "  # [계속(C)]"
        str = str + Chr(10) + "  "
    Case &H4000 'MB::IDOK = 1  MB::HELP = 16384
        str = str + " == " + "MB::IDOK"
        str = str + Chr(10) + "  # [확인]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "elsif " + "MB::HELP"
        str = str + Chr(10) + "  # [도움말]"
        str = str + Chr(10) + "  "
    End Select
    
    str = str + Chr(10) + "end"
    
    createIf = str
    
End Function


Private Function createCase() As String
    
    Dim str As String
    
    str = "case " + createNormal + Chr(10)
    
    Select Case dataButton
    Case 0 'MB::IDOK = 1
        str = str + "when " + "MB::IDOK"
        str = str + Chr(10) + "  # [확인]"
        str = str + Chr(10) + "  "
    Case 1 'MB::IDOK = 1        MB::IDCANCEL = 2
        str = str + "when " + "MB::IDOK"
        str = str + Chr(10) + "  # [확인]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDCANCEL"
        str = str + Chr(10) + "  # [취소]"
        str = str + Chr(10) + "  "
    Case 2 'MB::IDABORT = 3     MB::IDRETRY = 4     MB::IDIGNORE = 5
        str = str + "when " + "MB::IDABORT"
        str = str + Chr(10) + "  # [중단(A)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDRETRY"
        str = str + Chr(10) + "  # [다시 시도(R)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDIGNORE"
        str = str + Chr(10) + "  # [무시(I)]"
        str = str + Chr(10) + "  "
    Case 3 'MB::IDYES = 6       MB::IDNO = 7        MB::IDCANCEL = 2
        str = str + "when " + "MB::IDYES"
        str = str + Chr(10) + "  # [예(Y)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDNO"
        str = str + Chr(10) + "  # [아니오(N)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDCANCEL"
        str = str + Chr(10) + "  # [취소]"
        str = str + Chr(10) + "  "
    Case 4 'MB::IDYES = 6       MB::IDNO = 7
        str = str + "when " + "MB::IDYES"
        str = str + Chr(10) + "  # [예]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDNO"
        str = str + Chr(10) + "  # [아니오]"
        str = str + Chr(10) + "  "
    Case 5 'MB::IDRETRY = 4     MB::IDCANCEL = 2
        str = str + "when " + "MB::IDRETRY"
        str = str + Chr(10) + "  # [다시 시도(R)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDCANCEL"
        str = str + Chr(10) + "  # [취소]"
        str = str + Chr(10) + "  "
    Case 6 'MB::IDCANCEL = 2    MB::IDTRYAGAIN = 10     MB::IDCONTINUE = 11
        str = str + "when " + "MB::IDCANCEL"
        str = str + Chr(10) + "  # [취소]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDTRYAGAIN"
        str = str + Chr(10) + "  # [다시 시도(T)]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::IDCONTINUE"
        str = str + Chr(10) + "  # [계속(C)]"
        str = str + Chr(10) + "  "
    Case &H4000 'MB::IDOK = 1  MB::HELP = 16384
        str = str + "when " + "MB::IDOK"
        str = str + Chr(10) + "  # [확인]"
        str = str + Chr(10) + "  "
        str = str + Chr(10) + "when " + "MB::HELP"
        str = str + Chr(10) + "  # [도움말]"
        str = str + Chr(10) + "  "
    End Select
    
    str = str + Chr(10) + "end"
    
    createCase = str
    
End Function

Private Sub cmdGenerate_Click()
    setData
    If optNormal.Value = True Then txtScript.Text = createNormal
    If optIf.Value = True Then txtScript.Text = createIf
    If optCase.Value = True Then txtScript.Text = createCase
End Sub

Private Sub cmdTest_Click()
    Dim title As String
    setData
    title = Me.Caption
    If Len(txtTitle.Text) > 0 Then title = txtTitle
    Call MessageBox(Form1.hwnd, txtContent, title, dataButton Or dataIcon Or dataButtonFocus Or dataRight Or dataR2L Or dataTopMost)
End Sub

Private Sub Form_Initialize()
    txtTitle.Text = "제목"
    txtContent.Text = "내용"
    
    ' // 버튼 종류
    With cmbButton
        .AddItem ("확인")                               'MB::OK = 0
        .AddItem ("확인, 취소")                         'MB::OKCANCEL = 1
        .AddItem ("중단(A), 다시 시도(R), 무시(I)")     'MB::ABORTRETRYIGNORE = 2
        .AddItem ("예(Y), 아니요(N), 취소")             'MB::YESNOCANCEL = 3
        .AddItem ("예(Y), 아니요(N)")                   'MB::YESNO = 4
        .AddItem ("다시 시도(R), 취소")                 'MB::RETRYCANCEL = 5
        .AddItem ("취소, 다시 시도(T), 계속(C)")        'MB::CANCELTRYCONTINUE = 6
        .AddItem ("확인, 도움말")                       'MB::HELP = 16384 (0x00004000)
    End With
    ' // 아이콘
    With imgCmbIcon.ComboItems
        .Add , , "없음", 1                  '0
        .Add , , "오류", 2                  'MB::ICONSTOP = 16
        .Add , , "질문", 3                  'MB::ICONQUESTION = 32
        .Add , , "경고", 4                  'MB::ICONEXCLAMATION = 48
        .Add , , "알림", 5                  'MB::ICONINFORMATION = 64
    End With
    ' // 버튼 포커스
    With cmbFocus
        .AddItem ("첫 번째")
        .AddItem ("두 번째")
        .AddItem ("세 번째")
        .AddItem ("네 번째")
    End With
    
    cmbButton.ListIndex = 0
    imgCmbIcon.SelectedItem = imgCmbIcon.ComboItems(1)
    cmbFocus.ListIndex = 0
    
    optNormal.Value = True
End Sub


Private Sub setData()
' // 버튼 종류
    If cmbButton.ListIndex = 0 Then dataButton = 0: dataRGSSButton = "MB::OK"
    If cmbButton.ListIndex = 1 Then dataButton = 1: dataRGSSButton = "MB::OKCANCEL"
    If cmbButton.ListIndex = 2 Then dataButton = 2: dataRGSSButton = "MB::ABORTRETRYIGNORE"
    If cmbButton.ListIndex = 3 Then dataButton = 3: dataRGSSButton = "MB::YESNOCANCEL"
    If cmbButton.ListIndex = 4 Then dataButton = 4: dataRGSSButton = "MB::YESNO"
    If cmbButton.ListIndex = 5 Then dataButton = 5: dataRGSSButton = "MB::RETRYCANCEL"
    If cmbButton.ListIndex = 6 Then dataButton = 6: dataRGSSButton = "MB::CANCELTRYCONTINUE"
    If cmbButton.ListIndex = 7 Then dataButton = &H4000: dataRGSSButton = "MB::HELP"
' // 아이콘
    If imgCmbIcon.SelectedItem.Index = 1 Then dataIcon = 0: dataRGSSIcon = ""
    If imgCmbIcon.SelectedItem.Index = 2 Then dataIcon = 16: dataRGSSIcon = "MB::ICONSTOP"
    If imgCmbIcon.SelectedItem.Index = 3 Then dataIcon = 32: dataRGSSIcon = "MB::ICONQUESTION"
    If imgCmbIcon.SelectedItem.Index = 4 Then dataIcon = 48: dataRGSSIcon = "MB::ICONEXCLAMATION"
    If imgCmbIcon.SelectedItem.Index = 5 Then dataIcon = 64: dataRGSSIcon = "MB::ICONINFORMATION"
' // 버튼 포커스
    If cmbFocus.ListIndex = 0 Then dataButtonFocus = 0: dataRGSSButtonFocus = "MB::DEFBUTTON1"
    If cmbFocus.ListIndex = 1 Then dataButtonFocus = &H100: dataRGSSButtonFocus = "MB::DEFBUTTON2"
    If cmbFocus.ListIndex = 2 Then dataButtonFocus = &H200: dataRGSSButtonFocus = "MB::DEFBUTTON3"
    If cmbFocus.ListIndex = 3 Then dataButtonFocus = &H300: dataRGSSButtonFocus = "MB::DEFBUTTON4"
' // 오른쪽 정렬
    If chkRight.Value = 1 Then
        dataRight = &H80000
        dataRGSSRight = "MB::RIGHT"
    Else
        dataRight = 0
        dataRGSSRight = ""
    End If
' // 오른쪽에서 왼쪽 읽기
    If chkR2L.Value = 1 Then
        dataR2L = &H100000
        dataRGSSR2L = "MB::RTLREADING"
    Else
        dataR2L = 0
        dataRGSSR2L = ""
    End If
' // 항상 위로
    If chkTopMost.Value = 1 Then
        dataTopMost = &H40000
        dataRGSSTopMost = "MB::TOPMOST"
    Else
        dataTopMost = 0
        dataRGSSTopMost = ""
    End If
End Sub

