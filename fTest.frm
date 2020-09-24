VERSION 5.00
Begin VB.Form fTest 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8295
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ListBox lstSpeakers 
      Height          =   1230
      Left            =   4905
      TabIndex        =   5
      Top             =   300
      Width           =   3240
   End
   Begin VB.PictureBox picIcon 
      BorderStyle     =   0  'Kein
      Height          =   540
      Left            =   15
      Picture         =   "fTest.frx":08CA
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   3
      Top             =   345
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton btTest 
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   3270
      TabIndex        =   2
      Top             =   375
      Width           =   885
   End
   Begin VB.CommandButton btTest 
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   1897
      TabIndex        =   1
      Top             =   375
      Width           =   885
   End
   Begin VB.CommandButton btTest 
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   375
      Width           =   885
   End
   Begin VB.Label lb 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Speakers - Click on one to select"
      Height          =   195
      Left            =   4965
      TabIndex        =   6
      Top             =   60
      Width           =   3030
   End
   Begin VB.Label lbRes 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   2295
      TabIndex        =   4
      Top             =   1275
      Width           =   90
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long

Private Speaker As SpVoice
Private Token As ISpeechObjectToken
Private UserName As String
Private Internal As Boolean

Private Sub btTest_Click(Index As Integer)

  Dim Res As Long

    lbRes = vbNullString
    Select Case Index
      Case 0
        Res = MsgBoxEx("This Message Box speaks. Can you understand me?", vbYesNo Or vbInformation Or vbDefaultButton2, "Speak", 5500, -2, , 38, -40, picIcon.Picture.Handle, rid_Ja & "|" & rid_Nein, "&Yes, I do|&No, I can't", Speaker)
      Case 1
        Res = MsgBoxEx("This Message Box uses the standard Windows noises.", , "DingDong", , -2, , , -40, Icon.Handle, rid_OK, "Oh well...", vbCritical)
      Case 2
        Res = MsgBoxEx("This Message Box beeps; could you hear that?", vbOKCancel Or vbQuestion, "Beep", , -2, , -38, -40, , rid_OK & "|" & rid_Abbrechen, "&That's fine|Too &noisy here", 440.05)
    End Select

    Select Case Res
      Case vbYes
        lbRes = "vbYes"
      Case vbRetry
        lbRes = "vbRetry"
      Case vbOK
        lbRes = "vbOK"
      Case vbNo
        lbRes = "vbNo"
      Case vbIgnore
        lbRes = "vbIgnore"
      Case vbCancel
        lbRes = "vbCancel"
      Case vbAbort
        lbRes = "vbAbort"
      Case Else
        lbRes = "Timed out"
    End Select
    lbRes = "MsgBoxEx box returned: " & Res & " = " & lbRes

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

  Dim l As Long

    Set Speaker = New SpVoice
    With lstSpeakers
        For Each Token In Speaker.GetVoices
            .AddItem Token.GetDescription & "  [" & LCase$(Token.GetAttribute("Gender")) & "]" 'Add to listbox
        Next Token
        l = 128
        UserName = String$(l, 0)
        GetUserName UserName, l
        UserName = Left$(UserName, l - 1)
        Internal = True
        .ListIndex = 0
        Internal = False
    End With 'LSTSPEAKERS

End Sub

Private Sub lstSpeakers_Click()

    If Not Internal Then
        With Speaker
            Set .Voice = .GetVoices().Item(lstSpeakers.ListIndex)
            .Speak "Hallo, " & UserName & ", I am  " & Speaker.Voice.GetDescription & ".", SVSFlagsAsync
        End With 'SPEAKER
    End If

End Sub

':) Ulli's VB Code Formatter V2.22.5 (2007-Jan-21 16:40)  Decl: 9  Code: 76  Total: 85 Lines
':) CommentOnly: 0 (0%)  Commented: 3 (3,5%)  Empty: 17 (20%)  Max Logic Depth: 3
