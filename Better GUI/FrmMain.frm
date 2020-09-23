VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "PCS Interactive"
   ClientHeight    =   9750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FrmMain.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   2700
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.PictureBox PicRange 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1695
      Index           =   1
      Left            =   10380
      Picture         =   "FrmMain.frx":0442
      ScaleHeight     =   1695
      ScaleWidth      =   1935
      TabIndex        =   19
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox PicOutPut 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   6120
      Picture         =   "FrmMain.frx":AFC8
      ScaleHeight     =   1695
      ScaleWidth      =   1935
      TabIndex        =   18
      Top             =   7980
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox PicRange 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Index           =   0
      Left            =   8280
      Picture         =   "FrmMain.frx":15B4E
      ScaleHeight     =   1695
      ScaleWidth      =   1935
      TabIndex        =   17
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox TxtInput 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1860
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox Canvas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   180
      ScaleHeight     =   1695
      ScaleWidth      =   1725
      TabIndex        =   16
      Top             =   180
      Width           =   1725
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   10
      Left            =   5820
      TabIndex        =   26
      Top             =   7020
      Width           =   7035
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   9
      Left            =   5820
      TabIndex        =   25
      Top             =   6600
      Width           =   7035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   10
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":206D4
      TabIndex        =   24
      Top             =   7020
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   9
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":20B16
      TabIndex        =   23
      Top             =   6600
      Width           =   3915
   End
   Begin VB.Label LblNext 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2460
      TabIndex        =   22
      Top             =   8160
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label LblPrev 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   840
      TabIndex        =   21
      Top             =   8160
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PCS Software"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Index           =   0
      Left            =   2220
      TabIndex        =   20
      Tag             =   "NoSelect"
      Top             =   660
      Width           =   3855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   2
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":20F58
      TabIndex        =   1
      Top             =   3660
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   8
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":2139A
      TabIndex        =   7
      Top             =   6180
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   7
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":217DC
      TabIndex        =   6
      Top             =   5760
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   6
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":21C1E
      TabIndex        =   5
      Top             =   5340
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   5
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":22060
      TabIndex        =   4
      Top             =   4920
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   4
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":224A2
      TabIndex        =   3
      Top             =   4500
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "I Would Like To Register With PCS."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   3
      Left            =   1860
      MouseIcon       =   "FrmMain.frx":228E4
      TabIndex        =   2
      Top             =   4080
      Width           =   3915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "What Can We Do For You?"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   1
      Left            =   2340
      TabIndex        =   0
      Tag             =   "NoSelect"
      Top             =   1560
      Width           =   2835
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   8
      Left            =   5880
      TabIndex        =   14
      Top             =   6180
      Width           =   7035
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   7
      Left            =   5880
      TabIndex        =   13
      Top             =   5760
      Width           =   7035
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   6
      Left            =   5880
      TabIndex        =   12
      Top             =   5340
      Width           =   7035
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   5
      Left            =   5880
      TabIndex        =   11
      Top             =   4920
      Width           =   7035
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   10
      Top             =   4500
      Width           =   7035
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   9
      Top             =   4080
      Width           =   7035
   End
   Begin VB.Label LblOutPut 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   8
      Top             =   3660
      Width           =   7035
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte


Private Mode As Single
Private ChangeView As Boolean
Private Steps As Single

Private Type LabelText
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
            
Private Type LabelInfo
    LabelInfo() As LabelText
End Type

Private Labels() As LabelInfo

Private AllowClick As Boolean
Private Active As Single
Private PauseTime As Single
Private TRed As Variant, TGreen As Variant, TBlue As Variant
Private Red As Variant, Green As Variant, Blue As Variant

Private Sub Canvas_Click()
    Form_Unload (1)
End Sub

Private Sub Combo1_Click()
On Error Resume Next
        LblOutPut(Combo1.Tag).Caption = Combo1.List(Combo1.ListIndex)
        Combo1.Visible = False
        If LblOutPut(Combo1.Tag).Caption = "Other" Then
            MoveTextBox CSng(Combo1.Tag), 0
        Else
            If LblOutPut(Val(Combo1.Tag) + 1).Tag = "Input" Then
                MoveTextBox CSng(Combo1.Tag) + 1, 0
            End If
        End If
        
End Sub

Private Sub Form_Activate()
    
    'Load And Play The Backing Track
    SendToMCI "close MedCH1", Me
    SendToMCI "open """ & App.Path & "\Morcheeba.mp3" & """ alias MedCH1", Me
    SendToMCI "Play MedCH1 Repeat", Me
    
    
    Active = 0
    Steps = 10
    FadeItemsIn 0
    
End Sub
Private Sub LoadPageInformation()

    'this Info Will actually be in a resource file, I have left it here so you can see
    'How it works
    Dim LP As Single
    Dim NumberOfPages As Single
    NumberOfPages = 20
    
    ReDim Preserve Labels(NumberOfPages)
    
    For LP = 0 To UBound(Labels)
        ReDim Preserve Labels(LP).LabelInfo(10)
    Next LP
    
    Dim OneThird As Single
    
    OneThird = Int(Screen.Width / 4)
    
    'Page Information
    'Labels(Page Number).LabelInfo(Label Number).Property
    
    Labels(0).LabelInfo(0).Fade = 100
    Labels(0).LabelInfo(0).Caption = "Welcome To PCS Interactive."
    Labels(0).LabelInfo(0).Tag = "NoSelect"
    Labels(0).LabelInfo(1).Caption = "How Can We Help?"
    Labels(0).LabelInfo(1).Tag = "NoSelect"
    Labels(0).LabelInfo(1).Fade = 100
    Labels(0).LabelInfo(2).Caption = "I Would Like To Register With PCS."
    Labels(0).LabelInfo(2).Tag = "1"
    Labels(0).LabelInfo(2).Left = 1500
    Labels(0).LabelInfo(2).Fade = 100
    
    Labels(0).LabelInfo(4).Caption = "I Have Registered And Would Like To Report A Fault."
    Labels(0).LabelInfo(4).Tag = "1"
    Labels(0).LabelInfo(4).Left = 2000
    Labels(0).LabelInfo(4).Fade = 100
    Labels(0).LabelInfo(6).Caption = "I Have Registered, I Don't Have A Fault But Do Require A Service."
    Labels(0).LabelInfo(6).Tag = "10"
    Labels(0).LabelInfo(6).Left = 2500
    Labels(0).LabelInfo(6).Fade = 100
    Labels(0).LabelInfo(8).Caption = "Exit PCS Interactive."
    Labels(0).LabelInfo(8).Tag = "End"
    Labels(0).LabelInfo(8).Left = 3000
    Labels(0).LabelInfo(8).Fade = 100
    
    Labels(1).LabelInfo(0).Fade = 100
    Labels(1).LabelInfo(0).Caption = "Register A Home, Business Or Charity"
    Labels(1).LabelInfo(0).Tag = "NoSelect"
    Labels(1).LabelInfo(1).Caption = "Registering Your Benefits"
    Labels(1).LabelInfo(1).Tag = "NoSelect"
    Labels(1).LabelInfo(1).Fade = 100
    Labels(1).LabelInfo(2).Caption = "I Would Like To Register For Home Use."
    Labels(1).LabelInfo(2).Tag = "2"
    Labels(1).LabelInfo(2).Left = 1500
    Labels(1).LabelInfo(2).Fade = 100
    Labels(1).LabelInfo(4).Caption = "I Would Like To Register For Business Use."
    Labels(1).LabelInfo(4).Tag = "2"
    Labels(1).LabelInfo(4).Left = 2000
    Labels(1).LabelInfo(4).Fade = 100
    Labels(1).LabelInfo(6).Caption = "I Would Like To Register For Charity Use."
    Labels(1).LabelInfo(6).Tag = "2"
    Labels(1).LabelInfo(6).Left = 2500
    Labels(1).LabelInfo(6).Fade = 100
    Labels(1).LabelInfo(0).Previous = True
    Labels(1).LabelInfo(0).PrevTag = 0
    
    Labels(2).LabelInfo(0).Fade = 100
    Labels(2).LabelInfo(0).Caption = "Registering As A Home User."
    Labels(2).LabelInfo(0).Tag = "NoSelect"
    Labels(2).LabelInfo(1).Caption = "Register Name Details"
    Labels(2).LabelInfo(1).Tag = "NoSelect"
    Labels(2).LabelInfo(1).Fade = 100
    
    Labels(2).LabelInfo(2).Caption = "Title."
    Labels(2).LabelInfo(2).Tag = "ComboTitle"
    Labels(2).LabelInfo(2).Left = OneThird
    Labels(2).LabelInfo(2).Fade = 30
    Labels(2).LabelInfo(3).Caption = "Full Name."
    Labels(2).LabelInfo(3).Tag = "Input"
    Labels(2).LabelInfo(3).Left = OneThird
    Labels(2).LabelInfo(3).Fade = 30
    Labels(2).LabelInfo(5).Caption = "Christian Name"
    Labels(2).LabelInfo(5).Tag = "Input"
    Labels(2).LabelInfo(5).Left = OneThird
    Labels(2).LabelInfo(5).Fade = 30
    Labels(2).LabelInfo(6).Caption = "Surname"
    Labels(2).LabelInfo(6).Tag = "Input"
    Labels(2).LabelInfo(6).Left = OneThird
    Labels(2).LabelInfo(6).Fade = 30
    Labels(2).LabelInfo(7).Caption = "Middle Initials"
    Labels(2).LabelInfo(7).Tag = "Input"
    Labels(2).LabelInfo(7).Left = OneThird
    Labels(2).LabelInfo(7).Fade = 30
    Labels(2).LabelInfo(0).Previous = True
    Labels(2).LabelInfo(0).Next = True
    Labels(2).LabelInfo(0).PrevTag = 1
    Labels(2).LabelInfo(0).NextTag = 5
    Labels(2).LabelInfo(0).NextFrom = 1
    
    Labels(3).LabelInfo(0).Fade = 100
    Labels(3).LabelInfo(0).Caption = "Registering As A Business User."
    Labels(3).LabelInfo(0).Tag = "NoSelect"
    Labels(3).LabelInfo(1).Fade = 100
    Labels(3).LabelInfo(1).Caption = "Registering Your Benefits"
    Labels(3).LabelInfo(1).Tag = "NoSelect"
    
    Labels(3).LabelInfo(2).Caption = "Business Or Company Name."
    Labels(3).LabelInfo(2).Tag = "0"
    Labels(3).LabelInfo(2).Fade = 30
    Labels(3).LabelInfo(2).Left = 1500
    Labels(3).LabelInfo(2).Tag = "Input"
    Labels(3).LabelInfo(3).Caption = "Address Line 1."
    Labels(3).LabelInfo(3).Tag = "0"
    Labels(3).LabelInfo(3).Fade = 30
    Labels(3).LabelInfo(3).Left = 1500
    Labels(3).LabelInfo(3).Tag = "Input"
    Labels(3).LabelInfo(4).Caption = "Address Line 2."
    Labels(3).LabelInfo(4).Tag = "0"
    Labels(3).LabelInfo(4).Fade = 30
    Labels(3).LabelInfo(4).Left = 1500
    Labels(3).LabelInfo(4).Tag = "Input"
    Labels(3).LabelInfo(5).Caption = "Town Or City"
    Labels(3).LabelInfo(5).Tag = "0"
    Labels(3).LabelInfo(5).Fade = 30
    Labels(3).LabelInfo(5).Left = 1500
    Labels(3).LabelInfo(5).Tag = "Input"
    Labels(3).LabelInfo(6).Caption = "County"
    Labels(3).LabelInfo(6).Tag = "0"
    Labels(3).LabelInfo(6).Fade = 30
    Labels(3).LabelInfo(6).Left = 1500
    Labels(3).LabelInfo(6).Tag = "Input"
    Labels(3).LabelInfo(7).Caption = "PostCode"
    Labels(3).LabelInfo(7).Tag = "0"
    Labels(3).LabelInfo(7).Fade = 30
    Labels(3).LabelInfo(7).Left = 1500
    Labels(3).LabelInfo(7).Tag = "Input"
    Labels(3).LabelInfo(8).Caption = "Exit"
    Labels(3).LabelInfo(8).Tag = 0
    Labels(3).LabelInfo(8).Fade = 30
    Labels(3).LabelInfo(8).Left = 1000
    
    Labels(4).LabelInfo(0).Fade = 100
    Labels(4).LabelInfo(0).Caption = "Registering As A Charity User."
    Labels(4).LabelInfo(0).Tag = "NoSelect"
    Labels(4).LabelInfo(1).Fade = 100
    Labels(4).LabelInfo(1).Caption = "Registering Your Benefits"
    Labels(4).LabelInfo(1).Tag = "NoSelect"
    
    Labels(4).LabelInfo(2).Caption = "Charity Name."
    Labels(4).LabelInfo(2).Tag = "0"
    Labels(4).LabelInfo(2).Left = 1500
    Labels(4).LabelInfo(2).Fade = 30
    Labels(4).LabelInfo(2).Tag = "Input"
    Labels(4).LabelInfo(3).Caption = "Address Line 1."
    Labels(4).LabelInfo(3).Tag = "0"
    Labels(4).LabelInfo(3).Left = 1500
    Labels(4).LabelInfo(3).Fade = 30
    Labels(4).LabelInfo(3).Tag = "Input"
    Labels(4).LabelInfo(4).Caption = "Address Line 2."
    Labels(4).LabelInfo(4).Tag = "0"
    Labels(4).LabelInfo(4).Left = 1500
    Labels(4).LabelInfo(4).Fade = 30
    Labels(4).LabelInfo(4).Tag = "Input"
    Labels(4).LabelInfo(5).Caption = "Town Or City"
    Labels(4).LabelInfo(5).Tag = "0"
    Labels(4).LabelInfo(5).Left = 1500
    Labels(4).LabelInfo(5).Fade = 30
    Labels(4).LabelInfo(5).Tag = "Input"
    Labels(4).LabelInfo(6).Caption = "County"
    Labels(4).LabelInfo(6).Tag = "0"
    Labels(4).LabelInfo(6).Left = 1500
    Labels(4).LabelInfo(6).Fade = 30
    Labels(4).LabelInfo(6).Tag = "Input"
    Labels(4).LabelInfo(7).Caption = "PostCode"
    Labels(4).LabelInfo(7).Tag = "0"
    Labels(4).LabelInfo(7).Left = 1500
    Labels(4).LabelInfo(7).Fade = 30
    Labels(4).LabelInfo(7).Tag = "Input"
    Labels(4).LabelInfo(8).Caption = "Exit"
    Labels(4).LabelInfo(8).Tag = 0
    Labels(4).LabelInfo(8).Fade = 30
    Labels(4).LabelInfo(8).Left = 1000
    
    'Address Details
    Labels(5).LabelInfo(0).Caption = "Registering As A Home User."
    Labels(5).LabelInfo(0).Tag = "NoSelect"
    Labels(5).LabelInfo(1).Caption = "Register Contact Details"
    Labels(5).LabelInfo(1).Tag = "NoSelect"
    Labels(5).LabelInfo(2).Caption = "Daytime Phone Number."
    Labels(5).LabelInfo(2).Tag = "Input"
    Labels(5).LabelInfo(2).Left = OneThird
    Labels(5).LabelInfo(2).Fade = 30
    Labels(5).LabelInfo(3).Caption = "Evening Phone Number."
    Labels(5).LabelInfo(3).Tag = "Input"
    Labels(5).LabelInfo(3).Fade = 30
    Labels(5).LabelInfo(3).Left = OneThird
    Labels(5).LabelInfo(4).Caption = "Mobile Phone Number."
    Labels(5).LabelInfo(4).Tag = "Input"
    Labels(5).LabelInfo(4).Fade = 30
    Labels(5).LabelInfo(4).Left = OneThird
    Labels(5).LabelInfo(5).Caption = "Fax Number."
    Labels(5).LabelInfo(5).Tag = "Input"
    Labels(5).LabelInfo(5).Fade = 30
    Labels(5).LabelInfo(5).Left = OneThird
    Labels(5).LabelInfo(7).Caption = "Email Address."
    Labels(5).LabelInfo(7).Tag = "Input"
    Labels(5).LabelInfo(7).Fade = 30
    Labels(5).LabelInfo(7).Left = OneThird
    Labels(5).LabelInfo(0).Next = True
    Labels(5).LabelInfo(0).Previous = True
    Labels(5).LabelInfo(0).NextFrom = 1
    Labels(5).LabelInfo(0).PrevFrom = 1
    Labels(5).LabelInfo(0).PrevTag = 2
    Labels(5).LabelInfo(0).NextTag = 6
    
    Labels(6).LabelInfo(0).Fade = 100
    Labels(6).LabelInfo(0).Caption = "Registering As A Home User."
    Labels(6).LabelInfo(0).Tag = "NoSelect"
    Labels(6).LabelInfo(1).Caption = "Register Address Details"
    Labels(6).LabelInfo(1).Tag = "NoSelect"
    Labels(6).LabelInfo(1).Fade = 100
    
    Labels(6).LabelInfo(2).Caption = "House Name Or Number"
    Labels(6).LabelInfo(2).Tag = "Input"
    Labels(6).LabelInfo(2).Left = OneThird
    Labels(6).LabelInfo(2).Fade = 30
    Labels(6).LabelInfo(3).Caption = "Address Line 1"
    Labels(6).LabelInfo(3).Tag = "Input"
    Labels(6).LabelInfo(3).Left = OneThird
    Labels(6).LabelInfo(3).Fade = 30
    Labels(6).LabelInfo(4).Caption = "Address Line 2"
    Labels(6).LabelInfo(4).Tag = "Input"
    Labels(6).LabelInfo(4).Left = OneThird
    Labels(6).LabelInfo(4).Fade = 30
    Labels(6).LabelInfo(5).Caption = "Town Or City"
    Labels(6).LabelInfo(5).Tag = "Input"
    Labels(6).LabelInfo(5).Left = OneThird
    Labels(6).LabelInfo(5).Fade = 30
    Labels(6).LabelInfo(6).Caption = "County"
    Labels(6).LabelInfo(6).Tag = "Input"
    Labels(6).LabelInfo(6).Left = OneThird
    Labels(6).LabelInfo(6).Fade = 30
    Labels(6).LabelInfo(7).Caption = "PostCode"
    Labels(6).LabelInfo(7).Tag = "Input"
    Labels(6).LabelInfo(7).Left = OneThird
    Labels(6).LabelInfo(7).Fade = 30
    Labels(6).LabelInfo(0).Previous = True
    Labels(6).LabelInfo(0).PrevTag = 5
    Labels(6).LabelInfo(0).PrevFrom = 1
    
    
    Labels(10).LabelInfo(0).Fade = 100
    Labels(10).LabelInfo(0).Caption = "Services From PCS."
    Labels(10).LabelInfo(0).Tag = "NoSelect"
    Labels(10).LabelInfo(1).Caption = "How Can We Help?"
    Labels(10).LabelInfo(1).Tag = "NoSelect"
    Labels(10).LabelInfo(1).Fade = 100
    
    Labels(10).LabelInfo(2).Caption = "Hardware Installation And Upgrades"
    Labels(10).LabelInfo(2).Tag = "0"
    Labels(10).LabelInfo(2).Left = OneThird
    Labels(10).LabelInfo(2).Fade = 100
    Labels(10).LabelInfo(4).Caption = "Network Installation And Servicing"
    Labels(10).LabelInfo(4).Tag = "0"
    Labels(10).LabelInfo(4).Left = OneThird + 500
    Labels(10).LabelInfo(4).Fade = 100
    Labels(10).LabelInfo(6).Caption = "Software Upgrades"
    Labels(10).LabelInfo(6).Tag = "0"
    Labels(10).LabelInfo(6).Left = OneThird + 1000
    Labels(10).LabelInfo(6).Fade = 100
    Labels(10).LabelInfo(8).Caption = "Software Development Services"
    Labels(10).LabelInfo(8).Tag = "0"
    Labels(10).LabelInfo(8).Left = OneThird + 1500
    Labels(10).LabelInfo(8).Fade = 100
    Labels(10).LabelInfo(10).Caption = "Other Services"
    Labels(10).LabelInfo(10).Tag = "0"
    Labels(10).LabelInfo(10).Left = OneThird + 2000
    Labels(10).LabelInfo(10).Fade = 100
    Labels(10).LabelInfo(0).Previous = True
    Labels(10).LabelInfo(0).PrevTag = 0
End Sub

Private Sub SlideIcons()
    
    Dim Movement As Single
    Movement = 5
    
    If ChangeView = True Then Exit Sub
    
    ChangeView = True
    
    Dim LP  As Single
    Dim Start As Single
    
    Mode = Mode + 1
    If Mode > PicRange.UBound Then Mode = 0
    
    Dim X As Long, Y As Long, Rep As Long
    
    'Move Slide Out
    For Start = 0 To Steps - 1
        For Rep = 0 To T2P(Canvas.Width) Step Movement
            For Y = Start To T2P(Canvas.Height) Step Steps
                For X = 0 To T2P(Canvas.Width) - Rep
                    SetPixelV Canvas.hDC, X, Y, GetPixel(PicOutPut.hDC, X + Rep, Y)
                Next X
            Next Y
            DoEvents
        Next Rep
    Next Start
    
    Canvas.Picture = LoadPicture("")
    ChangeView = False
    
End Sub

Private Sub SlideIconsIn()

    Dim Movement As Single
    Movement = 5
    
    Dim X As Long, Y As Long, Rep As Long
    If ChangeView = True Then Exit Sub
    
    ChangeView = True
    
    Dim LP  As Single
    Dim Start As Single
    'Move Next Slide In
    For Start = 0 To Steps - 1
        For Rep = 0 To T2P(Canvas.Width) Step Movement
            For Y = Start To T2P(Canvas.Height) Step Steps
                For X = 0 To Rep
                    SetPixelV Canvas.hDC, T2P(Canvas.Width) - Rep + X, Y, GetPixel(PicRange(Mode).hDC, X, Y)
                Next X
            Next Y
            DoEvents
        Next Rep
    Next Start
    
    PicOutPut.Picture = PicRange(Mode).Picture
    Canvas.Picture = PicRange(Mode).Picture
    
    ChangeView = False
    
End Sub

Private Function T2P(Twip As Long) As Long
    
    'Convert Twip To Pixel For Graphics Routines
    T2P = Int(Twip / 15)

End Function
Private Sub Form_Load()
    
    Dim Control As Object
    Dim LP As Single
    
    
    TRed = 128
    TGreen = 128
    TBlue = 255
    
    LblNext.ForeColor = RGB(TRed, TGreen, TBlue)
    LblPrev.ForeColor = RGB(TRed, TGreen, TBlue)
    
    For LP = 0 To Label1.UBound
            Label1(LP).ForeColor = Me.BackColor
    Next LP
    
    
    
    
    
    LoadPageInformation
    
End Sub

Private Function SendToMCI(SendString As String, FromForm As Form)

    Dim RetStr As String * 255
    
    Call mciSendString(SendString, RetStr, 255, FromForm.hWnd)
    
    SendToMCI = Replace(RetStr, Chr(0), "")

End Function


Private Sub FadeItemsIn(FromControl As Single)

'Fade The Labels In

AllowClick = False

Dim RDiff As Variant
Dim BDiff As Variant
Dim GDiff As Variant
Dim LP As Single
Dim LblLP As Single

SetUpLabels

If FromControl = 0 Then
    SlideIconsIn
End If

'The TColor Are the target colours, I have chosen this bluey colour, The Target Colour has been set in the Load
'Load routine. During which time all the labels forecolors have been set to the same as the form backcolor.

'I get the differance between each colour then divide the differance by 10, this gives me a 10 step fade in

For LblLP = FromControl To Label1.UBound
    If Label1(LblLP).Caption <> "" Then
    ToRGB (Label1(LblLP).ForeColor), Red, Green, Blue
    
    'Get The Differance Between Current Colour And Target Colour
    RDiff = Red - TRed
    GDiff = Green - TGreen
    BDiff = Blue - TBlue
                            
    'Divide The Differances By 10
    RDiff = Int(RDiff / 10)
    GDiff = Int(GDiff / 10)
    BDiff = Int(BDiff / 10)
    
    'Fade The Labels In In  10 Steps
    For LP = 0 To 10
        Label1(LblLP).ForeColor = RGB(Red - (RDiff * LP), Green - (GDiff * LP), Blue - (BDiff * LP))
        Sleep Labels(Active).LabelInfo(LblLP).Fade
        Label1(LblLP).Refresh
        DoEvents
    Next LP
    End If
Next LblLP
            
            'If Display Next Button Is Tagged Display The Button
            If Labels(Active).LabelInfo(0).Next = True Then
                LblNext.Left = Me.Width - LblNext.Width - 1000
                LblNext.Top = Me.Height - LblNext.Height - 500
                LblNext.Tag = Labels(Active).LabelInfo(0).NextTag
                LblNext.Visible = True
            End If
            
            'If Display Previous Button Is Tagged Display The Button
            If Labels(Active).LabelInfo(0).Previous = True Then
                LblPrev.Left = 1000
                LblPrev.Top = Me.Height - LblPrev.Height - 500
                LblPrev.Tag = Labels(Active).LabelInfo(0).PrevTag
                LblPrev.Visible = True
            End If

AllowClick = True

End Sub

Private Sub SetUpLabels()
    
    
    'Load the information from the type arrey corresponding to the page number
    'Active is the chosen page number variable
    
    Dim LP As Single
    Dim Longest As Single
        For LP = 0 To Label1.UBound
            Label1(LP).Caption = ""
            Label1(LP).Tag = ""
            
            Label1(LP).Caption = Labels(Active).LabelInfo(LP).Caption
            Label1(LP).Tag = Labels(Active).LabelInfo(LP).Tag
            
            If Labels(Active).LabelInfo(LP).Left > 0 Then
                Label1(LP).Left = Labels(Active).LabelInfo(LP).Left
            End If
            
            If Labels(Active).LabelInfo(LP).Top > 0 Then
                Label1(LP).Top = Labels(Active).LabelInfo(LP).Top
            End If
        Next LP
            
            Longest = 2
            
            For LP = 3 To Label1.UBound
                If Len(Label1(LP).Caption) > Len(Label1(Longest).Caption) Then Longest = LP
            Next LP
            
            Label1(1).Left = (Label1(0).Width - (Label1(1).Width / 2)) + Label1(0).Left
            Label1(1).Top = Label1(0).Top + Label1(0).Height + 30
            
            TxtInput.Text = ""
            
            'if the type arrey doesn't hold data, don't display the output labels
            For LP = 2 To Label1.UBound
                If Label1(LP).Tag = "Input" Or Label1(LP).Tag = "ComboTitle" Then
                    LblOutPut(LP).Caption = ""
                    LblOutPut(LP).Visible = True
                    LblOutPut(LP).Top = Label1(LP).Top
                    LblOutPut(LP).Left = Label1(Longest).Left + Label1(Longest).Width + 250
                    LblOutPut(LP).Width = Me.Width - LblOutPut(LP).Left - 300
                Else
                    LblOutPut(LP).Visible = False
                End If
            Next LP
            
            
            'Set the time each label takes to fade
            PauseTime = Labels(Active).LabelInfo(0).Fade
End Sub

Private Sub FadeAllOut(FromControl As Single)


LblNext.Visible = False
LblPrev.Visible = False

'Useing the same method as fade in, but backwards

If FromControl = 0 Then
    SlideIcons
End If
AllowClick = False

Dim RDiff As Variant
Dim BDiff As Variant
Dim GDiff As Variant
Dim LP As Single
Dim LblLP As Single


    TxtInput.Visible = False
    
    For LblLP = FromControl To Label1.UBound
        If Label1(LblLP).Caption <> "" Then
        ToRGB (Label1(LblLP).ForeColor), Red, Green, Blue
                
        RDiff = Red - 255
        GDiff = Green - 255
        BDiff = Blue - 255
                            
        RDiff = Int(RDiff / 10)
        GDiff = Int(GDiff / 10)
        BDiff = Int(BDiff / 10)
                        
        For LP = 0 To 10
            Label1(LblLP).ForeColor = RGB(Red - (RDiff * LP), Green - (GDiff * LP), Blue - (BDiff * LP))
            Label1(LblLP).Refresh
            Sleep 10
        Next LP
                        
        DoEvents
        End If
        If LblLP > 1 Then
            LblOutPut(LblLP).Caption = ""
        End If
        
    Next LblLP



FadeItemsIn (FromControl)

End Sub


Private Function ToRGB(Colour As Long, Optional Red As Variant, Optional Green As Variant, Optional Blue As Variant)
        
        'Converts long colours to R G B
        Red = Colour And 255
        Green = (Colour \ 256) And 255
        Blue = (Colour \ 65536) And 255

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim LP As Single
    
    For LP = 0 To Label1.UBound
            If Label1(LP).FontUnderline = True Then Label1(LP).FontUnderline = False
    Next LP
    
    If LblPrev.FontUnderline = True Then LblPrev.FontUnderline = False
    If LblNext.FontUnderline = True Then LblNext.FontUnderline = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Stop The Music
    SendToMCI "Stop MedCH1", Me
    SendToMCI "close all", Me
    End
    
End Sub

Private Sub Label1_Click(Index As Integer)

Dim UserInput As String

If TxtInput.Visible = True Then
    LblOutPut(TxtInput.Tag).Caption = TxtInput.Text
End If

    'What action to take depending on the tag
 If AllowClick = True Then
    If Label1(Index).Tag <> "NoSelect" Then
        Select Case Label1(Index).Tag
        Case "Input" 'Display text Box When clicked On
            MoveTextBox CSng(Index), 0
        Case "ComboTitle" 'Display Combo Box
            DisplayComboTitle CSng(Index)
        Case Else
        If IsNumeric(Label1(Index).Tag) Then 'If The tag is numeric it is pointing to a new page
            Active = Label1(Index).Tag
        End If
            'Check To See If Data Needs Saveing
            If Label1(Index).Tag = "End" Then EndProgram 'If label tag is End then End.
            FadeAllOut 0
        End Select
    End If
End If

EHand:
End Sub

Private Sub DisplayComboTitle(Index As Single)
    
    
    'Called if a labels Tag is ComboTitle,
    'set the coordinates and content of the combo box
    
    'As there is actually only one text box, save the text boxes data to it's
    'perticulare label before hiding it
    If TxtInput.Tag <> "" Then
        LblOutPut(TxtInput.Tag).Caption = TxtInput.Text
    End If

    TxtInput.Visible = False
    Combo1.Clear
    Combo1.AddItem "Mr"
    Combo1.AddItem "Mrs"
    Combo1.AddItem "Ms"
    Combo1.AddItem "Miss"
    Combo1.AddItem "Dr"
    Combo1.AddItem "Sir"
    Combo1.AddItem "Other"
    Combo1.ListIndex = 0
    Combo1.Tag = Index
    Combo1.Left = LblOutPut(Index).Left - 190
    Combo1.Top = LblOutPut(Index).Top - 100
    Combo1.Visible = True
    
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LP As Single

    For LP = 0 To Label1.UBound
        If Label1(LP).Tag <> "NoSelect" Then
            If LP = Index Then
                If Label1(Index).FontUnderline = False Then Label1(Index).FontUnderline = True
            Else
                If Label1(LP).FontUnderline = True Then Label1(LP).FontUnderline = False
            End If
        End If
    Next LP
    
End Sub




Private Sub LblNext_Click()
Dim Lactive As Single

If AllowClick = True Then
    Lactive = Active
    Active = LblNext.Tag
    FadeAllOut Labels(Lactive).LabelInfo(0).NextFrom
    
End If
End Sub

Private Sub LblNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LblNext.FontUnderline = False Then LblNext.FontUnderline = True
End Sub

Private Sub LblOutPut_Click(Index As Integer)
    Select Case Labels(Active).LabelInfo(Index).Tag
    Case "Input", "ComboTitle"
    If Len(LblOutPut(Index).Caption) > 0 Then
        MoveTextBox CSng(Index), Labels(Active).LabelInfo(0).PrevFrom
    End If
    End Select
End Sub

Private Sub MoveTextBox(ToIndex As Single, IncrementIndex As Single)
        
        'As there is only actually one text box, if we need input in a differant place,
        'save the text to the corresponig label, then move the box to the new coordinates
        
        If TxtInput.Tag <> "" Then
            LblOutPut(TxtInput.Tag).Caption = TxtInput.Text
        End If
        
        If Combo1.Visible = True Then
            LblOutPut(Combo1.Tag).Caption = Combo1.List(Combo1.ListIndex)
            Combo1.Visible = False
        ElseIf TxtInput.Tag = Combo1.Tag And Combo1.Tag <> "" Then
            LblOutPut(Combo1.Tag).Caption = TxtInput.Text
        End If
        
        
If ToIndex <> 0 Then
    TxtInput.Top = LblOutPut(ToIndex).Top
    TxtInput.Left = LblOutPut(ToIndex).Left
    TxtInput.Width = LblOutPut(ToIndex).Width
    TxtInput.Text = LblOutPut(ToIndex).Caption
    TxtInput.Tag = ToIndex
    TxtInput.Visible = True
    TxtInput.SetFocus
    TxtInput.SelStart = Len(TxtInput.Text)

End If

If IncrementIndex <> 0 Then
    TxtInput.Top = LblOutPut(TxtInput.Tag + IncrementIndex).Top
    TxtInput.Left = LblOutPut(TxtInput.Tag + IncrementIndex).Left
    TxtInput.Width = LblOutPut(TxtInput.Tag + IncrementIndex).Width
    TxtInput.Text = LblOutPut(TxtInput.Tag + IncrementIndex).Caption
    TxtInput.Tag = TxtInput.Tag + IncrementIndex
    TxtInput.Visible = True
    TxtInput.SelStart = Len(TxtInput.Text)
End If


End Sub

Private Sub LblPrev_Click()
Dim Pactive As Single


    Pactive = Active
    Active = LblPrev.Tag
    FadeAllOut Labels(Pactive).LabelInfo(0).PrevFrom
    
End Sub

Private Sub LblPrev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LblPrev.FontUnderline = False Then LblPrev.FontUnderline = True
End Sub

Private Sub TxtInput_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
KeyAscii = 0



TxtInput.Text = ConvertProperCase(TxtInput.Text, Label1(TxtInput.Tag).Caption)
    
LblOutPut(TxtInput.Tag).Caption = TxtInput.Text

    If Val(TxtInput.Tag) < LblOutPut.UBound Then
        If Label1(Val(TxtInput.Tag) + 1).Tag = "Input" Then
            MoveTextBox 0, 1
        Else
            TxtInput.Visible = False
        End If
    Else
        TxtInput.Visible = False
    End If
End If

If KeyAscii = 8 And TxtInput.SelStart = 0 Then
    LblOutPut(TxtInput.Tag).Caption = TxtInput.Text
    If Label1(TxtInput.Tag - 1).Tag = "Input" Then
        MoveTextBox 0, -1
        KeyAscii = 0
    End If
End If

End Sub

Private Function ConvertProperCase(Text As String, DescLabel As String) As String

Dim TPos As Long, MiddleInitial As Long
Dim Temp As String
Dim SplitName() As String
Dim LP As Single

'Convert text case to something a bit more pretty.

Select Case DescLabel
    Case "Full Name."
            
            Text = StrConv(Text, vbProperCase)
            
            'Look For Double Barrel Names Seperated By -
            If InStr(1, Text, "-") > 0 Then
                TPos = InStr(1, Text, "-") + 1
                Temp = Text
                Mid$(Temp, TPos, 1) = UCase(Mid$(Temp, TPos, 1))
                Text = Temp
            End If
            
            'Look For Any McNames
            If InStr(1, UCase(Text), "MC") > 0 Then
                TPos = InStr(1, UCase(Text), "MC") + 2
                Temp = Text
                Mid$(Temp, TPos, 1) = UCase(Mid$(Temp, TPos, 1))
                Text = Temp
            End If
            
            If InStr(1, UCase(Text), "O'") > 0 Then
                TPos = InStr(1, UCase(Text), "O'") + 2
                Temp = Text
                Mid$(Temp, TPos, 1) = UCase(Mid$(Temp, TPos, 1))
                Text = Temp
            End If
            
            SplitName = Split(Text)
               
            For LP = 0 To Label1.UBound
                If Label1(LP).Caption = "Middle Initials" Then LblOutPut(LP).Caption = ""
            Next LP
                
            Select Case UBound(SplitName)
                Case 0 'Only One Name Entered
                    MsgBox "Please Enter Your Full Name."
                    'TxtName(0).SetFocus
                    'TxtName(0).SelStart = Len(TxtName(0))
                    Exit Function
                Case 1 'Surname And Christian Name Entered (One Assumes)
                    For LP = 0 To Label1.UBound
                        If Label1(LP).Caption = "Christian Name" Then LblOutPut(LP).Caption = SplitName(0)
                        If Label1(LP).Caption = "Surname" Then LblOutPut(LP).Caption = SplitName(1)
                    Next LP
                Case Else 'Full Name Entered With Middle Names Too.
                For LP = 0 To Label1.UBound
                        If Label1(LP).Caption = "Christian Name" Then LblOutPut(LP).Caption = SplitName(0) 'Assumes First Name Is Christian Name
                        If Label1(LP).Caption = "Surname" Then LblOutPut(LP).Caption = SplitName(UBound(SplitName)) 'Assumes Last Name Is Surname
                    
                    
                    If Label1(LP).Caption = "Middle Initials" Then
                    'Initial Middle Names
                    For MiddleInitial = 1 To UBound(SplitName) - 1
                        SplitName(MiddleInitial) = Left(SplitName(MiddleInitial), 1)
                        LblOutPut(LP).Caption = LblOutPut(LP).Caption & SplitName(MiddleInitial) & ", "
                    Next MiddleInitial
                    
                    LblOutPut(LP).Caption = Left(LblOutPut(LP).Caption, Len(LblOutPut(LP).Caption) - 2)
                    End If
                Next LP
            End Select
            
            ConvertProperCase = Text
        Case Else
            ConvertProperCase = StrConv(Text, vbProperCase)
    End Select
End Function


Private Sub TxtInput_LostFocus()

    LblOutPut(TxtInput.Tag).Caption = TxtInput.Text
    TxtInput.Visible = False
    
End Sub


Private Sub EndProgram()

LblNext.Visible = False
LblPrev.Visible = False

SlideIcons

AllowClick = False

Dim RDiff As Variant
Dim BDiff As Variant
Dim GDiff As Variant
Dim LP As Single
Dim LblLP As Single


    TxtInput.Visible = False
    
    For LblLP = 0 To Label1.UBound
        If Label1(LblLP).Caption <> "" Then
        ToRGB (Label1(LblLP).ForeColor), Red, Green, Blue
                
        RDiff = Red - 255
        GDiff = Green - 255
        BDiff = Blue - 255
                            
        RDiff = Int(RDiff / 10)
        GDiff = Int(GDiff / 10)
        BDiff = Int(BDiff / 10)
                        
        For LP = 0 To 10
            Label1(LblLP).ForeColor = RGB(Red - (RDiff * LP), Green - (GDiff * LP), Blue - (BDiff * LP))
            Label1(LblLP).Refresh
            Sleep 10
        Next LP
                        
        DoEvents
        End If
        If LblLP > 1 Then
            LblOutPut(LblLP).Caption = ""
        End If
        
    Next LblLP

Form_Unload (1)

End Sub
