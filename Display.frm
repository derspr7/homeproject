VERSION 5.00
Begin VB.Form Display 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const Istanbul = 0
Private Const Berlin = 1
Private Const London = 2
Private Const Ottova = 4
Private Const Pekin = 5

Dim SelectedCity As Integer

Private Sub Form_Load()

    With Combo1
    
      .Clear
      .AddItem "Istanbul"
      .ItemData(.NewIndex) = 1
      .AddItem "Berlin"
      .ItemData(.NewIndex) = 2
      .AddItem "London"
      .ItemData(.NewIndex) = 3
      .AddItem "Ottova"
      .ItemData(.NewIndex) = 4
      .AddItem "Pekin"
      .ItemData(.NewIndex) = 5

    End With
    


    Timer1.Interval = 60
End Sub



Private Sub Timer1_Timer()


    Dim Hour As String
    Dim Minute As String
    Dim Second As String

    Hour = VBA.Hour(Now)
    Minute = VBA.Minute(Now)
    Second = VBA.Second(Now)
    

    Label1.Caption = GetCurrentTime(SelectedCity, Hour, Minute, Second)
 
    

End Sub


Private Function GetCurrentTime(City As Integer, Hour As String, Minute As String, Second As String) As String

    

    If City = Istanbul Then
    
        Hour = Hour + 3
        Minute = Minute
        Second = Second
    
    ElseIf City = Berlin Then
    
        Hour = Hour + 2
        Minute = Minute
        Second = Second
    
    ElseIf City = London Then
    
        Hour = Hour + 1
        Minute = Minute
        Second = Second
    
    ElseIf City = Ottova Then
    
        Hour = Hour - 4
        Minute = Minute
        Second = Second
    
    ElseIf City = Pekin Then
    
        Hour = Hour + 8
        Minute = Minute
        Second = Second
    
    End If
    
    
    If Hour < 0 Then
        Hour = Hour + 24
    ElseIf Hour > 23 Then
        Hour = Hour - 24
    End If
    
    If Minute < 0 Then
        Minute = Minute + 60
    ElseIf Minute > 59 Then
        Minute = Minute - 60
    End If
    
    GetCurrentTime = Hour & " : " & Minute & " : " & Second

End Function



Private Sub Combo1_Click()


   If Combo1.Text = "Istanbul" Then
        SelectedCity = Istanbul
    ElseIf Combo1.Text = "Berlin" Then
        SelectedCity = Berlin
    ElseIf Combo1.Text = "London" Then
        SelectedCity = London
    ElseIf Combo1.Text = "Ottova" Then
        SelectedCity = Ottova
    ElseIf Combo1.Text = "Pekin" Then
        SelectedCity = Pekin
    End If
End Sub



