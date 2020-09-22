VERSION 5.00
Object = "{0A11A543-1724-465B-979B-2E3A27195497}#2.0#0"; "SkinnableButtonProject.ocx"
Begin VB.Form Form1 
   Caption         =   "Test"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   165
      Left            =   2700
      TabIndex        =   6
      Top             =   1575
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Columns         =   4
      Height          =   3765
      Left            =   300
      TabIndex        =   4
      Top             =   1800
      Width           =   3915
   End
   Begin SkinnableButtonProject.SkinnableButton SkinnableButton1 
      Height          =   915
      Left            =   300
      TabIndex        =   0
      Top             =   525
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1614
      Value           =   1
   End
   Begin SkinnableButtonProject.SkinnableButton SkinnableButton2 
      Height          =   915
      Left            =   2475
      TabIndex        =   3
      Top             =   525
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1614
      Value           =   1
   End
   Begin VB.Label Label3 
      Height          =   240
      Left            =   375
      TabIndex        =   5
      Top             =   5625
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Checkbox"
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   150
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Button"
      Height          =   195
      Left            =   750
      TabIndex        =   1
      Top             =   150
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Set SkinnableButton1.Image_Normal = LoadPicture(App.Path & "\Play1.jpg")
    Set SkinnableButton1.Image_Hover = LoadPicture(App.Path & "\Play2.jpg")
    Set SkinnableButton1.Image_Pressed = LoadPicture(App.Path & "\Play3.jpg")
    
    Set SkinnableButton2.Image_Normal = LoadPicture(App.Path & "\Play1.jpg")
    Set SkinnableButton2.Image_Hover = LoadPicture(App.Path & "\Play2.jpg")
    Set SkinnableButton2.Image_Pressed = LoadPicture(App.Path & "\Play3.jpg")

    Set SkinnableButton2.Image_Checked_Normal = LoadPicture(App.Path & "\Play4.jpg")
    Set SkinnableButton2.Image_Checked_Hover = LoadPicture(App.Path & "\Play5.jpg")
    Set SkinnableButton2.Image_Checked_Pressed = LoadPicture(App.Path & "\Play6.jpg")
    SkinnableButton2.Style = CheckBox

End Sub

Private Sub SkinnableButton1_Click()
    List1.AddItem "Click"
End Sub

Private Sub SkinnableButton1_DblClick()
    List1.AddItem "Dblclick"
End Sub

Private Sub SkinnableButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.AddItem "Down"
End Sub

Private Sub SkinnableButton1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label3.Caption = "(" & X & "," & Y & ")"
End Sub

Private Sub SkinnableButton1_MouseOut()
    List1.AddItem "Out"
End Sub

Private Sub SkinnableButton1_MouseOver()
    List1.AddItem "Over"
End Sub

Private Sub SkinnableButton1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    List1.AddItem "Up"
End Sub

Private Sub SkinnableButton2_Click()
    Check1.Value = SkinnableButton2.Value - 1
End Sub

