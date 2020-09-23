VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin prjTestISControls.ISCombo ISCombo1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   300
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   661
      Caption         =   "ISCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ISCombo1.AddItem "yahoo.com"
    ISCombo1.AddItem "Hotmail"
    ISCombo1.AddItem "Terra"
    ISCombo1.AddItem "aol"
    ISCombo1.AddItem "PSC"
End Sub
