VERSION 5.00
Object = "*\AProject2.vbp"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project2.RMWx RMWx1 
      Left            =   855
      Top             =   315
      _extentx        =   873
      _extenty        =   873
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1755
      Top             =   405
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   15
      TabIndex        =   0
      Top             =   -90
      Width           =   4650
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Gathering weather information. Please Wait ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   383
         TabIndex        =   1
         Top             =   308
         Width           =   3915
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Form_Load()
i = 0
With Me
Frame1.Left = -10
Frame1.Top = -100
Frame1.Height = .Height + 125
Frame1.Width = .Width + 25
End With

Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
i = i + 1
If i = 2 Then
RMWx1.GetWeather Form1.Text9
Form1.Pic1.Picture = LoadPicture(App.Path & "\images\" & RMWx1.Image & ".gif")
Form1.Text7 = RMWx1.Temp
Form1.Text8 = RMWx1.Current
Form1.Text6 = RMWx1.FeelsLike
Form1.Text1 = RMWx1.UVIndex
Form1.Text2 = RMWx1.Wind
Form1.Text3 = RMWx1.Humidity
Form1.Text4 = RMWx1.Barometer
Form1.Text10 = RMWx1.CityName
Form1.Text5 = RMWx1.DewPoint
Form1.Text11 = RMWx1.Visibility
Timer1.Enabled = False
Unload Me
End If
End Sub
