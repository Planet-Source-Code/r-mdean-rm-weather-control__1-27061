VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl RMWx 
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   99.939
   ScaleMode       =   0  'User
   ScaleWidth      =   154.545
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "RMWx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type cWx
cCityName As String
cCurrent As String
cTemp As String
cFeelsLike As String
cReport As String
cWind As String
cDewPoint As String
cUVIndex As String
cHumidity As String
cVisibality As String
cBarometer As String
cImageURL As String
cImage As String
End Type

Private cWx As cWx

Public Sub GetWeather(sZip As String)
Dim sStart, sEnd, sEnd1
Dim sTmp

DoEvents
sTmp = Inet1.OpenURL("http://www.weather.com/weather/local/" & sZip & ".htm")

On Error Resume Next
sStart = InStr(sTmp, "<!-- Insert City Name and Zip Code -->")
sEnd = InStrRev(sTmp, "<B>", Int(sStart)) + 3
cWx.cCityName = Right(Mid(sTmp, sEnd, sStart - sEnd), Len(Mid(sTmp, sEnd, sStart - sEnd)) - 1)

sStart = InStr(sStart, sTmp, "<!-- insert reported by and last updated info -->") + 49
sStart = InStr(sStart, sTmp, "reported")
sEnd1 = InStr(sStart, sTmp, "DT)") + 3
cWx.cReport = "as " & Replace(Mid(sTmp, sStart, sEnd1 - sStart), "&nbsp;", " ")

sStart = InStr(Int(sEnd1), sTmp, "<!-- insert wx icon --><img src=") + 33
sEnd = InStr(sStart, sTmp, "width") - 2
cWx.cImageURL = Mid(sTmp, sStart, sEnd - sStart)

sStart = InStr(Int(sEnd1), sTmp, "http://image.weather.com/web/common/wxicons/52/") + 47
sEnd = InStr(sStart, sTmp, "width") - 6
cWx.cImage = Mid(sTmp, sStart, sEnd - sStart)

sStart = InStr(Int(sEnd), sTmp, "<!-- insert forecast text -->") + 29
sEnd = InStr(sStart, sTmp, "</td>")
cWx.cCurrent = Mid(sTmp, sStart, sEnd - sStart)

sStart = InStr(Int(sEnd), sTmp, "<!-- insert current temp -->") + 28
sEnd = InStr(sStart, sTmp, "</B>")
cWx.cTemp = Mid(sTmp, sStart, sEnd - sStart)

sStart = InStr(Int(sEnd), sTmp, "<!-- insert feels like temp -->") + 31
sEnd = InStr(sStart, sTmp, "</font>")
cWx.cFeelsLike = Mid(sTmp, sStart, sEnd - sStart)

sStart = InStr(Int(sEnd), sTmp, "<!-- insert UV number -->") + 25
sEnd = InStr(sStart, sTmp, "</td>")
cWx.cUVIndex = Replace(Mid(sTmp, sStart, sEnd - sStart), "&nbsp;", " ")

sStart = InStr(Int(sEnd), sTmp, "<!-- insert wind information -->") + 32
sEnd = InStr(sStart, sTmp, "</td>")
cWx.cWind = Replace(Mid(sTmp, sStart, sEnd - sStart), "&nbsp;", " ")

sStart = InStr(Int(sEnd), sTmp, "<!-- insert dew point -->") + 25
sEnd = InStr(sStart, sTmp, "</td>")
cWx.cDewPoint = Replace(Mid(sTmp, sStart, sEnd - sStart), "&nbsp;", " ")

sStart = InStr(Int(sEnd), sTmp, "<!-- insert humidity -->") + 24
sEnd = InStr(sStart, sTmp, "</td>")
cWx.cHumidity = Replace(Mid(sTmp, sStart, sEnd - sStart), "&nbsp;", " ")

sStart = InStr(Int(sEnd), sTmp, "<!-- insert visibility -->") + 26
sEnd = InStr(sStart, sTmp, "</td>")
cWx.cVisibality = Replace(Mid(sTmp, sStart, sEnd - sStart), "&nbsp;", " ")

sStart = InStr(Int(sEnd), sTmp, "<!-- insert barometer information -->") + 37
sEnd = InStr(sStart, sTmp, "</td>")
cWx.cBarometer = Replace(Mid(sTmp, sStart, sEnd - sStart), "&nbsp;", " ")
Unload Form2

End Sub
Function CityName() As String
CityName = cWx.cCityName
End Function
Function Report() As String
Report = cWx.cReport
End Function
'This is for later use
Private Function ImageURL() As String
ImageURL = cWx.cImageURL
End Function
Function Image() As String
Image = cWx.cImage
End Function
Function Current() As String
Current = cWx.cCurrent
End Function
Function Temp() As String
Temp = cWx.cTemp & " " & Chr(176) & "F"
End Function
Function FeelsLike() As String
FeelsLike = Replace(cWx.cFeelsLike, "&nbsp;&deg;", " " & Chr(176))
End Function
Function UVIndex() As String
UVIndex = cWx.cUVIndex
End Function
Function Wind() As String
Wind = cWx.cWind
End Function
Function DewPoint() As String
DewPoint = Replace(cWx.cDewPoint, "&deg;", Chr(176))
End Function
Function Humidity() As String
Humidity = cWx.cHumidity
End Function
Function Visibility() As String
Visibility = cWx.cVisibality
End Function
Function Barometer() As String
Barometer = cWx.cBarometer
End Function
Private Sub UserControl_Initialize()
UserControl.Picture = LoadPicture(App.Path & "\rmwx.bmp")
End Sub
Private Sub UserControl_Resize()
With UserControl
.Height = "495"
.Width = "495"
End With
End Sub
