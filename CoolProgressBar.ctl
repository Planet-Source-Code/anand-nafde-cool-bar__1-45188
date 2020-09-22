VERSION 5.00
Begin VB.UserControl CoolProgressBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox pBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   600
      ScaleHeight     =   315
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Shape border 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   3735
   End
End
Attribute VB_Name = "CoolProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'How to use ?
'
'1. Compile this code to get an ActiveX control
'2. Place this control on your form
'3. Set the minimum / macimum values for this progress bar
'4. choose the colors using the resp. properties
'5. call the Value method by supplying some value
'6. Value supplied should be in the min/max range
'
'
'I am open for your kind suggestions how I can improve this
'I am reachable at <anandnafde@rediffmail.com>
'Best,
'Anand Nafde
'
'
'Global Declarations
Event Error(Description As String, AddInfo As String)

Dim Fore_Color As Long
Dim Minimum_Value As Integer, Maximum_Value As Integer
Dim Display_Percent As Boolean, Show_Border As Boolean

'All properties to change colors /text etc...
'
Property Let Border_Color(ByVal color As Long)
On Error GoTo Error
  border.BackColor = color
  
Exit Property
Error:
  RaiseEvent Error("Cannot set Border Color", "Supplied number is not a valid color value")
End Property
Property Let Back_Color(ByVal color As Long)
On Error GoTo Error
  pBar.BackColor = color
  
Exit Property
Error:
  RaiseEvent Error("Cannot set Back Color", "Supplied number is not a valid color value")
End Property
Property Let ForeColor(ByVal color As Long)
On Error GoTo Error
  Fore_Color = color
  
Exit Property
Error:
  RaiseEvent Error("Cannot set Fore Color", "Supplied number is not a valid color value")
End Property
Property Let Text_Color(ByVal color As Long)
On Error GoTo Error
  pBar.ForeColor = color
  
Exit Property
Error:
  RaiseEvent Error("Cannot set Text Color", "Supplied number is not a valid color value")
End Property
Property Let Text_Font(ByVal font As String)
On Error GoTo Error
  pBar.FontName = font
  
Exit Property
Error:
  RaiseEvent Error("Cannot set the font", "Supplied font is not a valid or is not available on this machine")
End Property
Property Let Text_Size(ByVal size As Integer)
On Error GoTo Error
  pBar.FontSize = size
  
Exit Property
Error:
  RaiseEvent Error("Cannot set the font size", "Supplied size is either too large or too small for the selected font")
End Property
Property Let MinimumValue(ByVal minimum As Integer)
On Error GoTo Error
  Minimum_Value = minimum
  
Exit Property
Error:
  RaiseEvent Error("Cannot set minimum value", "minimum value should be between 0 and 32767")
End Property
Property Let MaximumValue(ByVal maximum As Integer)
On Error GoTo Error
If maximum <= Minimum_Value Then
  RaiseEvent Error("Cannot set Maximum value", "Value must be greater than minimum value")
  Exit Property
End If
  Maximum_Value = maximum
  
Exit Property
Error:
  RaiseEvent Error("Cannot set Maximum value", "Maximum value should be between 0 and 32767")
End Property
Property Let DisplayPercentage(ByVal value As Boolean)
  Display_Percent = value
End Property
Property Let ShowBorder(ByVal value As Boolean)
  Show_Border = value
End Property


Private Sub UserControl_Initialize()
'initialize important variables here
Fore_Color = vbRed
Minimum_Value = 0
Maximum_Value = 100
Display_Percent = False
End Sub

Private Sub UserControl_Resize()
'Make sure background is not visible
If Height < 150 Then
  Height = 150
End If
If Width < 200 Then
  Width = 200
End If

If Height > 600 Then
  Height = 600
End If


If Show_Border Then
  border.Left = 0
  border.Top = 0
  border.Height = Height
  border.Width = Width
  
  pBar.Left = 50
  pBar.Top = 50
  pBar.Height = border.Height - 100
  pBar.Width = border.Width - 100

Else
  border.Visible = False
  pBar.Left = 0
  pBar.Top = 0
  pBar.Height = Height
  pBar.Width = Width
End If
End Sub

Public Sub value(ByVal value As Integer)
'This is the actual Sub where the display will take place
On Error GoTo Error

If (value > Maximum_Value) Or (value < Minimum_Value) Then
  RaiseEvent Error("Progress bar value out of range", "Error Number: " & Err.Number)
  Exit Sub
End If

'Clear the display area
pBar.Cls
'Scale the display area according to the min/max values
pBar.Scale (Minimum_Value, 1)-(Maximum_Value, 0)
'Draw the progress line here
pBar.Line (Minimum_Value, 0)-(value, 1), Fore_Color, BF

'print the percentage here
If Display_Percent Then
  pBar.CurrentX = (Maximum_Value / 2) - pBar.TextWidth(value & " %")
  pBar.CurrentY = 0.85
  pBar.Print (value & " %")
End If

Exit Sub
Error:
  RaiseEvent Error("An unhandled exception caught in displaying the progress value: " & vbCrLf & Err.Description, "Error Number: " & Err.Number)
End Sub
