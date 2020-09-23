VERSION 5.00
Object = "{34F681D0-3640-11CF-9294-00AA00B8A733}#1.0#0"; "danim.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text Automated Rotating & Moving Demo"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin DirectAnimationCtl.DAViewerControlWindowed DAViewerControlWindowed1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      OpaqueForHitDetect=   -1  'True
      UpdateInterval  =   0.033
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim metre, half, Font, clr, txtImg, pos, scl, xf, bgr
Set metre = DAViewerControlWindowed1.MeterLibrary

Set half = metre.DANumber(0.5)
Set clr = metre.ColorHslAnim(metre.Mul(metre.LocalTime, metre.DANumber(0.345)), half, half)

   'Text font
   Set Font = metre.Font("Comic Sans MS", 12, clr)
   'Text to be rotated
   Set txtImg = metre.StringImage("Text Rotating & Moving Demo", Font)
   
   Set pos = metre.Mul(metre.Sin(metre.LocalTime), metre.DANumber(0.05))
   'Zoom how many time(s)             *
   Set scl = metre.Add(metre.DANumber(1), metre.Abs(metre.Mul(metre.Sin(metre.LocalTime), metre.DANumber(3))))
   Set xf = metre.Compose2(metre.Translate2Anim(metre.DANumber(0), pos), _
                       metre.Scale2UniformAnim(scl))
   Set txtImg = txtImg.Transform(xf)
   'you can the degree of rotation here                     it is 30 as default
   Set bgr = metre.Rotate3RateDegrees(metre.Vector3(1, 1, 1), 30).ParallelTransform2
   Set txtImg = txtImg.Transform(bgr)
   
   DAViewerControlWindowed1.BackgroundImage = metre.SolidColorImage(metre.Black)
   DAViewerControlWindowed1.Image = txtImg
   DAViewerControlWindowed1.Start
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Delete this sub if you don't like a message box keep appearing when you exit
MsgBox "Please vote this code if you like. If you do not like this code, then give comments why you don't like. I will improve it. Just need your support...", , "Please..."
End Sub
