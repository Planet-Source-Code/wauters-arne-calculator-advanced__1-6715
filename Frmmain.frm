VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Wetenschappelijke rekenmachine"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Memo"
      Height          =   2055
      Left            =   6240
      TabIndex        =   42
      Top             =   720
      Width           =   735
      Begin VB.CommandButton cmdMplus 
         Caption         =   "M+"
         Height          =   435
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Width           =   510
      End
      Begin VB.CommandButton cmdMR 
         Caption         =   "MR"
         Height          =   435
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   510
      End
      Begin VB.CommandButton cmdMC 
         Caption         =   "MC"
         Height          =   435
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "View"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1680
         Width           =   405
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Wetenschappelijk"
      Height          =   2175
      Left            =   3600
      TabIndex        =   25
      Top             =   600
      Width           =   2535
      Begin VB.CommandButton CmdSIN 
         Caption         =   "Sin"
         Height          =   375
         Left            =   1320
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdCOS 
         Caption         =   "Cos"
         Height          =   375
         Left            =   1320
         TabIndex        =   39
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdTAN 
         Caption         =   "Tan"
         Height          =   375
         Left            =   1320
         TabIndex        =   38
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdAtn 
         Caption         =   "Atn"
         Height          =   375
         Left            =   1320
         TabIndex        =   37
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdSQR 
         Caption         =   "SQR"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdX2 
         Caption         =   "X^2"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdXdeel1 
         Caption         =   "X/1"
         Height          =   375
         Left            =   720
         TabIndex        =   34
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdX3 
         Caption         =   "X^3"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdPI 
         Caption         =   "PI"
         Height          =   375
         Left            =   720
         TabIndex        =   32
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdExp 
         Caption         =   "X^?"
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmdprocent 
         Caption         =   "%"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdLog 
         Caption         =   "Log"
         Height          =   375
         Left            =   1920
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdIn 
         Caption         =   "In"
         Height          =   375
         Left            =   1920
         TabIndex        =   28
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton CmdExp2 
         Caption         =   "Exp"
         Height          =   375
         Left            =   1920
         TabIndex        =   27
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton CmdFact 
         Caption         =   "X!"
         Height          =   375
         Left            =   720
         TabIndex        =   26
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "A - S"
         Height          =   255
         Left            =   2000
         TabIndex        =   41
         Top             =   1755
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   2160
      TabIndex        =   16
      Top             =   600
      Width           =   1335
      Begin VB.CommandButton Cmddelen 
         Caption         =   "/"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmdmaal 
         Caption         =   "X"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Cmdmin 
         Caption         =   "-"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Cmdplus 
         Caption         =   "+"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdCal 
         Caption         =   "="
         Height          =   375
         Left            =   720
         TabIndex        =   20
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdCE 
         Caption         =   "CE"
         Height          =   375
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton CmdC 
         Caption         =   "C"
         Height          =   375
         Left            =   720
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Cmdterug 
         Caption         =   "<-"
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   1200
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1935
      Begin VB.CommandButton Cmd7 
         Caption         =   "7"
         Height          =   375
         Left            =   120
         MaskColor       =   &H8000000F&
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmd4 
         Caption         =   "4"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Cmd1 
         Caption         =   "1"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Cmd8 
         Caption         =   "8"
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmd5 
         Caption         =   "5"
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "2"
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton cmd9 
         Caption         =   "9"
         Height          =   375
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Cmd6 
         Caption         =   "6"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Cmd3 
         Caption         =   "3"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Cmd0 
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton CmdPosORneg 
         Caption         =   "+/-"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Cmdcomma 
         Caption         =   ","
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1680
         Width           =   495
      End
   End
   Begin VB.CommandButton CmdPaste 
      Caption         =   "Paste"
      Height          =   255
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton CmdCopy 
      Caption         =   "Copy "
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtDisplay 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer
Dim getal2 As Single
Dim getal1 As Single
Dim gh As Double
Dim bewerk As String
Public Function Toontext(getal As String)
If txtDisplay <> "0" Then
    If Left(txtDisplay.Text, 1) <> "-" Then
        txtDisplay.Text = txtDisplay.Text + getal
    Else
        If txtDisplay.Text = "-0" Then
            txtDisplay.Text = "-" + getal
        Else
            txtDisplay.Text = txtDisplay.Text + getal
        End If
   End If
Else
txtDisplay.Text = getal
End If
End Function
Private Sub Cmd0_Click()
Toontext (Cmd0.Caption)
End Sub
Private Sub Cmd1_Click()
Toontext (Cmd1.Caption)
End Sub
Private Sub cmd2_Click()
Toontext (cmd2.Caption)
End Sub

Private Sub Cmd3_Click()
Toontext (Cmd3.Caption)
End Sub

Private Sub cmd4_Click()
Toontext (cmd4.Caption)
End Sub

Private Sub Cmd5_Click()
Toontext (Cmd5.Caption)
End Sub

Private Sub Cmd6_Click()
Toontext (Cmd6.Caption)
End Sub
Private Sub Cmd7_Click()
Toontext (Cmd7.Caption)
End Sub
Private Sub Cmd8_Click()
Toontext (Cmd8.Caption)
End Sub
Private Sub cmd9_Click()
Toontext (cmd9.Caption)
End Sub
Public Function bewerking(bewerkingetje As String)
Select Case bewerkingetje
    Case "+"
        getal1 = Val(txtDisplay.Text)
        bewerk = "+"
        txtDisplay.Text = ""
    Case "-"
        getal1 = Val(txtDisplay.Text)
        bewerk = "-"
        txtDisplay.Text = ""
    Case "/"
        getal1 = Val(txtDisplay.Text)
        bewerk = "/"
        txtDisplay.Text = ""
    Case "X"
        getal1 = Val(txtDisplay.Text)
        bewerk = "X"
        txtDisplay.Text = ""
    'Case "%"
        'getal1 = Val(txtDisplay.Text)
        'bewerk = "%"
        'txtDisplay.Text = ""
End Select
End Function
Private Sub CmdCal_Click()
getal2 = Val(txtDisplay.Text)
werkuit getal1, bewerk, getal2
End Sub
Public Function werkuit(getalleke As Single, bewerk As String, getalleke2 As Single)
Select Case bewerk
    Case "+"
        txtDisplay.Text = getalleke + getalleke2
    Case "-"
        txtDisplay.Text = getalleke - getalleke2
    Case "/"
        If getalleke2 = 0 Then
            MsgBox "wie deelt door nul is een sul", vbOKOnly, "Fout!"
        Else
            txtDisplay.Text = getalleke / getalleke2
        End If
    Case "X"
        txtDisplay.Text = getalleke * getalleke2
    'Case "%"
     'txtDisplay.Text = getalleke * getalleke / 100 * getalleke2
End Select
End Function
Private Sub Cmdcomma_Click()
If InStr(txtDisplay.Text, ".") Then
    MsgBox "U heeft reeds een komma in dit getal", vbOKOnly, "Fout!"
    Exit Sub
Else
    txtDisplay.Text = txtDisplay.Text + "."
End If
End Sub
Private Sub CmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtDisplay.Text
End Sub
Private Sub CmdCOS_Click()
txtDisplay.Text = Cos(Val(txtDisplay.Text))
End Sub
Private Sub Cmddelen_Click()
bewerking (Cmddelen.Caption)
End Sub
Private Sub CmdExp2_Click()
txtDisplay.Text = Exp(Val(txtDisplay.Text))
End Sub
Private Sub CmdFact_Click()
 txtDisplay.Text = fact(Val(txtDisplay.Text))
End Sub
Private Sub CmdIn_Click()
If Val(txtDisplay.Text) > 0 Then
    txtDisplay.Text = Log(Val(txtDisplay.Text))
Else
    MsgBox "Het huidige getal is negatief", vbOKOnly, "Fout!"
End If
End Sub
Private Sub CmdLog_Click()
If Val(txtDisplay.Text) > 0 Then
    txtDisplay.Text = Log(Val(txtDisplay.Text)) / Log(10)
Else
    MsgBox "Het huidige getal is negatief", vbOKOnly, "Fout!"
End If
End Sub
Private Sub Cmdmaal_Click()
bewerking (Cmdmaal.Caption)
End Sub

Private Sub cmdMC_Click()
gh = 0
End Sub

Private Sub Cmdmin_Click()
bewerking (Cmdmin.Caption)
End Sub

Private Sub cmdMplus_Click()
gh = txtDisplay.Text
End Sub

Private Sub cmdMR_Click()
txtDisplay.Text = gh
End Sub

Private Sub CmdPaste_Click()
txtDisplay.Text = ""
txtDisplay.Text = Clipboard.GetText()
End Sub
Private Sub CmdPI_Click()
txtDisplay.Text = 3.141592654
End Sub
Private Sub Cmdplus_Click()
bewerking (Cmdplus.Caption)
End Sub

Private Sub CmdPosORneg_Click()
If Left(txtDisplay.Text, 1) <> "-" Then
    txtDisplay.Text = "-" + txtDisplay.Text
Else
   tussenstring = Mid$(txtDisplay.Text, 2, Len(txtDisplay.Text))
   txtDisplay.Text = tussenstring
End If
End Sub

Private Sub Cmdprocent_Click()
'bewerking ("%")
MsgBox "Deze Functie is momenteel onder constructie :P", vbOKOnly, "Oops!"
End Sub

Private Sub CmdSIN_Click()
txtDisplay.Text = Sin(Val(txtDisplay.Text))
End Sub
Private Sub CmdSQR_Click()
On Error Resume Next
antwoord = Sqr(Val(txtDisplay.Text))
txtDisplay.Text = antwoord
End Sub
Private Sub CmdCE_Click()
getal1 = 0
getal2 = 0
bewerk = ""
txtDisplay.Text = "0"
End Sub

Private Sub CmdC_Click()
getal2 = 0
txtDisplay.Text = "0"
End Sub

Private Sub CmdTAN_Click()
txtDisplay.Text = Tan(Val(txtDisplay.Text))
End Sub

Private Sub Cmdterug_Click()
If Len(txtDisplay.Text) > 0 Then
    txtDisplay.Text = Mid$(txtDisplay.Text, 1, Len(txtDisplay.Text) - 1)
Else
    Beep
End If
End Sub
Private Sub CmdATN_Click()
On Error Resume Next
X = Atn(txtDisplay.Text)
txtDisplay.Text = X
End Sub
Private Sub CmdX2_Click()
On Error Resume Next
txtDisplay.Text = Val(txtDisplay.Text) * Val(txtDisplay.Text)
End Sub

Private Sub CmdXdeel1_Click()
tussen = txtDisplay.Text
If Val(tussen) <> 0 Then
                    txtDisplay.Text = (1 / tussen)
                Else
                     MsgBox "10 x schrijven, wie deelt door nul is een sul!", vbOKOnly, "Fout!"
                End If
End Sub
Private Sub CmdX3_Click()
On Error Resume Next
txtDisplay.Text = txtDisplay.Text ^ 3
End Sub
Private Sub CmdExp_Click()
fm = InputBox("Geef het exponent in", "Geef exponent")
If fm = "" Then fm = 1
txtDisplay.Text = Val(txtDisplay.Text) ^ fm
End Sub
Function fact(num As Long) As Long
    If (num < 0 Or num = 0) Then
        MsgBox ("Getal is negatief")
        fact = num
    Else
            opnieuw = 1
            While (num > 0)
                opnieuw = opnieuw * num
                num = num - 1
            Wend
            fact = opnieuw
    End If
End Function
Private Sub Label1_Click()
Form2.Show

End Sub

Private Sub Label2_Click()
MsgBox gh, vbOKOnly, "Inhoud van het Geheugen"
End Sub
