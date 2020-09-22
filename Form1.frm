VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password Module Demo...."
   ClientHeight    =   2568
   ClientLeft      =   36
   ClientTop       =   324
   ClientWidth     =   2508
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2568
   ScaleWidth      =   2508
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Caption         =   "&Set an new password"
      Height          =   252
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   2292
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Password Options"
      ForeColor       =   &H0000C000&
      Height          =   1452
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2292
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Write to File"
         ForeColor       =   &H0000C000&
         Height          =   252
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1572
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Write To Registry"
         ForeColor       =   &H0000C000&
         Height          =   252
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1572
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Encrypt Password"
         ForeColor       =   &H0000C000&
         Height          =   252
         Left            =   240
         MaskColor       =   &H0000C000&
         TabIndex        =   3
         Top             =   360
         Width           =   1692
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C000&
      Caption         =   "E&xit"
      Height          =   252
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1092
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "&Test"
      Height          =   252
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' *****************************************************************
' **
' **    THIS DEMO PROJECT IS MEANT TO ILLUSTRATE THE EASE OF USE
' **    OF THE MODULE CONTAINED WITHIN. PAY ATTEMTION TO THE WAY
' **    THE FUNCTIONS ARE USED IN THIS FORM TO GET AN IDEA OF HOW
' **    TO USE THEM IN YOUR PROJECT.
' **
' **    WHEN CHACNGING FROM SAVING IN THE REGISTRY OR  A FILE YOU
' **    MAY HAVE TO RESTART THE APPLICATION FOR THE NEW PASSWORD
' **    TO TAKE EFFECT.
' *****************************************************************




Private Sub Check1_Click()
    ' Encrypting?
    If Check1.Value = vbChecked Then
       pw.Encoded = True 'PW.Encoded = True
    Else
       pw.Encoded = False 'PW.Encoded = False
    End If
End Sub


Private Sub Command1_Click()
     Dim bCorrect As Boolean
     
     ' this is how you would get a users password
     ' and verify it. No forms to design, add, etc.
     
     ' Show Password Dialog, to verify users account
     GetPW_Dialog
     
     ' This function will return a boolean value
     ' sInput is the text the user types into the "Get Password"
     ' dialog
     bCorrect = CheckPW(sInput$)
     
     ' simple.... huh?
     If bCorrect Then
        MsgBox "Correct"
     Else
        MsgBox "Invalid"
     End If
        
End Sub

Private Sub Command2_Click()  ' EXIT
    Unload Me
End Sub


Private Sub Command3_Click()  ' SET\CHANGE PASSWORD
    GetNewPWord
End Sub


Private Sub Form_Load()
    ' make sure a saving palce has been set
    Option1.Value = True
       
    ' set defaults
    With pw
       .Save% = 1   '  registry
       .Encoded = False  ' don't encode
       ' specify where to store the password by setting
       ' the following variables
       .Path$ = App.Path & "\PW.dat"
       ' these values should be altered to hide the password better
       .AppName$ = "Demo"
       .Section$ = "PW_Section"
       .Key$ = "PW"
    End With
        
End Sub

Private Sub Option1_Click()
     ' Registry
     If Option1.Value = True Then
        pw.Save% = 1
     End If
       
End Sub

Private Sub Option2_Click()
     ' file
     If Option2.Value = True Then
        pw.Save% = 2
     End If
     
End Sub

