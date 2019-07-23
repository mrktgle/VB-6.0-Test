VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "FormUI"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11820
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox txtBday 
      Height          =   420
      Left            =   1800
      TabIndex        =   11
      Text            =   "mm/dd/yyyy"
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   1800
      TabIndex        =   8
      Top             =   3000
      Width           =   2895
      Begin VB.OptionButton radioFemale 
         Caption         =   "Female"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton radioMale 
         Caption         =   "Male"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox txtPhone 
      Height          =   420
      Left            =   2400
      TabIndex        =   7
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtEmail 
      Height          =   420
      Left            =   2280
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   420
      Left            =   1440
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Birthday:"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Gender:"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Phone Number:"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Email Address:"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FullName As String
Dim EmailAddress As String
Dim PhoneNumber As String
Dim Gender As String
Dim Birthdate As String

Public Sub Validations()
    If txtName.Text = "" Then MsgBox ("Please Enter your Name")
    If txtEmail.Text = "" Then MsgBox ("Please Enter your Email")
    If txtPhone.Text = "" Then MsgBox ("Please Enter your Phone")
    If txtBday.Text = "mm/dd/yyyy" Then MsgBox ("Please Enter your Bday")
End Sub

Private Sub btnSave_Click()
If Gender = "" Then Gender = "Please Select your Gender."
FullName = txtName.Text
EmailAddress = txtEmail.Text
PhoneNumber = txtPhone.Text
Birthdate = txtBday.Text
'MsgBox (FullName + EmailAddress + PhoneNumber + Gender + Birthdate)
Validations
End Sub

Private Sub radioFemale_Click()
Gender = "Female"
End Sub

Private Sub radioMale_Click()
Gender = "Male"
End Sub
