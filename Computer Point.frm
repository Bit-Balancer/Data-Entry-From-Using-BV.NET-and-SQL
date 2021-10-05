VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form main_frm 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Computer Point"
   ClientHeight    =   8988
   ClientLeft      =   108
   ClientTop       =   732
   ClientWidth     =   13248
   FillStyle       =   0  'Solid
   Icon            =   "Computer Point.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8988
   ScaleWidth      =   13248
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Txt4 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      TabIndex        =   30
      Top             =   4320
      Width           =   3372
   End
   Begin VB.TextBox Txt5 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      TabIndex        =   29
      Top             =   5040
      Width           =   3372
   End
   Begin VB.CommandButton Findbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   28
      Top             =   1440
      Width           =   972
   End
   Begin VB.CommandButton Previousbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   2160
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   26
      Top             =   8280
      Width           =   1572
   End
   Begin VB.CommandButton Newbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   4080
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   25
      Top             =   8280
      Width           =   1092
   End
   Begin VB.CommandButton Savebtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   24
      Top             =   8280
      Width           =   1092
   End
   Begin VB.CommandButton Updatebtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   23
      Top             =   8280
      Width           =   1332
   End
   Begin VB.CommandButton Deletebtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   22
      Top             =   8280
      Width           =   1332
   End
   Begin VB.CommandButton Nextbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   10080
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   21
      Top             =   8280
      Width           =   1092
   End
   Begin VB.CommandButton Lastbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   20
      Top             =   8280
      Width           =   1092
   End
   Begin VB.CommandButton AddPicturebtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Picture"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   19
      Top             =   4680
      Width           =   2412
   End
   Begin VB.CommandButton Firstbtn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   720
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   18
      Top             =   8280
      Width           =   1092
   End
   Begin VB.TextBox Txt7 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      TabIndex        =   17
      Top             =   7200
      Width           =   3372
   End
   Begin VB.PictureBox Picture1 
      Height          =   2412
      Left            =   8520
      ScaleHeight     =   2364
      ScaleWidth      =   2484
      TabIndex        =   16
      Top             =   2040
      Width           =   2532
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   492
      Left            =   3360
      TabIndex        =   15
      Top             =   6480
      Width           =   3492
      _ExtentX        =   6160
      _ExtentY        =   868
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   78905345
      CurrentDate     =   42765
   End
   Begin VB.TextBox Txt6 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      TabIndex        =   14
      Top             =   5760
      Width           =   3372
   End
   Begin VB.TextBox Txt3 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      TabIndex        =   13
      Top             =   3600
      Width           =   3372
   End
   Begin VB.OptionButton Op2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   5040
      TabIndex        =   12
      Top             =   2880
      Width           =   1572
   End
   Begin VB.OptionButton Op1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   3360
      TabIndex        =   11
      Top             =   2880
      Width           =   1212
   End
   Begin VB.TextBox Txt1 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      TabIndex        =   10
      Top             =   1440
      Width           =   3372
   End
   Begin VB.TextBox Txt2 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   16.2
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   3360
      TabIndex        =   9
      Top             =   2160
      Width           =   3372
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Computer Point"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   25.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   4680
      TabIndex        =   27
      Top             =   480
      Width           =   3852
   End
   Begin VB.Label LB9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Phone No."
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   1560
      TabIndex        =   8
      Top             =   7200
      Width           =   1332
   End
   Begin VB.Label LB8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Joining Date"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   1200
      TabIndex        =   7
      Top             =   6480
      Width           =   1692
   End
   Begin VB.Label LB6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Course Teacher"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   840
      TabIndex        =   6
      Top             =   5040
      Width           =   2052
   End
   Begin VB.Label LB7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Course Fee"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   1440
      TabIndex        =   5
      Top             =   5760
      Width           =   1452
   End
   Begin VB.Label LB5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Course Type"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   1320
      TabIndex        =   4
      Top             =   4320
      Width           =   1572
   End
   Begin VB.Label LB4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Guardian Name"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   840
      TabIndex        =   3
      Top             =   3600
      Width           =   2052
   End
   Begin VB.Label LB3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   1920
      TabIndex        =   2
      Top             =   2880
      Width           =   972
   End
   Begin VB.Label LB2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   960
      TabIndex        =   1
      Top             =   2160
      Width           =   1932
   End
   Begin VB.Label LB1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "@Adobe Gothic Std B"
         Size            =   13.2
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   372
      Left            =   2280
      TabIndex        =   0
      Top             =   1440
      Width           =   492
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnufilesave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnufilecancel 
         Caption         =   "&Cancel"
      End
      Begin VB.Menu mnufileexit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "main_frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub AddPicturebtn_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*.jpg"
str = CommonDialog1.FileName
Picture1.Picture = LoadPicture(str)
End Sub
Sub Relode()
rs.Close
rs.Open "Select * From Table1", con, adOpenStatic, adLockPessimistic
End Sub
Sub Display()
Txt1.Text = rs!Roll
Txt2.Text = rs!Student_Name
If rs!Gender = "Male" Then
Op1.Value = True
Else
Op2.Value = True
End If
Txt3.Text = rs!Guardian_Name
Txt4.Text = rs!Course_Type
Txt5.Text = rs!Course_Teacher
Txt6.Text = rs!Course_Fee
DTPicker1.Value = rs!Joining_Date
Txt7.Text = rs!Phone_No
End Sub
Sub Clear()
Txt1.Text = ""
Txt2.Text = ""
Op1.Value = False
Op2.Value = False
Txt3.Text = ""
Txt4.Text = ""
Txt5.Text = ""
Txt6.Text = ""
DTPicker1.Value = "1/1/2000"
Txt7.Text = ""
End Sub
Sub refreshdata()
rs.Close
rs.Open "Select * from Table1", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveNext
Display
End If
End Sub
Sub MoveLast()
rs.Close
rs.Open "Select * from Table1", con, adOpenStatic, adLockPessimistic
If Not rs.EOF Then
rs.MoveLast
Display
End If
End Sub
Private Sub Deletebtn_Click()
confirm = MsgBox("Do You Want Delet Profile........", vbYesNo + vbCritical, "Deleted....")
If confirm = vbYes Then
rs.Delete adAffectCurrent
MsgBox "Deleted Successfully", vbInformation, "Message"
rs.Update
refreshdata
Else
MsgBox "Not Deleted......", vbInformation, "Message"
End If
End Sub

Private Sub Findbtn_Click()
rs.Close
rs.Open "Select * from Table1 where Roll='" + Txt1.Text + "'", con, adOpenDynamic, adLockPessimistic
If Not rs.EOF Then
Display
Relode
Else
MsgBox "Record Not Found.....!", vbInformation
End If
End Sub

Private Sub Firstbtn_Click()
rs.MoveFirst
Display
End Sub
Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\RHIVED\Documents\Student_Record_Projects.mdb;Persist Security Info=False"
rs.Open "Select * from Table1", con, adOpenDynamic, adLockPessimistic
MoveLast
Display
End Sub
Private Sub Lastbtn_Click()
rs.MoveLast
Display
End Sub

Private Sub mnufilecancel_Click()
Relode
End Sub

Private Sub mnufileexit_Click()
End
End Sub

Private Sub mnufilenew_Click()
rs.AddNew
Clear
End Sub

Private Sub mnufilesave_Click()
If rs.Fields("Roll").Value = Txt1.Text Then
MsgBox "This Record Is Saved Already........!You Can Update That File!!!!!", vbExclamation
Else
If Txt2.Text = "" Or Txt2.Text = "" Or Txt3.Text = "" Or Txt4.Text = "" Or Txt5.Text = "" Or Txt6.Text = "" Or Txt7.Text = "" Then
MsgBox "Please Fill All Fields......!!!!!!", vbExclamation
Else
If DTPicker1.Value = "1/1/2000" Then
MsgBox "Please Enter Joining Date..........?", vbExclamation
Else
rs.Fields("Student_Name").Value = Txt2.Text
If Op1.Value = True Then
rs.Fields("Gender").Value = Op1.Caption
Else
rs.Fields("Gender").Value = Op2.Caption
End If
rs.Fields("Guardian_Name").Value = Txt3.Text
rs.Fields("Course_Type").Value = Txt4.Text
rs.Fields("Course_Teacher").Value = Txt5.Text
rs.Fields("Course_Fee").Value = Txt6.Text
rs.Fields("Joining_Date").Value = DTPicker1
rs.Fields("Phone_No").Value = Txt7.Text
MsgBox "Data Is Saved Successfully............", vbInformation
rs.Update
Txt1.Text = rs!Roll
MoveLast
End If
End If
End If
End Sub

Private Sub Newbtn_Click()
rs.AddNew
Clear
End Sub
Private Sub Nextbtn_Click()
rs.MoveNext
If Not rs.EOF Then
Display
Else
rs.MoveFirst
Display
End If
End Sub
Private Sub Previousbtn_Click()
rs.MovePrevious
If rs.BOF Then
rs.MoveLast
Display
Else
Display
End If
End Sub

Private Sub Savebtn_Click()
If rs.Fields("Roll").Value = Txt1.Text Then
MsgBox "This Record Is Saved Already........!You Can Update That File!!!!!", vbExclamation
Else
If Txt2.Text = "" Or Txt2.Text = "" Or Txt3.Text = "" Or Txt4.Text = "" Or Txt5.Text = "" Or Txt6.Text = "" Or Txt7.Text = "" Then
MsgBox "Please Fill All Fields......!!!!!!", vbExclamation
Else
If DTPicker1.Value = "1/1/2000" Then
MsgBox "Please Enter Joining Date..........?", vbExclamation
Else
rs.Fields("Student_Name").Value = Txt2.Text
If Op1.Value = True Then
rs.Fields("Gender").Value = Op1.Caption
Else
rs.Fields("Gender").Value = Op2.Caption
End If
rs.Fields("Guardian_Name").Value = Txt3.Text
rs.Fields("Course_Type").Value = Txt4.Text
rs.Fields("Course_Teacher").Value = Txt5.Text
rs.Fields("Course_Fee").Value = Txt6.Text
rs.Fields("Joining_Date").Value = DTPicker1
rs.Fields("Phone_No").Value = Txt7.Text
MsgBox "Data Is Saved Successfully............", vbInformation
rs.Update
Txt1.Text = rs!Roll
MoveLast
End If
End If
End If
End Sub

Private Sub Txt1_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyDecPt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Txt2_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
Beep
Else
Exit Sub
End If
End Sub

Private Sub Txt3_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
Beep
Else
Exit Sub
End If
End Sub

Private Sub Txt4_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
Beep
Else
Exit Sub
End If
End Sub

Private Sub Txt5_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
KeyAscii = 0
Beep
Else
Exit Sub
End If
End Sub

Private Sub Txt6_Keypress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyDecPt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Txt7_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyDecPt Or KeyAscii = vbKeyBack Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Updatebtn_Click()
If Not rs.Fields("Roll").Value = Txt1.Text Then
MsgBox "Access Denied.....!Roll Cannot Be Changed!!!", vbExclamation
Else
rs.Fields("Student_Name").Value = Txt2.Text
If Op1.Value = True Then
rs.Fields("Gender").Value = Op1.Caption
Else
rs.Fields("Gender").Value = Op2.Caption
End If
rs.Fields("Guardian_Name").Value = Txt3.Text
rs.Fields("Course_Type").Value = Txt4.Text
rs.Fields("Course_Teacher").Value = Txt5.Text
rs.Fields("Course_Fee").Value = Txt6.Text
rs.Fields("Joining_Date").Value = DTPicker1
rs.Fields("Phone_No").Value = Txt7.Text
MsgBox "Data Update Successfully............", vbInformation
rs.Update
End If
End Sub
