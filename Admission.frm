VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Admission 
   Caption         =   "Welcome to Student Admission Form"
   ClientHeight    =   8268.001
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   12900
   OleObjectBlob   =   "Admission.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Admission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'Start Code for upload photo and this code will be always placed on top
'===========================================================================
Dim fpath As String

Private Sub CommandButton10_Click()
Unload Me
Admission.Show
End Sub

'=============================================================================================================================
'Start code for select only from ComboBox (Copy the source code from https://www.omgstudy.com) and use for TextBox or ComboBox.
'=============================================================================================================================







'========================================================
'End code for select only from ComboBox
'========================================================

Private Sub CommandButton6_Click()
'===========================================================================
'Start Code for upload photo (This code use for upload button)
'===========================================================================
On Error Resume Next
Dim x As Integer
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
x = Application.FileDialog(msoFileDialogOpen).Show
If x <> 0 Then
fpath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
Image1.Picture = LoadPicture(fpath)
Image1.PictureSizeMode = 1
End If
End Sub
'===========================================================================
'Start Code for upload photo
'===========================================================================

Private Sub CommandButton7_Click()
'=========================================================================================================================
'start code for required field (Copy the source code from https://www.omgstudy.com) and use for TextBox or ComboBox.
'=========================================================================================================================
'==================================
'Required field for Student Name
'==================================
On Error Resume Next
If TextBox2.Text = "" Then
Cancel = 1
MsgBox "Please enter the student name", vbOKOnly, "Learn with OMGStudy.Com"
TextBox2.SetFocus
Exit Sub
End If

'==================================
'Required field for Student Name
'==================================
On Error Resume Next
If TextBox3.Text = "" Then
Cancel = 1
MsgBox "Please enter the father name", vbOKOnly, "Learn with OMGStudy.Com"
TextBox3.SetFocus
Exit Sub
End If

'==================================
'Required field for Student Name
'==================================
On Error Resume Next
If TextBox4.Text = "" Then
Cancel = 1
MsgBox "Please enter the mother name", vbOKOnly, "Learn with OMGStudy.Com"
TextBox4.SetFocus
Exit Sub
End If









'=============================================================
'End code for required field
'=============================================================
'========================================================================
'start Code for enter data from Textbox into Database Worksheet
'========================================================================
Dim x As Long
Dim y As Worksheet
Set y = Sheets("Database")
x = y.Range("B" & Rows.Count).End(xlUp).Row
With y
.Cells(x + 1, "B").Value = TextBox2.Text
.Cells(x + 1, "C").Value = TextBox3.Text
.Cells(x + 1, "D").Value = TextBox4.Text
.Cells(x + 1, "E").Value = TextBox5.Text
.Cells(x + 1, "F").Value = TextBox6.Text
.Cells(x + 1, "G").Value = TextBox7.Text
.Cells(x + 1, "H").Value = TextBox8.Text
.Cells(x + 1, "L").Value = TextBox9.Text
.Cells(x + 1, "M").Value = TextBox10.Text

.Cells(x + 1, "I").Value = ComboBox4.Text
.Cells(x + 1, "J").Value = ComboBox2.Text
.Cells(x + 1, "K").Value = ComboBox3.Text
.Cells(x + 1, "N").Value = ComboBox5.Text

End With
'================================================================
'start code for send uploaded photo in Photo folder of C Drive
'================================================================
On Error Resume Next
Dim i As String
i = TextBox1.Text
FileCopy fpath, "C:\Photo\" & i & ".jpg"
Me.TextBox1.Text = ""
'======================================================================================
'start code for clear Textbox after submit data from Textbox into Database Worksheet
'======================================================================================
TextBox2.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Text = ""
TextBox6.Text = ""
TextBox7.Text = ""
TextBox8.Text = ""
TextBox9.Text = ""
TextBox10.Text = ""

ComboBox4.Text = ""
ComboBox2.Text = ""
ComboBox3.Text = ""
ComboBox5.Text = ""

'=================================================================
'start code for clear uploaded photo after submit
'=================================================================
Image1.Picture = Nothing
'=================================================================
'start code for show message after submit data
'=================================================================
MsgBox "Admission successfully. Now click on the reset button for new admission.", vbOKOnly, "Learn with OMGStudy.Com"
End Sub


Private Sub CommandButton8_Click()
'============================================
'start code for Search Student Information
'============================================
On Error Resume Next
Dim x As Long
Dim y As Long
x = Sheets("Database").Range("B" & Rows.Count).End(xlUp).Row
For y = 2 To x
If Sheets("Database").Cells(y, 1).Text = TextBox11.Value Then
TextBox1.Text = Sheets("Database").Cells(y, 1)
TextBox2.Text = Sheets("Database").Cells(y, 2)
TextBox3.Text = Sheets("Database").Cells(y, 3)
TextBox4.Text = Sheets("Database").Cells(y, 4)
TextBox5.Text = Sheets("Database").Cells(y, 5)
TextBox6.Text = Sheets("Database").Cells(y, 6)
TextBox7.Text = Sheets("Database").Cells(y, 7)
TextBox8.Text = Sheets("Database").Cells(y, 8)
TextBox9.Text = Sheets("Database").Cells(y, 12)

TextBox10.Text = Sheets("Database").Cells(y, 13)

ComboBox4.Text = Sheets("Database").Cells(y, 9)
ComboBox2.Text = Sheets("Database").Cells(y, 10)
ComboBox3.Text = Sheets("Database").Cells(y, 11)
ComboBox5.Text = Sheets("Database").Cells(y, 12)

'=================================================
'start code for Search Student Image and Display
'=================================================
On Error Resume Next
Image1.Picture = LoadPicture("C:\Photo\" & TextBox11.Text & Value & ".jpg")
Image1.PictureSizeMode = 1
End If
Next y

End Sub

Private Sub CommandButton9_Click()
'==================================
'Required field for Student Name
'==================================
On Error Resume Next
If TextBox11.Text = "" Then
Cancel = 1
MsgBox "Please enter registration number", vbOKOnly, "Learn with OMGStudy.Com"
TextBox11.SetFocus
Exit Sub
End If
'=============================================
'start code for Update Student Information
'=============================================
On Error Resume Next
Dim x As Long
Dim y As Long
x = Sheets("Database").Range("A" & Rows.Count).End(xlUp).Row
For y = 2 To x
If Sheets("Database").Cells(y, 1).Text = TextBox11.Value Then
Sheets("Database").Cells(y, 2) = TextBox2.Text
Sheets("Database").Cells(y, 3) = TextBox3.Text
Sheets("Database").Cells(y, 4) = TextBox4.Text
Sheets("Database").Cells(y, 5) = TextBox5.Text
Sheets("Database").Cells(y, 6) = TextBox6.Text
Sheets("Database").Cells(y, 7) = TextBox7.Text
Sheets("Database").Cells(y, 8) = TextBox8.Text
Sheets("Database").Cells(y, 12) = TextBox9.Text
Sheets("Database").Cells(y, 13) = TextBox10.Text

Sheets("Database").Cells(y, 9) = ComboBox4.Text
Sheets("Database").Cells(y, 10) = ComboBox2.Text
Sheets("Database").Cells(y, 11) = ComboBox3.Text
Sheets("Database").Cells(y, 12) = ComboBox5.Text


'=====================================
'start code for Update Student photo
'=====================================
On Error Resume Next
Dim i As String
i = TextBox1.Text
FileCopy fpath, "C:\Photo\" & i & ".jpg"
Me.TextBox1.Text = ""
End If
Next y
MsgBox "Student Information updated successfully.", vbOKOnly, "Learn with OMGStudy.Com"
Unload Me
Admission.Show

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label19_Click()

End Sub

Private Sub Label32_Click()

End Sub

Private Sub Label34_Click()

End Sub

Private Sub Label35_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub




'=====================================================================================================================
'Start code for Enter Numbers only in TextBox (Copy the source code from https://www.omgstudy.com) and use for TextBox.
'=====================================================================================================================
Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'========================================================
'start code for write  Numbers only in TextBox6
'========================================================
On Error Resume Next
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Invalid Key Pressed, You Can Enter Numbers Only", vbOKOnly, "Learn with OMGStudy.Com"
End If
End Sub
Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'========================================================
'start code for write  Numbers only in TextBox7
'========================================================
On Error Resume Next
If (KeyAscii > 47 And KeyAscii < 58) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "Invalid Key Pressed, You Can Enter Numbers Only", vbOKOnly, "Learn with OMGStudy.Com"
End If
End Sub








'========================================================
'End code for write  Numbers only in TextBox
'========================================================
'=====================================================================================================================
'start code for write Proper Text in TextBox (Copy the source code from https://www.omgstudy.com) and use for TextBox.
'=====================================================================================================================
'=============================
'Proper Sentence for TextBox2
'=============================
Private Sub TextBox2_Change()
Me.TextBox2 = Application.WorksheetFunction.Proper(TextBox2)
End Sub

'Proper Sentence for TextBox3
'=============================
Private Sub TextBox3_Change()
Me.TextBox3 = Application.WorksheetFunction.Proper(TextBox3)
End Sub
'=============================
'Proper Sentence for TextBox4
'=============================
Private Sub TextBox4_Change()
Me.TextBox4 = Application.WorksheetFunction.Proper(TextBox4)
End Sub

'=============================
'Proper Sentence for TextBox8
'=============================
Private Sub TextBox8_Change()
Me.TextBox8 = Application.WorksheetFunction.Proper(TextBox8)
End Sub

'=============================
'Proper Sentence for TextBox10
'=============================
Private Sub TextBox10_Change()
Me.TextBox10 = Application.WorksheetFunction.Proper(TextBox10)
End Sub



'========================================================
'End code for write  Proper Text in TextBox
'========================================================

Private Sub UserForm_Initialize()
On Error Resume Next
'==============================================================================
'start code for show autometic student Registration No. in Reg. No. TextBox
'==============================================================================
lrf = Sheets("Database").Range("B" & Rows.Count).End(xlUp).Row
TextBox1.Value = Sheets("Database").Cells(lrf + 1, 1).Value
TextBox2.SetFocus
'==============================================================================
'Start code for show autometic current Date in Admission Date
'==============================================================================
TextBox9.Text = Date
End Sub

