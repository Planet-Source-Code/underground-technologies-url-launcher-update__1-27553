VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   375
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   900
      Left            =   360
      Top             =   0
   End
   Begin VB.ComboBox cboAuto 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   330
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "type the URL here or select one from the list. press ""ENTER"" to visit the website."
      Top             =   30
      Width           =   7095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
Option Explicit

'Flag for the ComboBox.
Dim Backspaced As Boolean
Sub StayOnTop(Theform As Form)
Dim SetWinOnTop1
SetWinOnTop1 = SetWindowPos(Theform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Public Sub UnloadAllFrms2()
Dim Form As Form
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
End Sub

Sub Pause(interval)
    Dim Current
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub


Function if_file_excists(ByVal strng As String) As Integer
Dim num As Integer
On Error Resume Next
    num = Len(Dir$(strng))
    If Err Or num = 0 Then
        if_file_excists = False
        Else
        if_file_excists = True
    End If
End Function
Sub Load_ComboBox(path As String, Combo As ComboBox)
'example:
' Call Load_ComboBox("c:\windows\desktop\combo.cmb", Combo1)
    Dim What As String
    On Error Resume Next
    Open path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        'Combo.AddItem What$
        Call no_combo_dupes(Combo, What$)
    Wend
    Close #1
End Sub
Sub Save_ComboBox(path As String, Combo As ComboBox)
'Ex: Call Save_ComboBox("c:\windows\desktop\combo.cmb", combo1)
    Dim Savez As Long
    On Error Resume Next
    Open path$ For Output As #1
    For Savez& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(Savez&)
    Next Savez&
    Close #1
End Sub

Sub no_combo_dupes(Comb As ComboBox, txt As String)
    On Error GoTo Err_Proc
Dim poo, poo2, poo4, poo5, poo6
If txt = "" Then Exit Sub
For poo = 0 To Comb.ListCount - 1
   DoEvents
   poo2 = Comb.List(poo)
   poo4 = InStr(1, poo2, txt, 1)
   If poo4 Then
      poo5 = Len(poo2)
      poo6 = Len(txt)
      If poo5 = poo6 Then
       '  Txt = ""
         GoTo 890
      End If
   End If
Next poo
Comb.AddItem txt
890:
Err_Proc:
Exit Sub
End Sub
Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String, txt7 As ListBox) As String
    On Error GoTo Err_Proc
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, newstring As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            newstring$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = newstring$
        Else
            newstring$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    Pause 0.5
   ' client.List1.AddItem "" + newstring$
Err_Proc:
End Function

Private Sub cboAuto_Click()
Dim Result
Dim MessageBox%
If cboAuto.Text = "[clear All URL's]" Then
MessageBox% = MsgBox("are you sure you want to clear ALL URL's ?", vbYesNo + vbApplicationModal + vbQuestion + vbDefaultButton2, "Clear URL's?")
If MessageBox% = vbYes Then
cboAuto.Clear
Call Save_ComboBox(App.path & "\url's.txt", cboAuto)
combo_clear.Show
Exit Sub
Else
Exit Sub
End If
End If
' add the url to the combobox
Call no_combo_dupes(cboAuto, cboAuto)

' execute internet explorer and the selected url
Result = Shell("start.exe " & cboAuto.Text, vbHide)

' pause the program for 1/2 a second
Pause 0.5
' start the unload project timer
Timer1.Enabled = True: Exit Sub
End Sub

Private Sub cboAuto_KeyPress(KeyAscii As Integer)
Dim Result
If KeyAscii = 13 Then ' check if the user has hit "enter" key
' execute internet explorer and selected url
Result = Shell("start.exe " & cboAuto.Text, vbHide)
' add the selected url to the combobox
Call no_combo_dupes(cboAuto, cboAuto)
KeyAscii = 0 ' this will stop the anoying beep when pressing enter
Unload Me ' unload the form
End If
End Sub

Private Sub Form_Load()
Me.Top = (Screen.Height * 1.8) / 2 - Me.Height / 2
Me.Left = 100
' place the form on top of ALL other windows
StayOnTop Me
' add some test items into the combo
    With cboAuto
    .AddItem "[clear All URL's]"
    .AddItem "http://www.planet-source-code.com"
    .AddItem "http://www.vbthunder.com"
    .AddItem "http://www.mvps.org/vb/"
    .AddItem "http://www.mvps.org/vbnet/"
    .AddItem "http://members.aol.com/btmtz/vb/index.htm"
    .AddItem "http://www.zonecorp.com"
    .AddItem "http://www.microsoft.com"
    .AddItem "http://www.mvps.org/ccrp/"
    End With
    
' Check for saved url's.
If if_file_excists(App.path & "\url's.txt") = 0 Then
' If none then create a saved url file
Open (App.path & "\url's.txt") For Output As #1
End If
Close #1

' load the saved url's
    Call Load_ComboBox(App.path & "\url's.txt", cboAuto)
End Sub

Private Sub cboAuto_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If cboAuto.Text <> "" Then
            ' let the Change event know that it
            ' shouldn't respond to this change.
            Backspaced = True
        End If
    End If

End Sub

Private Sub cboAuto_Change()
    
    If Backspaced = True Or cboAuto.Text = "" Then
        Backspaced = False
        Exit Sub
    End If

    Dim i As Long
    Dim nSel As Long
    ' run through the available items and
    ' grab the first matching one.
    For i = 0 To cboAuto.ListCount - 1
        If InStr(1, cboAuto.List(i), cboAuto.Text, _
        vbTextCompare) = 1 Then
            ' save the SelStart property.
            nSel = cboAuto.SelStart
            cboAuto.Text = cboAuto.List(i)
            ' set the selection in the combo.
            cboAuto.SelStart = nSel
            cboAuto.SelLength = Len(cboAuto.Text) - nSel
            Exit For
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)

' save the url's to a file
Call Save_ComboBox(App.path & "\url's.txt", cboAuto)

' pause the program for one second
Pause 1

' end the program
End
End Sub



Public Sub Check_URL()
Dim strURL, rest As String
    If strURL = "" Then
        Exit Sub
    Else
    rest$ = strURL
      strURL = Left$(rest$, InStr(rest$, ";") - 1)
        rest$ = Right(rest$, Len(rest$) - InStr(rest$, ";"))
        ' Check_URL = rest$
          Exit Sub
    End If
End Sub

Private Sub Timer1_Timer()
'the only reason this timer is here is
'because i was having trouble unloading the project in
'the cboAuto_Click() event so i added the timer to do
'work for me
UnloadAllFrms2
End Sub


