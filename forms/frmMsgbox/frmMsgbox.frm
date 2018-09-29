VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsgbox 
   Caption         =   "Msgbox"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9705
   OleObjectBlob   =   "frmMsgbox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMsgbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public result As Integer
Private buttons_status As Integer
' -1 = undefined
' -2 = all
' vbcancel = Cancel
' vbok = OK
' vbno = No
' vbyes = Yes
' 100 = all options
' 101 = option1
' 102 = option2
' 103 = option3

Private Sub UserForm_Initialize()
    Me.result = -1
    Me.setText
    Me.setTitle
    Me.setReadonly
    Me.setButtons
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If buttons_status = -1 Then If CloseMode = 0 Then Cancel = True
End Sub


Public Function XXX(ByVal text As String, Optional ByVal title As String, Optional ByVal buttons As Integer) As Integer
    Dim f As New frmMsgbox
    f.setText text
    If Not IsMissing(title) Then f.setTitle title
    If Not IsMissing(buttons) Then f.setButtons buttons
    f.Show
    XXX = f.result
    Unload f
    Unload Me
End Function

' =========================== GETTERS/SETTERS
Public Sub setText(Optional ByVal str As String = "")
    txtText.text = str
    txtText.SetFocus ' to make next line work, first we must set focus on Textbox
    txtText.CurLine = 0
End Sub
Public Function getText() As String
    getText = txtText.text
End Function

Public Sub setTitle(Optional ByVal title As String = "msgboXXX")
    Me.Caption = title
End Sub
Public Function getTitle() As String
    getTitle = Me.Caption
End Function

Public Sub setButtons(Optional ByVal buttons As Integer = -1)
    buttons_status = buttons
    'resetting buttons
    Me.btnOk.Visible = False
    Me.btnCancel.Visible = False
    Me.btnYes.Visible = False
    Me.btnNo.Visible = False
    Me.btnOption1.Visible = False
    Me.btnOption2.Visible = False
    Me.btnOption3.Visible = False
    
    'setting new option
    Select Case buttons
        Case -1 ' none
        Case -2 ' all
            Me.btnOk.Visible = True
            Me.btnCancel.Visible = True
            Me.btnYes.Visible = True
            Me.btnNo.Visible = True
            Me.btnOption1.Visible = True
            Me.btnOption2.Visible = True
            Me.btnOption3.Visible = True
        Case vbOK
            Me.btnOk.Visible = True
        Case vbCancel
            Me.btnCancel.Visible = True
        Case vbYes
            Me.btnYes.Visible = True
        Case vbNo
            Me.btnNo.Visible = True
        Case vbOKCancel
            Me.btnOk.Visible = True
            Me.btnCancel.Visible = True
        Case vbYesNo
            Me.btnYes.Visible = True
            Me.btnNo.Visible = True
        Case vbYesNoCancel
            Me.btnYes.Visible = True
            Me.btnNo.Visible = True
            Me.btnCancel.Visible = True
            
    End Select
End Sub
Public Function getButtons() As Integer
    getButtons = Me.buttons_status
End Function

Public Sub setReadonly(Optional ByVal readonly As Boolean = True)
    txtText.Locked = readonly
End Sub
Public Function getReadonly() As Boolean
    getReadonly = txtText.Locked
End Function

' =========================== BUTTONS
Private Sub btnCancel_Click()
    Me.result = vbCancel
    Unload Me
End Sub

Private Sub btnOk_Click()
    Me.result = vbOK
    Unload Me
End Sub

Private Sub btnNo_Click()
    Me.result = vbNo
    Unload Me
End Sub

Private Sub btnYes_Click()
    Me.result = vbYes
    Unload Me
End Sub

Private Sub btnOption1_Click()
    Me.result = 101
    Unload Me
End Sub

Private Sub btnOption2_Click()
    Me.result = 102
    Unload Me
End Sub

Private Sub btnOption3_Click()
    Me.result = 103
    Unload Me
End Sub

