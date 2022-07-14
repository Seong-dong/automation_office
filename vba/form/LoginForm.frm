VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login"
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim lg As New Ctr_DB
    Dim result As Variant
    
    Dim ws As Worksheet
    Dim strSht As Worksheet
    
    ThisWorkbook.Activate
    
    pngDBloginId = Me.identityBox
    Debug.Print "로그인된계정 : " & pngDBloginId
        
    lg.accountId = UCase(Me.identityBox) '대문자로 치환.
    lg.accountPw = Me.pwBox
    Debug.Print (lg.accountId)
    
    lg.selectString = "*"
    lg.tableString = "info_person"
    lg.whereString = "ID = '" & lg.accountId & "' AND PW = '" & lg.accountPw & "'"

    Debug.Print lg.whereString
    result = lg.loginCk
    
    If IsEmpty(result) Then
        MsgBox "계정정보가 없습니다."
    Else
        authorInfo = result(4, 0)
        If authorInfo = 9999 Then
            MsgBox "Ecount등록권한으로 로그인되었습니다."
        Else
            MsgBox "로그인되었습니다."
        End If
        
        Unload LoginForm
        
    End If
   
    
 
End Sub

Private Sub CommandButton2_Click()
    Unload LoginForm
    ThisWorkbook.Close saveChanges:=True
End Sub


Private Sub CommandButton3_Click()
'//비번바꾸기창
Show
End Sub

Private Sub text_box_Change()

End Sub

Private Sub UserForm_Initialize()

    Me.pwBox.PasswordChar = "*"
'    Me.text_box.Value = "ttt"
    
    
End Sub
