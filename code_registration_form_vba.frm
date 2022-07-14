VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "MakeName"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15495
   OleObjectBlob   =   "code_registration_form_vba.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cb1_value As String
Private cb2_value As String
Private cb3_value As String
Private Sub btn_run_Click()
    
    '//SUPPORT시트 컨트롤 변수
    Dim ws As Worksheet
    Dim lastRowIndex As Integer
    Dim insertRowIndex As Integer
    Dim stdColumnIndex As Integer
        
    Dim ingArry As Variant
    Dim i As Integer
    '//품목조회 및 등록을 위한 변수
    Dim sQuery As New Ctr_DB
    Dim itemCode As String
    Dim arry As Variant
    Dim iQuery As New Insert_DB
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ThisWorkbook.Activate
    
    On Error GoTo errMessage
    
    If pngDBloginId <> "" Then
        
        Set ws = ThisWorkbook.Worksheets("SUPPORT")
        
        
        With TextBox_spec
            .Value = Me.specType & " @ " & Me.spec_2 & " @ " & Me.spec_3 & " @ " & Me.spec_4 & " @ " & Me.specColorNum & " " & Me.spec_5
        End With
        
        TextBox_ingredient.Value = ""
        
        Dim mkIng As New FormCtr_IngredientMerge
        mkIng.arryX = 1 '//2열
        mkIng.arryY = 4 '//5행
            
        ingArry = mkIng.TwoDimentionArry
        '//성분값할당완료(0, #) 타이틀 / (1,#) 값
        
        '//성분값 문자열 조합
        If ingArry(0, 0) <> "" Then '//성분1 속성값이 있을때
            TextBox_ingredient = ingArry(1, 0) & "% " & ingArry(0, 0)
                Debug.Print "시작 : " & TextBox_ingredient
    
            For i = 1 To 4
                If ingArry(0, i) <> "" Then
                    TextBox_ingredient = TextBox_ingredient & " / " & ingArry(1, i) & "% " & ingArry(0, i)
                        Debug.Print TextBox_ingredient
                End If
            Next i
        End If
    
        
    
        If Me.radio_1 = True Then
            Debug.Print "재고성 품목등록 체크되었음."
        
            sQuery.selectString = "*"
            sQuery.tableString = "ecount_item_code"
            sQuery.whereString = "NAME = '" & Me.TextBox_itemName & "' AND  SPEC = '" & Me.TextBox_spec & "'"
        
            arry = sQuery.selectQurey
            Debug.Print "조회값이 있는지확인하였음"
            
            '//조회값이 비어있으면 코드를 생성함.
            If IsEmpty(arry) Then
                '//테이블이름
                    Debug.Print "값이없으므로 생성함"
                With iQuery
                    .tableString = "ecount_item_code"
                    .nameString = Me.TextBox_itemName
                    .specString = Me.TextBox_spec
                    .ingredientString = Me.TextBox_ingredient
                    .unitString = Me.spec_6
                    .styleNameString = Me.rawPageBomName
                    If Me.radio_1 = True Then
                        .purchase_type = "KS"
                    Else
                        .purchase_type = "KM"
                    End If
                End With
        
                '//등록실행
                iQuery.InsertQurey
        
                '//등록 후 생성된코드조회기능
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_code"
                    .whereString = "SPEC = '" & Me.TextBox_spec & "'"
                End With
                
                arry = sQuery.selectQurey
        
                itemCode = "KS" & Format(arry(0, 0), "0000")
                        
                Me.ExportMsg.Value = "코드가 등록되었습니다. (IDX CODE : " & itemCode & ")"
        
                stdColumnIndex = 2 '//NAME컬럼 기준
                lastRowIndex = ws.Cells(Rows.Count, stdColumnIndex).End(xlUp).Row
                insertRowIndex = lastRowIndex + 1
        
                With ws.Cells(insertRowIndex, stdColumnIndex)
                    .Offset(0, -1) = itemCode
                    .Offset(0, 0) = Me.TextBox_itemName
                    .Offset(0, 1) = Me.TextBox_spec
                    .Offset(0, 2) = Me.TextBox_ingredient
                    .Offset(0, 3) = Me.spec_6
                    Debug.Print InStr(Me.TextBox_itemName, "FABRIC")
                    If Me.MultiPage1.SelectedItem.Name = "P_raw" Then
                        .Offset(0, 4) = 0
                    Else
                        .Offset(0, 4) = 4
                    End If
                End With
        
            Else
                    Debug.Print "값이 있으므로 생성하지않음"
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_code"
                    .whereString = "SPEC = '" & Me.TextBox_spec & "'"
                End With
                    
                    arry = sQuery.selectQurey
                    
                    itemCode = "KS" & Format(arry(0, 0), "0000")
                    
                        Debug.Print "itemCode : " & itemCode
                        
                    Me.ExportMsg.Value = "규격이 중복되어 등록되지 않았습니다. (IDX CODE : " & itemCode & ")"
            End If
        
        ElseIf Me.radio_2 = True Then
            Debug.Print "시장성 품목등록 체크되었음."
        
        '//시장성품목_등록가능여부 코드조회기능
            With sQuery
                .selectString = "*"
                .tableString = "ecount_item_market"
                .whereString = "NAME = '" & Me.TextBox_itemName & "' AND  SPEC = '" & Me.TextBox_spec & "'"
            End With
            
                arry = sQuery.selectQurey
        
            '//조회값이 비어있으면 코드를 생성함.
            If IsEmpty(arry) Then
                '//테이블이름
                With iQuery
                    .tableString = "ecount_item_market"
                    .nameString = Me.TextBox_itemName
                    .specString = Me.TextBox_spec
                    .ingredientString = Me.TextBox_ingredient
                    .styleNameString = Me.rawPageBomName
                    If Me.radio_1 = True Then
                        .purchase_type = "KS"
                    Else
                        .purchase_type = "KM"
                    End If
                End With
        
                '//등록실행
                iQuery.InsertQurey
        
                '//등록 후 생성된코드조회기능
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_market"
                    .whereString = "SPEC = '" & Me.TextBox_spec & "'"
                End With
        
                    arry = sQuery.selectQurey
        
                itemCode = "KM" & Format(arry(0, 0), "0000")
        
                Me.ExportMsg.Value = "코드가 등록되었습니다. (IDX CODE : " & itemCode & ")"
        
                stdColumnIndex = 2 '//NAME컬럼 기준
                lastRowIndex = ws.Cells(Rows.Count, stdColumnIndex).End(xlUp).Row
                insertRowIndex = lastRowIndex + 1
                
                '//SUPPORT시트에 값을 입력함
                With ws.Cells(insertRowIndex, stdColumnIndex)
                    .Offset(0, -1) = itemCode
                    .Offset(0, 0) = Me.TextBox_itemName
                    .Offset(0, 1) = Me.TextBox_spec
                    .Offset(0, 2) = Me.TextBox_ingredient
                    .Offset(0, 3) = Me.spec_6
                    If Me.MultiPage1.SelectedItem.Name = "P_raw" Then
                        .Offset(0, 4) = 0
                    Else
                        .Offset(0, 4) = 4
                    End If
                    .Offset(0, 5) = Me.rawPageBomName
                End With
        
            Else
                    With sQuery
                        .selectString = "*"
                        .tableString = "ecount_item_market"
                        .whereString = "SPEC = '" & Me.TextBox_spec & "'"
                    End With
                    
                    arry = sQuery.selectQurey
                    
                    itemCode = "KM" & Format(arry(0, 0), "0000")
                    
                        Debug.Print "itemCode : " & itemCode
                        
                    Me.ExportMsg.Value = "규격이 중복되어 등록되지 않았습니다. (IDX CODE : " & itemCode & ")"
            End If
        
        Else
            MsgBox "오류발생으로 프로그램 종료됩니다."
            Exit Sub
        
        End If
    Else
        MsgBox "로그인 정보가 없습니다. 프로그램을 재실행해주세요. 입력창이 종료됩니다."
        Unload UserForm1
        
    End If
    ws.Activate
    Exit Sub
      
errMessage:
    ExportMsg.Value = "에러메세지 : " & Err.Description
    Exit Sub

    
End Sub

Private Sub CommandButton1_Click()
    
    Dim uQry As New Update_DB
    
    ThisWorkbook.Activate
    
    If Me.mod_radio_1 = True And InStr(Me.mod_searchCode, "KS") Then
        uQry.tableString = "ecount_item_code"
        uQry.UpdateQurey
    ElseIf Me.mod_radio_2 = True And InStr(Me.mod_searchCode, "KM") Then
        uQry.tableString = "ecount_item_market"
        uQry.UpdateQurey
        
        
    End If
    
    Call select_Union_db
    
End Sub

Private Sub CommandButton2_Click()

    Unload UserForm1
    
End Sub

Private Sub delPageBtnSearch_Click()

    Dim sq As New Ctr_DB
    Dim arry As Variant
    
    sq.selectString = "*"
    sq.tableString = "ecount_item_code"
    
    arry = sq.selectQurey
    Me.ListBox1.ColumnCount = 5
    
    Me.ListBox1.Column = arry
    
    
End Sub
'//수정 페이지
Private Sub listView_Click()
'//검색 기능
    Dim sq As New Ctr_DB
    Dim arry As Variant
    
    sq.selectString = "*"
    sq.tableString = "ecount_item_code"
    
    arry = sq.selectQurey
    Me.ListBox1.ColumnCount = 5
    
    Me.ListBox1.Column = arry

End Sub
'//수정 페이지
Private Sub listView_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub delete_button_Click()
    '//delete기능
    Dim mysql As New Insert_DB
    Dim yn As Integer
    
    
    If Me.mod_radio_1 = True And InStr(Me.mod_searchCode, "KS") Then
        mysql.tableString = "ecount_item_code"
        
        If authorInfo = 9999 Then
            yn = MsgBox("코드를삭제합니다.", vbYesNo)
            If yn = 6 Then
                
                    mysql.code_idx = Mid(Me.mod_searchCode, 3, Len(Me.mod_searchCode))
                Debug.Print (mysql.code_idx)
                mysql.delete_func
                MsgBox "삭제되었습니다."
            Else
                Debug.Print ("취소")
            End If
        Else
            MsgBox "삭제권한이없습니다."
        End If
    
    ElseIf Me.mod_radio_2 = True And InStr(Me.mod_searchCode, "KM") Then
        mysql.tableString = "ecount_item_market"
        
        If authorInfo = 9999 Then
            yn = MsgBox("코드를삭제합니다.", vbYesNo)
            If yn = 6 Then
                
                    mysql.code_idx = Mid(Me.mod_searchCode, 3, Len(Me.mod_searchCode))
                Debug.Print (mysql.code_idx)
                mysql.delete_func
                MsgBox "삭제되었습니다."
            Else
                Debug.Print ("취소")
            End If
        Else
            MsgBox "삭제권한이없습니다."
        End If
        
    Else
        MsgBox "코드입력 오류, 프로그램을 재실행해주세요."
        Exit Sub
    End If
    
End Sub

Private Sub Frame22_Click()

End Sub

Private Sub Frame21_Click()

End Sub

'//수정 페이지
Private Sub modPageBtnSearch_Click()
    
    Dim sqr As New Ctr_DB
    Dim arry As Variant
    
    If Me.mod_radio_1 = True And InStr(Me.mod_searchCode, "KS") Then
        Debug.Print "월마감 코드"
        With sqr
            .selectString = "*"
            .tableString = "ecount_item_code"
            .whereString = "IDX = '" & Mid(Me.mod_searchCode, 3, Len(Me.mod_searchCode)) & "'"
        End With
        
        If IsEmpty(sqr.selectQurey) Then
            MsgBox "조회값이 없습니다."
        Else
            arry = sqr.selectQurey
            
            With Me
                .mod_itemCode = Me.mod_searchCode
                .mod_itemName = arry(2, 0)
                .mod_itemSpec = arry(3, 0)
                .mod_Ingredient = arry(4, 0)
                .mod_itemUnit = arry(5, 0)
                .modifyPageBomName = arry(1, 0)
            End With
            
        End If
        
    ElseIf Me.mod_radio_2 = True And InStr(Me.mod_searchCode, "KM") Then
        Debug.Print "시장 코드"
        With sqr
            .selectString = "*"
            .tableString = "ecount_item_market"
            .whereString = "IDX = '" & Mid(Me.mod_searchCode, 3, Len(Me.mod_searchCode)) & "'"
        End With
        
        If IsEmpty(sqr.selectQurey) Then
            MsgBox "조회값이 없습니다."
        Else
            arry = sqr.selectQurey
            
            With Me
                .mod_itemCode = Me.mod_searchCode
                .mod_itemName = arry(2, 0)
                .mod_itemSpec = arry(3, 0)
                .mod_Ingredient = arry(4, 0)
                .mod_itemUnit = arry(5, 0)
                .modifyPageBomName = arry(1, 0)
            End With
        
        End If
            
        
    
    Else
        MsgBox "stock/market선택과 코드유형이 맞는지 확인해주세요."
    End If
    
    
    
    '등록일 arry(5,0)
        
    
End Sub
Private Sub subPageBtnMake_Click()

    Dim i As Integer
    '//SUPPORT시트 컨트롤 변수
    Dim ws As Worksheet
    Dim lastRowIndex As Integer
    Dim insertRowIndex As Integer
    Dim stdColumnIndex As Integer
    
    '//품목조회 및 등록을 위한 변수
    Dim sQuery As New Ctr_DB
    Dim itemCode As String
    Dim arry As Variant
    Dim iQuery As New Insert_DB
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ThisWorkbook.Activate
   
   On Error GoTo errMessage
    
    If pngDBloginId <> "" Then
    
        Set ws = Worksheets("SUPPORT")
        
        '//spec정보 입력
        Me.subPageInfoSpec.Value = Me.subPageSpec_1.Value & " @ " & Me.subPageSpecMaterials & " @ " & Me.subPageSpecSize & " @ " & Me.subPageSpecPartner & " @ " & Me.subPageSpecColor & " @ " & Me.subPageSpecRemark
        
        If Me.subPageRadio_1 = True Then
            Debug.Print "재고성 품목등록 체크되었음."
            sQuery.selectString = "*"
            sQuery.tableString = "ecount_item_code"
            sQuery.whereString = "NAME = '" & Me.subPageInfoName & "' AND  SPEC = '" & Me.subPageInfoSpec & "'"
            
            arry = sQuery.selectQurey
            
            '//조회값이 비어있으면 코드를 생성함.
            If IsEmpty(arry) Then
                '//테이블이름
                With iQuery
                    .tableString = "ecount_item_code"
                    .nameString = Me.subPageInfoName
                    .specString = Me.subPageInfoSpec
                    .unitString = Me.subPageSpecUnit
                    .styleNameString = Me.subPageBomName
                    If Me.subPageRadio_1 = True Then
                        .purchase_type = "KS"
                    Else
                        .purchase_type = "KM"
                    End If
                End With
                
                '//등록실행
                iQuery.InsertQurey
                    
                '//등록 후 생성된코드조회기능
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_code"
                    .whereString = "SPEC = '" & Me.subPageInfoSpec & "'"
                End With
                
                arry = sQuery.selectQurey
                
                itemCode = "KS" & Format(arry(0, 0), "0000")
                    Debug.Print "itemCode : " & itemCode
                Me.subPageMessage.Value = "코드가 등록되었습니다. (IDX CODE : " & itemCode & ")"
                
                stdColumnIndex = 2 '//NAME컬럼 기준
                lastRowIndex = ws.Cells(Rows.Count, stdColumnIndex).End(xlUp).Row
                insertRowIndex = lastRowIndex + 1
                
                '//SUPPORT시트에 등록된 데이터 입력하여 보이기
                With ws.Cells(insertRowIndex, stdColumnIndex)
                    .Offset(0, -1) = itemCode
                    .Offset(0, 0) = Me.subPageInfoName
                    .Offset(0, 1) = Me.subPageInfoSpec
                    .Offset(0, 3) = Me.subPageSpecUnit
                    If Me.MultiPage1.SelectedItem.Name = "P_raw" Then
                        .Offset(0, 4) = 0
                    Else
                        .Offset(0, 4) = 4
                    End If
                End With
                
            Else
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_code"
                    .whereString = "NAME = '" & Me.subPageInfoName & "' AND  SPEC = '" & Me.subPageInfoSpec & "'"
                End With
                
                arry = sQuery.selectQurey
                
                itemCode = "KS" & Format(arry(0, 0), "0000")
                
                Me.subPageMessage.Value = "규격이 중복되어 등록되지 않았습니다. (IDX CODE : " & itemCode & ")"
                
            End If
        
        ElseIf subPageRadio_2 = True Then
                Debug.Print "재고성 품목등록 체크되었음."
            With sQuery
                .selectString = "*"
                .tableString = "ecount_item_market"
                .whereString = "NAME = '" & Me.subPageInfoName & "' AND  SPEC = '" & Me.subPageInfoSpec & "'"
            End With
            
            arry = sQuery.selectQurey
            
            '//조회값이 비어있으면 코드를 생성함.
            If IsEmpty(arry) Then
                '//테이블이름
               With iQuery
                    .tableString = "ecount_item_market"
                    .nameString = Me.subPageInfoName
                    .specString = Me.subPageInfoSpec
                    .unitString = Me.subPageSpecUnit
                    .styleNameString = Me.subPageBomName
                    If Me.subPageRadio_1 = True Then
                        .purchase_type = "KS"
                    Else
                        .purchase_type = "KM"
                    End If
                End With
                
                '//등록실행
                iQuery.InsertQurey
                    
                '//등록 후 생성된코드조회기능
                '//등록 후 생성된코드조회기능
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_market"
                    .whereString = "SPEC = '" & Me.subPageInfoSpec & "'"
                End With
                
                arry = sQuery.selectQurey
                
                itemCode = "KM" & Format(arry(0, 0), "0000")
                    Debug.Print "itemCode : " & itemCode
                Me.subPageMessage.Value = "코드가 등록되었습니다. (IDX CODE : " & itemCode & ")"
                
                stdColumnIndex = 2 '//NAME컬럼 기준
                lastRowIndex = ws.Cells(Rows.Count, stdColumnIndex).End(xlUp).Row
                insertRowIndex = lastRowIndex + 1
                
                '//SUPPORT시트에 값을 입력함
                With ws.Cells(insertRowIndex, stdColumnIndex)
                    .Offset(0, -1) = itemCode
                    .Offset(0, 0) = Me.subPageInfoName
                    .Offset(0, 1) = Me.subPageInfoSpec
                    .Offset(0, 3) = Me.subPageSpecUnit
                    If Me.MultiPage1.SelectedItem.Name = "P_raw" Then
                        .Offset(0, 4) = 0
                    Else
                        .Offset(0, 4) = 4
                    End If
                    .Offset(0, 5) = Me.subPageBomName
                End With
            Else
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_market"
                    .whereString = "NAME = '" & Me.subPageInfoName & "' AND  SPEC = '" & Me.subPageInfoSpec & "'"
                End With
                '//조회실행
                arry = sQuery.selectQurey
                
                itemCode = "KM" & Format(arry(0, 0), "0000")
                    
                Me.subPageMessage.Value = "규격이 중복되어 등록되지 않았습니다. (IDX CODE : " & itemCode & ")"
                
            End If
                
        Else
                MsgBox "오류발생으로 프로그램 종료됩니다."
                Exit Sub
            
        End If
    Else
        MsgBox "로그인 정보가 없습니다. 프로그램을 재실행해주세요. 입력창이 종료됩니다."
        Unload UserForm1
        'Exit Sub
    End If
    ws.Activate
    Exit Sub
    
errMessage:
    MsgBox Err.Description
    ExportMsg.Value = Err.Description
    Exit Sub

End Sub
Private Sub D_midList_Change()

    Dim sl As New Ctr_DB
    Dim arry As Variant
    Dim changeCb2_value As String
    Dim i As Long
    
    If cb2_value <> changeCb2_value Then
        Me.D_lowList.Clear
    End If
    
    '//중분류 하위 쿼리
    If Me.D_midList.Value <> "" Then
        With sl
            .selectString = "distinct SUB_NAME"
            .tableString = "info_item_category"
            .whereString = "MAIN_NAME = '" & Me.D_midList.Value & "'"
        
            arry = sl.selectQurey
        End With
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
            Me.D_lowList.AddItem arry(0, i)
        Next i
    End If
    
    cb2_value = Me.D_midList.Value
    Debug.Print cb2_value
End Sub
Private Sub D_lowList_Change()
    
    With Me.TextBox_itemName
        .Value = ""
        .Value = Me.D_lowList & " / " & Me.D_midList
    End With
    
End Sub
Private Sub B_cancle_Click()

Unload UserForm1

End Sub

Private Sub UserForm_Initialize()
    '//유저폼 나타날떄 동작
    '//원재료 시트 선택
    Me.MultiPage1.Value = 0
    
    Dim sl As New Ctr_DB
    Dim arry As Variant
    Dim i As Integer
        
        Me.D_midList.Value = "FABRIC"
        Me.spec_6 = "YDS"
        Me.specType.Clear
        With sl
            .selectString = "distinct SUB_NAME"
            .tableString = "info_ingredient"
            .whereString = "MAIN_NAME = 'FABRIC' ORDER BY SUB_NAME ASC"
            
            arry = .selectQurey
            
        End With
                
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
            Me.specType.AddItem arry(0, i)
        Next i
        Erase arry
        
        '//중분류 호출
        sl.selectString = "distinct SUB_NAME"
        sl.tableString = "info_item_category"
        sl.whereString = "MAIN_NAME = 'RAW'"
        
        arry = sl.selectQurey
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
            Me.D_midList.AddItem arry(0, i)
        Next i
        Erase arry
        
        '//단위정보 호출
        sl.selectString = "distinct NAME_STRING"
        sl.tableString = "ecount_code"
        sl.whereString = "TYPE_NAME = 'UNIT'"
        
        arry = sl.selectQurey
        Debug.Print "첫번째 : " & arry(0, 0)
        Debug.Print "반복횟수 : " & UBound(arry, 2) - LBound(arry, 2)
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
            If IsNull(arry(0, 0)) Then
                Debug.Print "값이없음"
            Else
                Me.spec_6.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
        
        '//색상코드 호출
        sl.selectString = "distinct NAME"
        sl.tableString = "color_std_code"
        'sl.whereString = "SUB_NAME = 'Blue'"
        sl.whereString = ""
        
        arry = sl.selectQurey
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
        
            If IsEmpty(arry) Then
                Debug.Print "값이없음"
            Else
                Me.spec_5.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
        
        '//Ingredient 정보 호출
        sl.selectString = "distinct SUB_NAME"
        sl.tableString = "info_ingredient"
        sl.whereString = "MAIN_NAME = 'FABRIC' ORDER BY SUB_NAME ASC"
        
        arry = sl.selectQurey
        Debug.Print "두번째 : " & arry(0, 0)
        Debug.Print "반복횟수 : " & UBound(arry, 2) - LBound(arry, 2)
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
        
            If IsNull(arry(0, 0)) Then
                Debug.Print "값이없음"
            Else
                Me.ingredient_1.AddItem arry(0, i)
                Me.ingredient_2.AddItem arry(0, i)
                Me.ingredient_3.AddItem arry(0, i)
                Me.ingredient_4.AddItem arry(0, i)
                Me.ingredient_5.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
    
End Sub
Private Sub MultiPage1_Change()
    '//UserForm Page change
    
    Dim sl As New Ctr_DB
    Dim arry As Variant
    Dim i As Integer
    
    Me.subPageMidList.Clear
    Me.subPageSpecMaterials.Clear
    Me.subPageSpecUnit.Clear
    
    If Me.MultiPage1.SelectedItem.Name = "P_sub" Then
        '//subPageMidList업데이트
        sl.selectString = "distinct SUB_NAME"
        sl.tableString = "info_item_category"
        sl.whereString = "MAIN_NAME = 'SUB' ORDER BY SUB_NAME ASC"
        
        arry = sl.selectQurey
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        
            Me.subPageMidList.AddItem arry(0, i)
            Debug.Print "표시값 : " & arry(0, i)
        Next i
    
    End If
    
    '//단위정보 호출
        sl.selectString = "distinct NAME_STRING"
        sl.tableString = "ecount_code"
        sl.whereString = "TYPE_NAME = 'UNIT'"
        
        arry = sl.selectQurey
        Debug.Print "첫번째 : " & arry(0, 0)
        Debug.Print "반복횟수 : " & UBound(arry, 2) - LBound(arry, 2)
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
            If IsNull(arry(0, 0)) Then
                Debug.Print "값이없음"
            Else
                Me.subPageSpecUnit.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
        
    '//색상정보 호출
        sl.selectString = "distinct NAME"
        sl.tableString = "color_std_code"
        'sl.whereString = "SUB_NAME = 'Blue'"
        sl.whereString = ""
        
        arry = sl.selectQurey
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
        
            If IsEmpty(arry) Then
                Debug.Print "값이없음"
            Else
                Me.subPageSpecColor.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
        
    '//재질정보 호출
        sl.selectString = "distinct NAME_STRING"
        sl.tableString = "ecount_code"
        sl.whereString = "TYPE_NAME = 'MATERIAL TYPE'"
        
        arry = sl.selectQurey
        Debug.Print "첫번째 : " & arry(0, 0)
        Debug.Print "반복횟수 : " & UBound(arry, 2) - LBound(arry, 2)
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
            If IsNull(arry(0, 0)) Then
                Debug.Print "값이없음"
            Else
                Me.subPageSpecMaterials.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
End Sub
Private Sub subPageMidList_Change()
    Dim sl As New Ctr_DB
    Dim arry As Variant
    Dim i As Long
    
    Me.subPageLowList.Clear
    
    '//중분류 하위 쿼리
    If Me.subPageMidList.Value <> "" Then
        With sl
            .selectString = "distinct SUB_NAME"
            .tableString = "info_item_category"
            .whereString = "MAIN_NAME = '" & Me.subPageMidList.Value & "' ORDER BY SUB_NAME ASC"
        
            arry = sl.selectQurey
        End With
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
            Me.subPageLowList.AddItem arry(0, i)
        Next i
    End If
    
End Sub
Private Sub subPageLowList_Change()
    Dim sl As New Ctr_DB
    Dim arry As Variant
    Dim i As Long
    
    Me.subPageSpec_1.Clear
    
    '//중분류 하위 쿼리
    If Me.subPageLowList.Value <> "" Then
        With sl
            .selectString = "distinct SUB_NAME"
            .tableString = "info_item_category"
            .whereString = "MAIN_NAME = '" & Me.subPageLowList.Value & "' AND CATEGORY =  '" & Me.subPageMidList.Value & "' ORDER BY SUB_NAME ASC"
        
            arry = sl.selectQurey
        End With
        
        If IsEmpty(arry) Then
            Debug.Print "값이없음"
            
        Else
            For i = 0 To UBound(arry, 2) - LBound(arry, 2)
                Me.subPageSpec_1.AddItem arry(0, i)
            Next i
        End If
    End If
        
    With Me.subPageInfoName
        .Value = ""
        .Value = Me.subPageLowList & " / " & Me.subPageMidList
    End With
    
End Sub
Private Sub subPageBtnCancle_Click()

    Unload UserForm1
    
End Sub


