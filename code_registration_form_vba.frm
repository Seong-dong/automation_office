VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "MakeName"
   ClientHeight    =   11130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15495
   OleObjectBlob   =   "code_registration_form_vba.frx":0000
   StartUpPosition =   1  '������ ���
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
    
    '//SUPPORT��Ʈ ��Ʈ�� ����
    Dim ws As Worksheet
    Dim lastRowIndex As Integer
    Dim insertRowIndex As Integer
    Dim stdColumnIndex As Integer
        
    Dim ingArry As Variant
    Dim i As Integer
    '//ǰ����ȸ �� ����� ���� ����
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
        mkIng.arryX = 1 '//2��
        mkIng.arryY = 4 '//5��
            
        ingArry = mkIng.TwoDimentionArry
        '//���а��Ҵ�Ϸ�(0, #) Ÿ��Ʋ / (1,#) ��
        
        '//���а� ���ڿ� ����
        If ingArry(0, 0) <> "" Then '//����1 �Ӽ����� ������
            TextBox_ingredient = ingArry(1, 0) & "% " & ingArry(0, 0)
                Debug.Print "���� : " & TextBox_ingredient
    
            For i = 1 To 4
                If ingArry(0, i) <> "" Then
                    TextBox_ingredient = TextBox_ingredient & " / " & ingArry(1, i) & "% " & ingArry(0, i)
                        Debug.Print TextBox_ingredient
                End If
            Next i
        End If
    
        
    
        If Me.radio_1 = True Then
            Debug.Print "��� ǰ���� üũ�Ǿ���."
        
            sQuery.selectString = "*"
            sQuery.tableString = "ecount_item_code"
            sQuery.whereString = "NAME = '" & Me.TextBox_itemName & "' AND  SPEC = '" & Me.TextBox_spec & "'"
        
            arry = sQuery.selectQurey
            Debug.Print "��ȸ���� �ִ���Ȯ���Ͽ���"
            
            '//��ȸ���� ��������� �ڵ带 ������.
            If IsEmpty(arry) Then
                '//���̺��̸�
                    Debug.Print "���̾����Ƿ� ������"
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
        
                '//��Ͻ���
                iQuery.InsertQurey
        
                '//��� �� �������ڵ���ȸ���
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_code"
                    .whereString = "SPEC = '" & Me.TextBox_spec & "'"
                End With
                
                arry = sQuery.selectQurey
        
                itemCode = "KS" & Format(arry(0, 0), "0000")
                        
                Me.ExportMsg.Value = "�ڵ尡 ��ϵǾ����ϴ�. (IDX CODE : " & itemCode & ")"
        
                stdColumnIndex = 2 '//NAME�÷� ����
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
                    Debug.Print "���� �����Ƿ� ������������"
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_code"
                    .whereString = "SPEC = '" & Me.TextBox_spec & "'"
                End With
                    
                    arry = sQuery.selectQurey
                    
                    itemCode = "KS" & Format(arry(0, 0), "0000")
                    
                        Debug.Print "itemCode : " & itemCode
                        
                    Me.ExportMsg.Value = "�԰��� �ߺ��Ǿ� ��ϵ��� �ʾҽ��ϴ�. (IDX CODE : " & itemCode & ")"
            End If
        
        ElseIf Me.radio_2 = True Then
            Debug.Print "���强 ǰ���� üũ�Ǿ���."
        
        '//���强ǰ��_��ϰ��ɿ��� �ڵ���ȸ���
            With sQuery
                .selectString = "*"
                .tableString = "ecount_item_market"
                .whereString = "NAME = '" & Me.TextBox_itemName & "' AND  SPEC = '" & Me.TextBox_spec & "'"
            End With
            
                arry = sQuery.selectQurey
        
            '//��ȸ���� ��������� �ڵ带 ������.
            If IsEmpty(arry) Then
                '//���̺��̸�
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
        
                '//��Ͻ���
                iQuery.InsertQurey
        
                '//��� �� �������ڵ���ȸ���
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_market"
                    .whereString = "SPEC = '" & Me.TextBox_spec & "'"
                End With
        
                    arry = sQuery.selectQurey
        
                itemCode = "KM" & Format(arry(0, 0), "0000")
        
                Me.ExportMsg.Value = "�ڵ尡 ��ϵǾ����ϴ�. (IDX CODE : " & itemCode & ")"
        
                stdColumnIndex = 2 '//NAME�÷� ����
                lastRowIndex = ws.Cells(Rows.Count, stdColumnIndex).End(xlUp).Row
                insertRowIndex = lastRowIndex + 1
                
                '//SUPPORT��Ʈ�� ���� �Է���
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
                        
                    Me.ExportMsg.Value = "�԰��� �ߺ��Ǿ� ��ϵ��� �ʾҽ��ϴ�. (IDX CODE : " & itemCode & ")"
            End If
        
        Else
            MsgBox "�����߻����� ���α׷� ����˴ϴ�."
            Exit Sub
        
        End If
    Else
        MsgBox "�α��� ������ �����ϴ�. ���α׷��� ��������ּ���. �Է�â�� ����˴ϴ�."
        Unload UserForm1
        
    End If
    ws.Activate
    Exit Sub
      
errMessage:
    ExportMsg.Value = "�����޼��� : " & Err.Description
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
'//���� ������
Private Sub listView_Click()
'//�˻� ���
    Dim sq As New Ctr_DB
    Dim arry As Variant
    
    sq.selectString = "*"
    sq.tableString = "ecount_item_code"
    
    arry = sq.selectQurey
    Me.ListBox1.ColumnCount = 5
    
    Me.ListBox1.Column = arry

End Sub
'//���� ������
Private Sub listView_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub delete_button_Click()
    '//delete���
    Dim mysql As New Insert_DB
    Dim yn As Integer
    
    
    If Me.mod_radio_1 = True And InStr(Me.mod_searchCode, "KS") Then
        mysql.tableString = "ecount_item_code"
        
        If authorInfo = 9999 Then
            yn = MsgBox("�ڵ带�����մϴ�.", vbYesNo)
            If yn = 6 Then
                
                    mysql.code_idx = Mid(Me.mod_searchCode, 3, Len(Me.mod_searchCode))
                Debug.Print (mysql.code_idx)
                mysql.delete_func
                MsgBox "�����Ǿ����ϴ�."
            Else
                Debug.Print ("���")
            End If
        Else
            MsgBox "���������̾����ϴ�."
        End If
    
    ElseIf Me.mod_radio_2 = True And InStr(Me.mod_searchCode, "KM") Then
        mysql.tableString = "ecount_item_market"
        
        If authorInfo = 9999 Then
            yn = MsgBox("�ڵ带�����մϴ�.", vbYesNo)
            If yn = 6 Then
                
                    mysql.code_idx = Mid(Me.mod_searchCode, 3, Len(Me.mod_searchCode))
                Debug.Print (mysql.code_idx)
                mysql.delete_func
                MsgBox "�����Ǿ����ϴ�."
            Else
                Debug.Print ("���")
            End If
        Else
            MsgBox "���������̾����ϴ�."
        End If
        
    Else
        MsgBox "�ڵ��Է� ����, ���α׷��� ��������ּ���."
        Exit Sub
    End If
    
End Sub

Private Sub Frame22_Click()

End Sub

Private Sub Frame21_Click()

End Sub

'//���� ������
Private Sub modPageBtnSearch_Click()
    
    Dim sqr As New Ctr_DB
    Dim arry As Variant
    
    If Me.mod_radio_1 = True And InStr(Me.mod_searchCode, "KS") Then
        Debug.Print "������ �ڵ�"
        With sqr
            .selectString = "*"
            .tableString = "ecount_item_code"
            .whereString = "IDX = '" & Mid(Me.mod_searchCode, 3, Len(Me.mod_searchCode)) & "'"
        End With
        
        If IsEmpty(sqr.selectQurey) Then
            MsgBox "��ȸ���� �����ϴ�."
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
        Debug.Print "���� �ڵ�"
        With sqr
            .selectString = "*"
            .tableString = "ecount_item_market"
            .whereString = "IDX = '" & Mid(Me.mod_searchCode, 3, Len(Me.mod_searchCode)) & "'"
        End With
        
        If IsEmpty(sqr.selectQurey) Then
            MsgBox "��ȸ���� �����ϴ�."
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
        MsgBox "stock/market���ð� �ڵ������� �´��� Ȯ�����ּ���."
    End If
    
    
    
    '����� arry(5,0)
        
    
End Sub
Private Sub subPageBtnMake_Click()

    Dim i As Integer
    '//SUPPORT��Ʈ ��Ʈ�� ����
    Dim ws As Worksheet
    Dim lastRowIndex As Integer
    Dim insertRowIndex As Integer
    Dim stdColumnIndex As Integer
    
    '//ǰ����ȸ �� ����� ���� ����
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
        
        '//spec���� �Է�
        Me.subPageInfoSpec.Value = Me.subPageSpec_1.Value & " @ " & Me.subPageSpecMaterials & " @ " & Me.subPageSpecSize & " @ " & Me.subPageSpecPartner & " @ " & Me.subPageSpecColor & " @ " & Me.subPageSpecRemark
        
        If Me.subPageRadio_1 = True Then
            Debug.Print "��� ǰ���� üũ�Ǿ���."
            sQuery.selectString = "*"
            sQuery.tableString = "ecount_item_code"
            sQuery.whereString = "NAME = '" & Me.subPageInfoName & "' AND  SPEC = '" & Me.subPageInfoSpec & "'"
            
            arry = sQuery.selectQurey
            
            '//��ȸ���� ��������� �ڵ带 ������.
            If IsEmpty(arry) Then
                '//���̺��̸�
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
                
                '//��Ͻ���
                iQuery.InsertQurey
                    
                '//��� �� �������ڵ���ȸ���
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_code"
                    .whereString = "SPEC = '" & Me.subPageInfoSpec & "'"
                End With
                
                arry = sQuery.selectQurey
                
                itemCode = "KS" & Format(arry(0, 0), "0000")
                    Debug.Print "itemCode : " & itemCode
                Me.subPageMessage.Value = "�ڵ尡 ��ϵǾ����ϴ�. (IDX CODE : " & itemCode & ")"
                
                stdColumnIndex = 2 '//NAME�÷� ����
                lastRowIndex = ws.Cells(Rows.Count, stdColumnIndex).End(xlUp).Row
                insertRowIndex = lastRowIndex + 1
                
                '//SUPPORT��Ʈ�� ��ϵ� ������ �Է��Ͽ� ���̱�
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
                
                Me.subPageMessage.Value = "�԰��� �ߺ��Ǿ� ��ϵ��� �ʾҽ��ϴ�. (IDX CODE : " & itemCode & ")"
                
            End If
        
        ElseIf subPageRadio_2 = True Then
                Debug.Print "��� ǰ���� üũ�Ǿ���."
            With sQuery
                .selectString = "*"
                .tableString = "ecount_item_market"
                .whereString = "NAME = '" & Me.subPageInfoName & "' AND  SPEC = '" & Me.subPageInfoSpec & "'"
            End With
            
            arry = sQuery.selectQurey
            
            '//��ȸ���� ��������� �ڵ带 ������.
            If IsEmpty(arry) Then
                '//���̺��̸�
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
                
                '//��Ͻ���
                iQuery.InsertQurey
                    
                '//��� �� �������ڵ���ȸ���
                '//��� �� �������ڵ���ȸ���
                With sQuery
                    .selectString = "*"
                    .tableString = "ecount_item_market"
                    .whereString = "SPEC = '" & Me.subPageInfoSpec & "'"
                End With
                
                arry = sQuery.selectQurey
                
                itemCode = "KM" & Format(arry(0, 0), "0000")
                    Debug.Print "itemCode : " & itemCode
                Me.subPageMessage.Value = "�ڵ尡 ��ϵǾ����ϴ�. (IDX CODE : " & itemCode & ")"
                
                stdColumnIndex = 2 '//NAME�÷� ����
                lastRowIndex = ws.Cells(Rows.Count, stdColumnIndex).End(xlUp).Row
                insertRowIndex = lastRowIndex + 1
                
                '//SUPPORT��Ʈ�� ���� �Է���
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
                '//��ȸ����
                arry = sQuery.selectQurey
                
                itemCode = "KM" & Format(arry(0, 0), "0000")
                    
                Me.subPageMessage.Value = "�԰��� �ߺ��Ǿ� ��ϵ��� �ʾҽ��ϴ�. (IDX CODE : " & itemCode & ")"
                
            End If
                
        Else
                MsgBox "�����߻����� ���α׷� ����˴ϴ�."
                Exit Sub
            
        End If
    Else
        MsgBox "�α��� ������ �����ϴ�. ���α׷��� ��������ּ���. �Է�â�� ����˴ϴ�."
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
    
    '//�ߺз� ���� ����
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
    '//������ ��Ÿ���� ����
    '//����� ��Ʈ ����
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
        
        '//�ߺз� ȣ��
        sl.selectString = "distinct SUB_NAME"
        sl.tableString = "info_item_category"
        sl.whereString = "MAIN_NAME = 'RAW'"
        
        arry = sl.selectQurey
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
            Me.D_midList.AddItem arry(0, i)
        Next i
        Erase arry
        
        '//�������� ȣ��
        sl.selectString = "distinct NAME_STRING"
        sl.tableString = "ecount_code"
        sl.whereString = "TYPE_NAME = 'UNIT'"
        
        arry = sl.selectQurey
        Debug.Print "ù��° : " & arry(0, 0)
        Debug.Print "�ݺ�Ƚ�� : " & UBound(arry, 2) - LBound(arry, 2)
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
            If IsNull(arry(0, 0)) Then
                Debug.Print "���̾���"
            Else
                Me.spec_6.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
        
        '//�����ڵ� ȣ��
        sl.selectString = "distinct NAME"
        sl.tableString = "color_std_code"
        'sl.whereString = "SUB_NAME = 'Blue'"
        sl.whereString = ""
        
        arry = sl.selectQurey
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
        
            If IsEmpty(arry) Then
                Debug.Print "���̾���"
            Else
                Me.spec_5.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
        
        '//Ingredient ���� ȣ��
        sl.selectString = "distinct SUB_NAME"
        sl.tableString = "info_ingredient"
        sl.whereString = "MAIN_NAME = 'FABRIC' ORDER BY SUB_NAME ASC"
        
        arry = sl.selectQurey
        Debug.Print "�ι�° : " & arry(0, 0)
        Debug.Print "�ݺ�Ƚ�� : " & UBound(arry, 2) - LBound(arry, 2)
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
        
            If IsNull(arry(0, 0)) Then
                Debug.Print "���̾���"
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
        '//subPageMidList������Ʈ
        sl.selectString = "distinct SUB_NAME"
        sl.tableString = "info_item_category"
        sl.whereString = "MAIN_NAME = 'SUB' ORDER BY SUB_NAME ASC"
        
        arry = sl.selectQurey
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        
            Me.subPageMidList.AddItem arry(0, i)
            Debug.Print "ǥ�ð� : " & arry(0, i)
        Next i
    
    End If
    
    '//�������� ȣ��
        sl.selectString = "distinct NAME_STRING"
        sl.tableString = "ecount_code"
        sl.whereString = "TYPE_NAME = 'UNIT'"
        
        arry = sl.selectQurey
        Debug.Print "ù��° : " & arry(0, 0)
        Debug.Print "�ݺ�Ƚ�� : " & UBound(arry, 2) - LBound(arry, 2)
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
            If IsNull(arry(0, 0)) Then
                Debug.Print "���̾���"
            Else
                Me.subPageSpecUnit.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
        
    '//�������� ȣ��
        sl.selectString = "distinct NAME"
        sl.tableString = "color_std_code"
        'sl.whereString = "SUB_NAME = 'Blue'"
        sl.whereString = ""
        
        arry = sl.selectQurey
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
        
            If IsEmpty(arry) Then
                Debug.Print "���̾���"
            Else
                Me.subPageSpecColor.AddItem arry(0, i)
            End If
            Debug.Print arry(0, i)
            
        Next i
        
    '//�������� ȣ��
        sl.selectString = "distinct NAME_STRING"
        sl.tableString = "ecount_code"
        sl.whereString = "TYPE_NAME = 'MATERIAL TYPE'"
        
        arry = sl.selectQurey
        Debug.Print "ù��° : " & arry(0, 0)
        Debug.Print "�ݺ�Ƚ�� : " & UBound(arry, 2) - LBound(arry, 2)
        
        For i = 0 To UBound(arry, 2) - LBound(arry, 2)
        'For i = 0 To 3
            If IsNull(arry(0, 0)) Then
                Debug.Print "���̾���"
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
    
    '//�ߺз� ���� ����
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
    
    '//�ߺз� ���� ����
    If Me.subPageLowList.Value <> "" Then
        With sl
            .selectString = "distinct SUB_NAME"
            .tableString = "info_item_category"
            .whereString = "MAIN_NAME = '" & Me.subPageLowList.Value & "' AND CATEGORY =  '" & Me.subPageMidList.Value & "' ORDER BY SUB_NAME ASC"
        
            arry = sl.selectQurey
        End With
        
        If IsEmpty(arry) Then
            Debug.Print "���̾���"
            
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


