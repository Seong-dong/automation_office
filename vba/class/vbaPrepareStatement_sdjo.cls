VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbControll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pstmt As String

Private svr_name As String
Private svr_port As String
Private svr_database_name As String
Private svr_user_name As String
Private svr_password As String
Private svr_conn_info As String
Private Function svrInformation_func() As String
    svr_name = "serverAddress"
    svr_port = "portNumber"
    svr_database_name = "databaseName"
    svr_user_name = "connectID"
    svr_password = "pw"
    svr_conn_info = "Driver={MariaDB ODBC 2.0 Driver};Server=" + svr_name + ";Port=" + svr_port + ";Database=" + _
                    svr_database_name + ";User=" + svr_user_name + ";Password=" + svr_password + ";Option=2;"
    Debug.Print (svr_conn_info)
    
    svrInformation_func = svr_conn_info
    
End Function
Public Function prepareStatment(query_string As String) As String
    'Debug.Print (query_string)
    pstmt = query_string
End Function
Public Function setString(ByVal arry As Collection) As String
    Dim countChar As Integer
    
'    If IsEmpty(arry) Then
'        Debug.Print ("where���� ����")
'    Else
        Dim i As Integer, j As Integer
        Dim singleChr As String
        Dim tmpString As String
        countChar = Len(pstmt) - Len(Replace(pstmt, "?", ""))
       ' Debug.Print ("��������:" & countChar)
        
        '���ڿ��� ���ư��鼭 ?�� �߰��ϸ� ��ü�Ұ�.
        j = 1
        For i = 0 To Len(pstmt)
            singleChr = Mid(pstmt, i + 1, 1)
            If singleChr = "?" Then
                singleChr = "'" & arry(j) & "'"
                j = j + 1
            End If
            
           ' Debug.Print (singleChr)
            
            tmpString = tmpString + singleChr
        Next i
           ' Debug.Print (setString)
            
        setString = tmpString
    
End Function
Function executeQury_func(query As String) As Variant
'//ǰ���ڵ� ��ȸ���
    '//Microsoft ActiveX Data Objects 6.1 Library �߰�.
    Dim conn As ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim ws As Worksheet
    
'    Set ws = ThisWorkbook.Worksheets(result_sheet_name)
    Set conn = New ADODB.Connection
    
'    ws.Activate
    '//������������
    Dim svr_info As String
    
    svr_info = svrInformation_func()
    Debug.Print (svr_info)
            
    conn.ConnectionString = svr_info
    conn.Open
    '//��������

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open Source:=query, ActiveConnection:=conn, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly, Options:=adCmdText

    If rs.EOF Then
        Debug.Print "��ȸ���ǿ� �ش��ϴ� �ڷᰡ �����ϴ�."
        MsgBox "��ȸ���ǿ� �ش��ϴ� �ڷᰡ �����ϴ�."

    Else
'
'        For i = 0 To rs.Fields.count - 1
'            Cells(20, i + 1).Value = rs.Fields(i).Name
'        Next
'
'        rs.MoveFirst
'
'        With ws
'           .Cells(21, 1).CopyFromRecordset rs
'        End With
'
'        'ws.Columns("A:Z").AutoFit
'
    End If
   executeQury_func = rs.GetRows
   rs.Close
   conn.Close

End Function
