VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "databaseControll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//make by.icurfer
Private pstmt As String

Private svr_name As String
Private svr_port As String
Private svr_database_name As String
Private svr_user_name As String
Private svr_password As String
Private svr_conn_info As String
Private Function svrInformation_func() As String
    svr_name = "data.pngcap.com"
    svr_port = "3307"
    svr_database_name = "png"
    svr_user_name = "sdjo"
    svr_password = "Png1234!@!@"
    svr_conn_info = "Driver={MariaDB ODBC 2.0 Driver};Server=" + svr_name + ";Port=" + svr_port + ";Database=" + _
                    svr_database_name + ";User=" + svr_user_name + ";Password=" + svr_password + ";Option=2;"
    Debug.Print ("접속정보")
    
    svrInformation_func = svr_conn_info
    
End Function
Public Function prepareStatment(query_string As String) As String
    
    pstmt = query_string
    Debug.Print ("query: " + pstmt)
    
    prepareStatment = pstmt
End Function
Public Function setString(ByVal arry As Collection) As String
    Dim countChar As Integer
    
'    If IsEmpty(arry) Then

'    Else
        Dim i As Integer, j As Integer
        Dim singleChr As String
        Dim tmpString As String
        countChar = Len(pstmt) - Len(Replace(pstmt, "?", ""))
       
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

    Dim conn As ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim ws As Worksheet

    Set conn = New ADODB.Connection
    
    Dim svr_info As String
    
    svr_info = svrInformation_func()
            
    conn.ConnectionString = svr_info
    conn.Open
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    rs.Open Source:=query, ActiveConnection:=conn, CursorType:=adOpenForwardOnly, LockType:=adLockReadOnly, Options:=adCmdText

    If rs.EOF Then
        
        MsgBox "No search DATA."

    Else
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
    End If
   executeQury_func = rs.GetRows
   rs.Close
   conn.Close

End Function
Function executeUpdate_func(query As String)
    Dim conn As ADODB.Connection
    Dim ws As Worksheet
    Dim strId As String
        
        strId = pngDBloginId '//전역변수호출
        
    Set conn = New ADODB.Connection
    
    Dim svr_info As String
    
    svr_info = svrInformation_func()
            
    conn.ConnectionString = svr_info
    
    conn.Open
    
    '//서버연결
    'strSQL = "INSERT INTO " & tableString & "(NAME, SPEC, INGREDIENT, UNIT) VALUES('" & nameString & "','" & specString & "','" & ingredientString & "','" & unitString & "');"
    
        
        Debug.Print ("insert : " + query)
    
    conn.Execute query
    conn.Close
    
    Set conn = Nothing
End Function
