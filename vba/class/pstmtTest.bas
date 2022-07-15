Attribute VB_Name = "pstmtTest"
Sub pstmtInsert()
    '//make by.icurfer
    Dim pstmt As New databaseControll
    Dim insert_query As String
    Dim colArry As New Collection
    
    pstmt.prepareStatment ("INSERT INTO AD_INFO_ECOUNT_LOGIN(`time`, `key`) VALUES(?, ?)")
            
    colArry.Add "2", "key1"
    colArry.Add "3", "key2"
    
    insert_query = pstmt.setString(colArry)
    pstmt.executeUpdate_func (insert_query)
    
    Debug.Print ("stop")
    
    
End Sub
Sub pstmtSelect()
    '//make by.icurfer
    Dim pstmt As New databaseControll
    Dim select_query As String
    Dim colArry As New Collection
    Dim resultSet As Variant
    
    '//No have terms
'    select_query = pstmt.prepareStatment("SELECT * FROM AD_INFO_ECOUNT_LOGIN")
'    resultSet = pstmt.executeQury_func(select_query)
'    Debug.Print (resultSet(0, 0))

    '//Have some terms
    pstmt.prepareStatment ("SELECT * FROM AD_INFO_ECOUNT_LOGIN WHERE `time` = ?")
    colArry.Add "2", "key1"
    select_query = pstmt.setString(colArry)
    Debug.Print (select_query)
    resultSet = pstmt.executeQury_func(select_query)
    Debug.Print (resultSet(0, 0))

End Sub
