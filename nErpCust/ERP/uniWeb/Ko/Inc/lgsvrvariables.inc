<%'------------------------------------------------------------------------------------------------------------------------------
' Common variables for server side script
'------------------------------------------------------------------------------------------------------------------------------
Dim lgOpModeCRUD                                                                     '��:"C"(Create),"R"(Query),"U"(Update),"D"(Delete)
Dim lgPrevNext                                                                       '��:"P"(Prev),"N"(Next
Dim lgStrPrevKeyIndex                                                                '��:"0"(First),"1"(Second),"2"(Third),"3"(...)
Dim lgObjRs                                                                          '��: RecordSet
Dim lgObjConn                                                                        '��: Connection
Dim lgObjComm                                                                        '��: Command
Dim lgStrSQL                                                                         '��: Sql statements
Dim lgIntFlgMode                                                                     '��: OPMD_CMODE,OPMD_UMODE
Dim lgErrorStatus                                                                    '��: "YES"(Error),"NO"(Success)
Dim lgKeyStream
'------------------------------------------------------------------------------------------------------------------------------
' Common variables for server side script of Multi(SpreadSheet)
'------------------------------------------------------------------------------------------------------------------------------
Dim lgLngMaxRow                                                                      '��: Total cout for fetched data
Dim lgstrData                                                                        '��: Tempory variable for spreadsheet
Dim lgstrData1                                                                       '��: Tempory variable for spreadsheet
Dim lgstrData2                                                                       '��: Tempory variable for spreadsheet
Dim lgstrData3                                                                       '��: Tempory variable for spreadsheet
Dim lgstrData4                                                                       '��: Tempory variable for spreadsheet
Dim lgstrData5                                                                       '��: Tempory variable for spreadsheet
Dim lgMaxCount                                                                       '��: Fetch count at a time
Dim lgErrorPos                                                                       '��: Row number When error occurs
Dim lgCurrentSpd                                                                     '��: "M"(Spread #1) "S"(Spread #2)

'------------------------------------------------------------------------------------------------------------------------------
' Common constants for server side script
'------------------------------------------------------------------------------------------------------------------------------

Const C_EQGT = ">="
Const C_EQ   = "="
Const C_EQLT = "<="
Const C_LT   = "<"
Const C_GT   = ">"
Const C_LIKE = "LIKE"
%>