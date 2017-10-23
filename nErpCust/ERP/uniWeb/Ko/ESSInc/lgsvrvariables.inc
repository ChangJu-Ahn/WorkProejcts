<%'------------------------------------------------------------------------------------------------------------------------------
' Common variables for server side script
'------------------------------------------------------------------------------------------------------------------------------
Dim lgOpModeCRUD                                                                     '¢Ð:"C"(Create),"R"(Query),"U"(Update),"D"(Delete)
Dim lgPrevNext                                                                       '¢Ð:"P"(Prev),"N"(Next
Dim lgStrPrevKeyIndex                                                                '¢Ð:"0"(First),"1"(Second),"2"(Third),"3"(...)
Dim lgObjRs                                                                          '¢Ð: RecordSet
Dim lgObjConn                                                                        '¢Ð: Connection
Dim lgObjComm                                                                        '¢Ð: Command
Dim lgStrSQL                                                                         '¢Ð: Sql statements
Dim lgIntFlgMode                                                                     '¢Ð: OPMD_CMODE,OPMD_UMODE
Dim lgErrorStatus                                                                    '¢Ð: "YES"(Error),"NO"(Success)
Dim lgKeyStream
'------------------------------------------------------------------------------------------------------------------------------
' Common variables for server side script of Multi(SpreadSheet)
'------------------------------------------------------------------------------------------------------------------------------
Dim lgLngMaxRow                                                                      '¢Ð: Total cout for fetched data
Dim lgstrData                                                                        '¢Ð: Tempory variable for spreadsheet
Dim lgstrData1                                                                       '¢Ð: Tempory variable for spreadsheet
Dim lgstrData2                                                                       '¢Ð: Tempory variable for spreadsheet
Dim lgstrData3                                                                       '¢Ð: Tempory variable for spreadsheet
Dim lgstrData4                                                                       '¢Ð: Tempory variable for spreadsheet
Dim lgstrData5                                                                       '¢Ð: Tempory variable for spreadsheet
Dim lgMaxCount                                                                       '¢Ð: Fetch count at a time
Dim lgErrorPos                                                                       '¢Ð: Row number When error occurs
Dim lgCurrentSpd                                                                     '¢Ð: "M"(Spread #1) "S"(Spread #2)

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