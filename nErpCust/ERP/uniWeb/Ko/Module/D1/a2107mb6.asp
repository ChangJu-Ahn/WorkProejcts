<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<% 
	Call LoadBasisGlobalInf() 


    On Error Resume Next
    Err.Clear


    Call HideStatusWnd
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    'Single
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    'Multi SpreadSheet
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveSingleCreate()
             Call SubBizDelete()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'==========================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'==========================================================================================
Sub SubBizQuery()
End Sub 
'==========================================================================================
' Name : SubBizQuery
' Desc : Date data 
'==========================================================================================
Sub SubBizSave()

    On Error Resume Next
    Err.Clear

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '☜: Update
              Call SubBizSaveSingleUpdate()
    End Select
 
End Sub 

'==========================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'==========================================================================================
Sub SubBizDelete()
 Dim strTrans_Type
    Dim strAcct_Cd

    strTrans_Type = Trim(Request("txtTransType"))
    strAcct_Cd  = Trim(Request("txtAcctCd"))

    On Error Resume Next
    Err.Clear

 '---------- Developer Coding part (Start) ---------------------------------------------------------------
 'A developer must define field to update record
 '--------------------------------------------------------------------------------------------------------
 ' ==> A_JNL_CTRL_ASSN = > DELETE (해당조건에 맞는 레코드 를 삭제)
 '  DELETE 
 '  FROM A_JNL_CTRL_ASSN
 '  WHERE ACCT_CD =  '11100007'
 '  AND TRANS_TYPE = 'AP001'
 '  AND NOT  EXISTS
 '    (SELECT CTRL_CD,ACCT_CD, 'AP001'
 '    FROM A_ACCT_CTRL_ASSN  (NOLOCK)
 '    WHERE ACCT_CD =  '11100007'
 '    AND CTRL_CD =  A_JNL_CTRL_ASSN.CTRL_CD)



	lgStrSQL = "DELETE"
	lgStrSQL = lgStrSQL & "	FROM A_JNL_CTRL_ASSN "
	lgStrSQL = lgStrSQL & " WHERE ACCT_CD =  " & FilterVar(strAcct_Cd, "''", "S")		'11100007'
	lgStrSQL = lgStrSQL & " AND TRANS_TYPE =  " & FilterVar(strTrans_Type, "''", "S")	'AP001'
	lgStrSQL = lgStrSQL & " AND CTRL_CD NOT IN"
	lgStrSQL = lgStrSQL & " (SELECT CTRL_CD "
	lgStrSQL = lgStrSQL & " FROM A_ACCT_CTRL_ASSN  (NOLOCK)"
	lgStrSQL = lgStrSQL & " WHERE ACCT_CD =  " & FilterVar(strAcct_Cd, "''", "S") '11100007'
	lgStrSQL = lgStrSQL & " ) "

	'Call ServerMesgBox(lgStrSQL, vbInformation, I_MKSCRIPT)
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
 Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub


'==========================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'==========================================================================================
Sub SubBizSaveSingleCreate()
    Dim strTrans_Type
    Dim strSeq
    Dim strJnlCd
    Dim strDrCrFg
    Dim strAcct_Cd


	strTrans_Type = Trim(Request("txtTransType"))
	strSeq   = UNIConvNum(Trim(Request("txtFormSeq")),0)
	strJnlCd  = Trim(Request("txtJnlCd"))
	strDrCrFg  = Trim(Request("txtDrCrFgCd"))
	strAcct_Cd  = Trim(Request("txtAcctCd"))

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

 '==> A_JNL_CTRL_ASSN = > INSERT (해당조건에 맞는 레코드 를 삽입)
 '==> INSERT INTO A_JNL_CTRL_ASSN
 ' SELECT  'AP001' , '3', '20', 'DR', ACCT_CD, CTRL_CD,'','','','','','','','','','','','','','unierp',getdate(),'unierp',getdate()
 ' FROM A_ACCT_CTRL_ASSN  (NOLOCK)
 ' WHERE ACCT_CD =  '11100007'
 ' AND NOT  EXISTS
 '   (SELECT CTRL_CD,ACCT_CD ,TRANS_TYPE
 '   FROM A_JNL_CTRL_ASSN  (NOLOCK)
 '   WHERE ACCT_CD =  '11100007'
 '   AND CTRL_CD =  A_ACCT_CTRL_ASSN.CTRL_CD
 '   AND TRANS_TYPE = 'AP001'  )

    lgStrSQL = "INSERT INTO A_JNL_CTRL_ASSN"
	lgStrSQL = lgStrSQL & " SELECT  "	
	lgStrSQL = lgStrSQL & FilterVar(strTrans_Type, "''", "S") & ","	'AP001' 
	lgStrSQL = lgStrSQL & FilterVar(strSeq,"''", "") & ","	'3
	lgStrSQL = lgStrSQL & FilterVar(strJnlCd, "''", "S") & ","	'20'
	lgStrSQL = lgStrSQL & FilterVar(strDrCrFg, "''", "S") & ","  'DR'
	lgStrSQL = lgStrSQL & " ACCT_CD, CTRL_CD,'','','','','','','','','','','','','', "
	lgStrSQL = lgStrSQL &	FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL &	"getdate()," 
    lgStrSQL = lgStrSQL &	FilterVar(gUsrId, "''", "S") & "," 
    lgStrSQL = lgStrSQL &	"getdate()" 
    lgStrSQL = lgStrSQL & " FROM A_ACCT_CTRL_ASSN  (NOLOCK)"
	lgStrSQL = lgStrSQL & " WHERE ACCT_CD = " & FilterVar(strAcct_Cd, "''", "S") '11100007'
	lgStrSQL = lgStrSQL & " AND CTRL_CD NOT IN"
	lgStrSQL = lgStrSQL & "	(SELECT CTRL_CD"
	lgStrSQL = lgStrSQL & "	FROM A_JNL_CTRL_ASSN  (NOLOCK) "
	lgStrSQL = lgStrSQL & "	WHERE ACCT_CD =  " & FilterVar(strAcct_Cd, "''", "S") '11100007'
	lgStrSQL = lgStrSQL & "	AND TRANS_TYPE = " & FilterVar(strTrans_Type, "''", "S") 'AP001'
	lgStrSQL = lgStrSQL & "	AND SEQ = " & FilterVar(strSeq,"0", "N") 'SEQ'
	lgStrSQL = lgStrSQL & ")"

	'Call ServerMesgBox(lgStrSQL, vbInformation, I_MKSCRIPT)
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
 Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'==========================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'==========================================================================================
Sub SubBizSaveSingleUpdate()
End Sub

'==========================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'==========================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
End Sub
'==========================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'==========================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
End Sub


'==========================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'==========================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
End Sub


'==========================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'==========================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
End Sub

'==========================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'==========================================================================================
Sub CommonOnTransactionCommit()
End Sub

'==========================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'==========================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'==========================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'==========================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"
End Sub

'==========================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'==========================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "SD"
     If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990025", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990025", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MR"
        Case "MU"
    End Select
End Sub

%>

<Script Language=vbscript>
	Parent.BeforeDbQuery_TwoOk(<%=Request("txtRow")%>)
</Script>
