<%@ LANGUAGE=VBSCript  CODEPAGE=949 TRANSACTION=Required %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Call CheckVersion(lgKeyStream(1), lgKeyStream(2))	' 2005-03-11 버전관리기능 추가 
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002) 
                                                          '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    Call SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

        strWhere = " a.co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
		strWhere = strWhere & "  and a.fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
		strWhere = strWhere & "  and a.rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
		strWhere = strWhere & "  and a.acct_cd like " &  FilterVar(Trim(lgKeyStream(3)),"'%'","S")
		strWhere = strWhere & "  and a.doc_type2 like " &  FilterVar(Trim(lgKeyStream(4)),"'%'","S")
		
		if Trim(lgKeyStream(5)) = "1" then
		   strWhere = strWhere & "  and doc_amt <= 50000 "
		   
		elseif  Trim(lgKeyStream(5)) = "2" then
		
		   strWhere = strWhere & "  and doc_amt > 50000 "
		end if
		   strWhere = strWhere & "  and a.doc_type like " &  FilterVar(Trim(lgKeyStream(6)),"'%'","S")
		   strWhere = strWhere & "  and a.co_cd =b.co_cd and a.fisc_year = b.fisc_year and a.rep_type = b.rep_type "
		   strWhere = strWhere & "  and a.acct_cd =b.acct_cd and match_cd = '10' "

    Call SubMakeSQLStatements("MR",strWhere,"X","")                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_GP_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_CD"))
			lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
            lgstrData = lgstrData & Chr(11) & UniConvDateDbToCompany(lgObjRs("DOC_DT"),"")
			lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DOC_AMT"), ggAmtOfMoney.DecPoint, 0)
			if  lgObjRs("DOC_TYPE2") = "Y" then
				lgstrData = lgstrData & Chr(11) & 1
			else
			    lgstrData = lgstrData & Chr(11) & 0
			end if
			if  lgObjRs("DOC_TYPE") = "Y" then
				lgstrData = lgstrData & Chr(11) & 1
			else
			    lgstrData = lgstrData & Chr(11) & 0
			end if
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOC_DESC"))
 
 
 
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

   Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
  

	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data

        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
        
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    if	arrColVal(7)="0" then
		arrColVal(7) = "N"
    Else
		arrColVal(7) = "Y"
    End if
    
    if  arrColVal(8)="0" then
        arrColVal(8) = "N"
    Else
        arrColVal(8) = "Y"
    End if
                    
                    'lgStrSQL = "Declare @seq smallint "
					'lgStrSQL = lgStrSQL	& " Select @seq = isnull(max(seq_no),0) + 1 from TB_WORK_3 "
					'lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
					'lgStrSQL = lgStrSQL & "		and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
					'lgStrSQL = lgStrSQL & "		and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")
										
					lgStrSQL = lgStrSQL	& " Insert into TB_WORK_3 (CO_CD,FISC_YEAR,REP_TYPE,ACCT_CD,ACCT_NM,DOC_DT,DOC_AMT,CREDIT_DEBIT,DOC_DESC,DOC_TYPE2,DOC_TYPE,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) Values ("
					lgStrSQL = lgStrSQL	&	FilterVar(Trim(UCase(lgKeyStream(0))),"","S")      & ","			
					lgStrSQL = lgStrSQL	&	FilterVar(Trim(UCase(lgKeyStream(1))),"","S")      & ","	
					lgStrSQL = lgStrSQL	&	FilterVar(Trim(UCase(lgKeyStream(2))),"","S")      & ","	
					'lgStrSQL = lgStrSQL	& " @seq"											        & ","			
					lgStrSQL = lgStrSQL	&	FilterVar(Ucase(arrColVal(2)),"''","S")					& ","
					lgStrSQL = lgStrSQL	&	FilterVar(Ucase(arrColVal(3)),"''","S")					& ","
					lgStrSQL = lgStrSQL	&	FilterVar(Ucase(arrColVal(4)),"''","S")					& ","
					lgStrSQL = lgStrSQL	&	UniConvNum(arrColVal(5),0)								& ","	
					lgStrSQL = lgStrSQL	&	"'DR' ,"			
				    lgStrSQL = lgStrSQL	&	FilterVar(Ucase(arrColVal(6)),"''","S")					& ","
				    lgStrSQL = lgStrSQL	&	FilterVar(Ucase(arrColVal(7)),"''","S")					& ","			
				    lgStrSQL = lgStrSQL	&	FilterVar(Ucase(arrColVal(8)),"''","S")					& ","	
					lgStrSQL = lgStrSQL &	FilterVar(gUsrId,"''","S")								& "," 
				    lgStrSQL = lgStrSQL &	"getdate()"					& "," 
					lgStrSQL = lgStrSQL &	FilterVar(gUsrId,"''","S")								& "," 
					lgStrSQL = lgStrSQL &	"getdate()"
					lgStrSQL = lgStrSQL &	")"

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear  
    
     if	arrColVal(7)="0" then
		arrColVal(7) = "N"
    Else
		arrColVal(7) = "Y"
    End if
    
    if  arrColVal(8)="0" then
        arrColVal(8) = "N"
    Else
        arrColVal(8) = "Y"
    End if

    
    lgStrSQL = " Update TB_WORK_3 set"
    lgStrSQL = lgStrSQL & "		DOC_DT  = " &  FilterVar(Trim(UCase(arrColVal(4))),"","S") & ","    
    lgStrSQL = lgStrSQL & "		DOC_AMT  = " &  UniConvNum(arrColVal(5),0) & ","    
    lgStrSQL = lgStrSQL & "		DOC_DESC  = " &  FilterVar(Trim(UCase(arrColVal(6))),"","S") & ","   
    lgStrSQL = lgStrSQL & "		doc_type2  = " &  FilterVar(Trim(UCase(arrColVal(7))),"","S") & ","    
    lgStrSQL = lgStrSQL & "		doc_type   = " &  FilterVar(Trim(UCase(arrColVal(8))),"","S") & ","                  
    lgStrSQL = lgStrSQL & "		UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & "		UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "		and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "		and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")
    lgStrSQL = lgStrSQL & "		and acct_cd =" &   FilterVar(Trim(UCase(arrColVal(3))),"","S") 
    lgStrSQL = lgStrSQL & "	    and seq_no =" &  UNIConvNum(arrColVal(2),0)
    


    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next
    Err.Clear
     
     
     lgStrSQL = "Delete From TB_WORK_3 "

    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "		and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "		and rep_type =" & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")
    lgStrSQL = lgStrSQL & "	    and seq_no =" &  UNIConvNum(arrColVal(2),0)
    


    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "SELECT TOP " & iSelCount  & "  A.SEQ_NO , dbo.ufn_GetCodeName('W1056',ACCT_GP_CD) ACCT_GP_NM , "
                       lgStrSQL = lgStrSQL & "                   A.ACCT_CD, B.ACCT_NM , DOC_DT, CASE WHEN CREDIT_DEBIT = 'CR' THEN DOC_AMT * -1  ELSE DOC_AMT END DOC_AMT ,"
                       lgStrSQL = lgStrSQL & "                    DOC_TYPE2, DOC_TYPE,DOC_DESC  "
                       lgStrSQL = lgStrSQL & " FROM  TB_WORK_3 a ,TB_ACCT_MATCH b"
                       lgStrSQL = lgStrSQL & " WHERE  " & pComp & pCode 
                       lgStrSQL = lgStrSQL & " ORDER BY A.ACCT_CD , A.ACCT_NM  , A.DOC_DT "

           End Select             
    End Select
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	