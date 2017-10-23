<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<%Response.Buffer = True%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%    
	Dim lgStrPrevKey
	Dim txtInsur_type
	dim txtyyyymm
	dim txtEmp_No
	
	Const C_SHEETMAXROWS_D = 100

    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),1)    
    
    txtInsur_type = Request("txtInsur_type")
    txtyyyymm = Request("txtyyyymm")
    txtEmp_No = Request("txtEmp_No")

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002) 
      
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear  
    call SubBizSaveMultiDeleteHDR()                                                                      '¢Ð: Clear Error status
End Sub


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim strInsur_type
    Dim strPAY_YYMM
    Dim strWhere

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    
    Call SubMakeSQLStatements("MR",strWhere,"X","")                                 '¡Ù : Make sql statements
     
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
    
    call ListupDataGrid (lgObjRs.getRows,"","","vspdData")
    
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet
    END IF
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To ubound(arrRowVal)
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
      
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

Const C_PAY_YYMM         = 2
Const C_EMP_NO           = 3
Const C_INSUR_TYPE       = 4
Const C_INCOME_YY        = 5
Const C_INCOME_TOT_AMT   = 6
Const C_WORK_MONTH       = 7
Const C_INCOME_AVR_AMT   = 8
Const C_DATE1			 = 9
Const C_DATE2			 = 10
CONST C_ORG_CHANGE_ID    = 11

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "INSERT INTO HDB230T         ("

    lgStrSQL = lgStrSQL & " PAY_YYMM     ," 
    lgStrSQL = lgStrSQL & " EMP_NO     ," 
    lgStrSQL = lgStrSQL & " INSUR_TYPE         ," 
    lgStrSQL = lgStrSQL & " INCOME_YY     ," 
    lgStrSQL = lgStrSQL & " INCOME_TOT_AMT     ," 
    lgStrSQL = lgStrSQL & " WORK_MONTH     ," 
    lgStrSQL = lgStrSQL & " INCOME_AVR_AMT     ," 
    lgStrSQL = lgStrSQL & " MED_ACQ_DT     ," 
    lgStrSQL = lgStrSQL & " MED_LOSS_DT     ," 
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID     ," 
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO         ," 
    lgStrSQL = lgStrSQL & " ISRT_DT         ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO     ," 
    lgStrSQL = lgStrSQL & " UPDT_DT         )" 
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_PAY_YYMM)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_EMP_NO)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_INSUR_TYPE)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_INCOME_YY)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_INCOME_TOT_AMT)), "", "SNM")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_WORK_MONTH)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_INCOME_AVR_AMT)), "", "SNM")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_date1)),"''", "S")     & ","
    if arrColVal(C_date2)="" then
		lgStrSQL = lgStrSQL & " null,"
    else
		lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_date2)), "''", "S")     & ","
    end if
    
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_ORG_CHANGE_ID)),"''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
      
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "UPDATE  HDB230T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " INCOME_YY  = " & FilterVar(UCase(arrColVal(C_INCOME_YY)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & " INCOME_TOT_AMT  = " & UNIConvNum(arrColVal(C_INCOME_TOT_AMT),0)   & ","
    lgStrSQL = lgStrSQL & " INCOME_AVR_AMT  = "      & UNIConvNum(arrColVal(C_INCOME_AVR_AMT),0)   & ","
    lgStrSQL = lgStrSQL & " WORK_MONTH  = " & FilterVar(UCase(arrColVal(C_WORK_MONTH)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & " MED_ACQ_DT  = " & FilterVar(arrColVal(C_date1), "''", "S")     & ","
    lgStrSQL = lgStrSQL & " MED_LOSS_DT  = " & FilterVar(arrColVal(C_date2), "''", "S")     & ","
    lgStrSQL = lgStrSQL & " ORG_CHANGE_ID  = " & FilterVar(arrColVal(C_ORG_CHANGE_ID), "''", "S")     & ","
    
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO  = " & FilterVar(gUsrId, "''", "S")   & ","
    lgStrSQL = lgStrSQL & " UPDT_DT      = " & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " INSUR_TYPE   = "     & FilterVar(UCase(arrColVal(C_INSUR_TYPE)), "''", "S")
    lgStrSQL = lgStrSQL & " And PAY_YYMM   = " & FilterVar(UCase(arrColVal(C_PAY_YYMM)), "''", "S")
    lgStrSQL = lgStrSQL & " And EMP_NO   = "      & FilterVar(UCase(arrColVal(C_EMP_NO)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  HDB230T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " INSUR_TYPE   = "     & FilterVar(UCase(arrColVal(C_INSUR_TYPE)), "''", "S")
    lgStrSQL = lgStrSQL & " And PAY_YYMM   = " & FilterVar(UCase(arrColVal(C_PAY_YYMM)), "''", "S")
    lgStrSQL = lgStrSQL & " And EMP_NO   = "      & FilterVar(UCase(arrColVal(C_EMP_NO)), "''", "S")
   
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

Sub SubBizSaveMultiDeleteHDR()
	dim txtyyyymm
	dim txtInsuretype
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear    
                                                                        '¢Ð: Clear Error status
    txtInsuretype = Request("txtInsuretype")
    txtyyyymm = Request("txtyyyymm")
    
    lgStrSQL = "DELETE  HDB230T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " INSUR_TYPE   = "     & FilterVar(txtInsuretype, "''", "S")
    lgStrSQL = lgStrSQL & " And PAY_YYMM   = " & FilterVar(txtyyyymm, "''", "S")

   
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
        
         '  iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
          
           Select Case Mid(pDataType,2,1)
               Case "R"
              
               
                       lgStrSQL = "Select " &  " TOP " & lgStrPrevKey * C_SHEETMAXROWS_D 
                       lgStrSQL = lgStrSQL & " A.EMP_NO ,'',B.NAME, DEPT_CD,DBO.UFN_GETDEPTNAME(DEPT_CD, '2007-02-03') DEPT_NM,  "
                       lgStrSQL = lgStrSQL & " MED_ACQ_DT,MED_LOSS_DT,A.INCOME_YY, A.INCOME_TOT_AMT,A.WORK_MONTH , A.INCOME_AVR_AMT, A.INCOME_YY  ,A.INCOME_AVR_AMT"
                       lgStrSQL = lgStrSQL & " ,A.INCOME_YY  "
                       lgStrSQL = lgStrSQL & " FROM  HDB230T A "
                       lgStrSQL = lgStrSQL & " LEFT JOIN HAA010T B ON A.EMP_NO = B.EMP_NO "
                       lgStrSQL = lgStrSQL & " WHERE 1=1 "
                       lgStrSQL = lgStrSQL & " AND INSUR_TYPE =" &  FilterVar( txtInsur_type, "''", "S")
                       lgStrSQL = lgStrSQL & " AND PAY_YYMM =" &  FilterVar( txtyyyymm, "''", "S") 
                       if txtEmp_No<>"" then lgStrSQL = lgStrSQL & " AND a.EMP_NO like " &  FilterVar( txtEmp_No&"%", "''", "S") 
                       

           End Select             
    End Select
End Sub


'============================================================================
'ListupDataGrid
'============================================================================


 Sub ListupDataGrid(pArr,dataFormatCol,NFormatCol,grid)
	Dim strData
	Dim i,j,moveLine,RowCnt
	On Error resume next

	RowCnt=0
	
	
	moveLine = (lgStrPrevKey - 1) * C_SHEETMAXROWS_D
	

		for i=moveLine to uBound(pArr,2)
			RowCnt=RowCnt+1
			for j=0 to uBound(pArr,1)
			
			if inStr(dataFormatCol,"," & j&",") > 0 then
				strData = strData & Chr(11) & UniConvDateDbToCompany(pArr(j,i),"")
			elseif inStr(NFormatCol,"," & j&",") > 0 then
				strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(trim(ConvSPChars(pArr(j,i))),0)
			else
				strData = strData & Chr(11) & trim(ConvSPChars(pArr(j,i)))
			end if	
			next 
			strData =  strData & Chr(11) & i &  Chr(11) & Chr(12) 
		next 
		
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	Call .ggoOper.ClearField(Document, ""2"")	" & vbCr
		Response.Write "	.ggoSpread.Source       = .frm1."&grid&" "			& vbCr
		Response.Write "    .frm1."&grid&".Redraw = False   "                  & vbCr   
		Response.Write "	.ggoSpread.SSShowData     """ & strData	 & """" & ",""F""" & vbCr
		Response.Write "	.DbQueryOk " & vbCr 
		Response.Write  "   .frm1."&grid&".Redraw = True " & vbCr
		Response.Write "	.lgStrPrevKey  = """ & lgStrPrevKey + 1 & """" & vbCr 
		Response.Write "	.frm1.txtyyyy.year		= """ & ConvSPChars(pArr(11,0)) & """"	& vbCr
		
		if RowCnt<C_SHEETMAXROWS_D then
			Response.Write "    .lgStrPrevKey= """"  "                  & vbCr 
		
		end if
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
		
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                '.lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
            Parent.DBSaveOk
          Else
            Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
