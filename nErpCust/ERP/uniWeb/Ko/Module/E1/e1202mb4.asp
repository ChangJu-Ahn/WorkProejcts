  <%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd_uniSIMS

    Dim emp_no
    Dim name
    Dim dept_nm
	Dim gDecimal_day
	Dim gDecimal_time
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space

    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        
   'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    ' mode = UID_M0001                                                        			 '☜: Query
    Call SubBizQuery()
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1,iRet
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call get_decimal()
	
    Call SubMakeSQLStatements("1")                                       '☜ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then
        %>
           <Script Language=vbscript>
            With parent.frm1
                .name.Value			= "<%=ConvSPChars(lgObjRs("name"))%>"
                .emp_no.Value		= "<%=ConvSPChars(lgObjRs("Emp_no"))%>"
                .grade.Value		= "<%=ConvSPChars(lgObjRs("PAY_GRD1_NM"))%>"
                .dept_cd.Value		= "<%=ConvSPChars(lgObjRs("DEPT_NM"))%>"

                .pay_tot.Value		= "<%=UNINumClientFormat(lgObjRs("PROV_TOT_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
                .sub_tot.Value		= "<%=UNINumClientFormat(lgObjRs("SUB_TOT_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
                .real.Value			= "<%=UNINumClientFormat(lgObjRs("REAL_PROV_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
            End With          
           </Script>   
        <%		
    End If    

    Call SubCloseRs(lgObjRs)
    
    Call SubMakeSQLStatements("2")   
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then
        Do While Not lgObjRs.EOF
            If Trim(lgObjRs("YEAR_TYPE")) = "1" Then
            %>
               <Script Language=vbscript>
                With parent.frm1
                    .YEAR_LEFT.Value	= "<%=UNINumClientFormat(lgObjRs("YEAR_LEFT"), 0, 0)%>"
                    .YEAR_PART.Value	= "<%=UNINumClientFormat(lgObjRs("YEAR_PART"), 0, 0)%>"
                    .YEAR_USE.Value		= "<%=UNINumClientFormat(lgObjRs("YEAR_USE"), 1, 0)%>"
                    .YEAR_CNT.Value		= "<%=UNINumClientFormat(lgObjRs("YEAR_CNT"), 1, 0)%>"
                    .YEAR_PER.Value		= "<%=UNINumClientFormat(lgObjRs("BAS_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
                End With          
               </Script>   
            <%		
            ElseIf Trim(lgObjRs("YEAR_TYPE")) = "2" Then
            %>    
    			<Script Language=vbscript>
    			 With parent.frm1
    			         .MONTH_LEFT.Value	= "<%=UNINumClientFormat(lgObjRs("MONTH_SAVE"), 0, 0)%>"
    			         .MONTH_CNT.Value	= "<%=UNINumClientFormat(lgObjRs("MONTH_CNT"), 1, 0)%>"
    			         .MONTH_USE.Value	= "<%=UNINumClientFormat(lgObjRs("MONTH_USE"), 1, 0)%>"
    			         .MONTH_PER.Value	= "<%=UNINumClientFormat(lgObjRs("BAS_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
    			 End With          
    			</Script> 
            <%
            End If   

		    lgObjRs.MoveNext 
        Loop 
	End If

    Call SubCloseRs(lgObjRs)
    	
    Call SubMakeSQLStatements("3")   
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then 
         Do While Not lgObjRs.EOF
            If Trim(lgObjRs("YEAR_TYPE")) = "1" Then
             %>
            <Script Language=vbscript>
                With parent.frm1
                    .YEAR_SUM.Value	= "<%=UNINumClientFormat(lgObjRs("ALLOW_SUM"), ggAmtOfMoney.DecPoint, 0)%>"
                End With          
            </Script>  
             <% 		
            ElseIf Trim(lgObjRs("YEAR_TYPE")) = "2" Then
		    %>
		    	<Script Language=vbscript>
		    	 With parent.frm1
		    	     .MONTH_SUM.Value	= "<%=UNINumClientFormat(lgObjRs("ALLOW_SUM"), ggAmtOfMoney.DecPoint, 0)%>"        
		    	 End With          
		    	</Script>                
            <%
		    End if
			lgObjRs.MoveNext
        Loop 
	End If 

    Call SubCloseRs(lgObjRs)

    Call SubMakeSQLStatements("4")                                       '☜ : Make sql statements
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then
        i = 0
        Do While Not lgObjRs.EOF
        %>
    	       <Script Language=vbscript>
    	        With parent.frm1
     	            .sub_nm<%=i%>.Value = "<%=ConvSPChars(lgObjRs("ALLOW_NM"))%>"
'        	        .sub_amt<%=i%>.Value = "<%=FormatNumber(lgObjRs("SUB_AMT"),0)%>"
        	        .sub_amt<%=i%>.Value = "<%=UNINumClientFormat(lgObjRs("SUB_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
          	  End With          
	        </Script>       
        <%				
		    lgObjRs.MoveNext
            i = i + 1
        Loop 
	End If

    Call SubCloseRs(lgObjRs)

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
    Dim strRowBak
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub    

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
'    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
'    Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
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

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText+adExcuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pSection)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

    Select Case pSection 
			'----------------basic data--------------------
      Case "1"		
   					lgStrSQL = "SELECT "
   					lgStrSQL = lgStrSQL & "     EMP_NO, "
   					lgStrSQL = lgStrSQL & "     dbo.ufn_H_GetEmpName(emp_no) name, "   					
   					lgStrSQL = lgStrSQL & "     PAY_GRD1,"
   					lgStrSQL = lgStrSQL & "		dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ",PAY_GRD1) PAY_GRD1_NM,"   					
   					lgStrSQL = lgStrSQL & "     dbo.ufn_GetDeptName(dept_cd,getdate()) DEPT_NM, "
   					lgStrSQL = lgStrSQL & "     BONUS_BAS, PROV_TOT_AMT, SUB_TOT_AMT, REAL_PROV_AMT "   					   					   					
   					lgStrSQL = lgStrSQL & " FROM HDF070T "
   					lgStrSQL = lgStrSQL & " WHERE EMP_NO =  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND PAY_YYMM  =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND PROV_TYPE =  " & FilterVar(lgKeyStream(2), "''", "S") & ""

      Case "2"
   					lgStrSQL = "SELECT "
   					lgStrSQL = lgStrSQL & " YEAR_TYPE, "   					
   					lgStrSQL = lgStrSQL & " YEAR_SAVE+YEAR_BONUS YEAR_LEFT, "
   					lgStrSQL = lgStrSQL & " YEAR_PART, "
   					lgStrSQL = lgStrSQL & " YEAR_USE, "
   					lgStrSQL = lgStrSQL & " YEAR_CNT, "
   					lgStrSQL = lgStrSQL & " MONTH_SAVE, "
   					lgStrSQL = lgStrSQL & " CASE WHEN MONTH_USE > MONTH_DUTY_CNT THEN MONTH_USE ELSE MONTH_DUTY_CNT END MONTH_USE,"
   					lgStrSQL = lgStrSQL & " MONTH_CNT, BAS_AMT "
   					lgStrSQL = lgStrSQL & " FROM HFB020T "
   					lgStrSQL = lgStrSQL & " WHERE EMP_NO =  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND PROV_YYMM  =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND PROV_TYPE =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
   					lgStrSQL = lgStrSQL & " ORDER BY YEAR_TYPE "  
				
      Case "3"  '---------- 연/월차기준금액 
   					lgStrSQL = "SELECT  b.YEAR_TYPE ,SUM(b.ALLOW) ALLOW_SUM "
   					lgStrSQL = lgStrSQL & " FROM HFB030T b, HFB020T c"
   					lgStrSQL = lgStrSQL & " WHERE b.EMP_NO =  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND c.PROV_YYMM =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND c.PROV_TYPE =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.EMP_NO    = c.EMP_NO "  
   					lgStrSQL = lgStrSQL & "	  AND b.YEAR_YYMM = c.YEAR_YYMM "
   					lgStrSQL = lgStrSQL & "	  AND b.YEAR_TYPE = c.YEAR_TYPE "
   					lgStrSQL = lgStrSQL & " GROUP BY b.YEAR_TYPE "
   					
			'---------------공제내역 
      Case "4"
   					lgStrSQL = "SELECT	b.SUB_AMT, a.ALLOW_NM"
   					lgStrSQL = lgStrSQL & " FROM HDA010T a, HDF060T b"
   					lgStrSQL = lgStrSQL & " WHERE b.EMP_NO   =  " & FilterVar(lgKeyStream(0), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.SUB_YYMM =  " & FilterVar(lgKeyStream(1), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.SUB_TYPE =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
   					lgStrSQL = lgStrSQL & "   AND b.SUB_CD = a.ALLOW_CD"
   					lgStrSQL = lgStrSQL & "   AND a.PAY_CD=" & FilterVar("*", "''", "S") & " "
   					lgStrSQL = lgStrSQL & "   AND a.CODE_TYPE=" & FilterVar("2", "''", "S") & ""
   					lgStrSQL = lgStrSQL & "	order by a.ALLOW_SEQ"
'Response.Write  "**lgStrSQL:" & lgStrSQL
				
   End Select
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
        Case "MD"
        Case "MR"
        Case "MU"
    End Select
End Sub

sub get_decimal()
	dim intRetCd
	gDecimal_day  = 0
	gDecimal_time = 0
	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("1", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if IntRetCd = true then
		gDecimal_day  = Trim(Replace(lgF0,Chr(11),""))
	end if

	IntRetCd = CommonQueryRs(" DECI_PLACE "," HDA041T "," ATTEND_TYPE = " & FilterVar("2", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	if IntRetCd = true then
		gDecimal_time = Trim(Replace(lgF0,Chr(11),""))
	end if

end sub


%>




