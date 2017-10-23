<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%
	Dim lgStrPrevKey,lgStrPrevKey1
    Dim lgSpreadFlg

	Const C_SHEETMAXROWS_D = 100
	Call HideStatusWnd                                                               '☜: Hide Processing message

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB") 
    
    Dim lgSvrDateTime
    dim lgDayPoint
    dim lgTimePoint
    Dim cntEmpNo
    
    lgSvrDateTime = GetSvrDateTime

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
	lgDayPoint        = request("daypoint")
	lgTimePoint       = request("timePoint")
    lgSpreadFlg       = Request("gSpreadFlg")

	if lgSpreadFlg = "1" then
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	elseif lgSpreadFlg = "2" then            		
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	end if

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim txtGlNo
    Dim iLcNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If lgKeyStream(2) = "" then
       txtGlNo =  " " & FilterVar("%", "''", "S") & ")"
    Else
       txtGlNo =  FilterVar(lgKeyStream(2), "''", "S") & ")"
    End if
    
    If lgKeyStream(0) = "" then
       txtGlNo = txtGlNo & " AND  (WK_YYMM = " & FilterVar("%", "''", "S") & " )"
    Else
       txtGlNo = txtGlNo & "  AND  (WK_YYMM = " & FilterVar(lgKeyStream(0), "''", "S") & " )"  '☜ :WK_YYMM은 DB에 TYPE이 VARCHAAR로 되어있다..
    End if

        
    Call SubMakeSQLStatements("SR",txtGlNo,"X",C_EQ)                                  '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
		  if lgSpreadFlg = "1" then
	          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the starting data. 
		  end if
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
		  if lgSpreadFlg = "1" then       
			Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the ending data.
		  end if
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
%>
<Script Language=vbscript>

      With Parent.Frm1
             .txtEmp_no.Value  = "<%=ConvSPChars(lgObjRs("emp_no"))%>"                   'Set condition area
			 
			 call parent.txtEmp_no_Onchange()

             .txtTot_day.text          = "<%=UNINumClientFormat(lgObjRs("TOT_DAY"), 0,0)%>"
             .txtSun_day.text          = "<%=UNINumClientFormat(lgObjRs("SUN_DAY"), 0,0)%>"            
             .txtHol_day.text          = "<%=UNINumClientFormat(lgObjRs("HOL_DAY"), 0,0)%>"
             .txtAttend_day.text       = "<%=UNINumClientFormat(lgObjRs("ATTEND_DAY"), 0,0)%>"
'             .txtWeek_hol_day.text     = "<%=UNINumClientFormat(lgObjRs("SUN_DAY"), 0,0)%>" - "<%=UNINumClientFormat(lgObjRs("NON_WEEK_DAY"), 0,0)%>"
             .txtWeek_hol_day.text     = "<%=UNINumClientFormat(lgObjRs("WEEK_DAY"), 0,0)%>"
             .txtNon_week_day.text     = "<%=UNINumClientFormat(lgObjRs("NON_WEEK_DAY"), 0,0)%>"            
             .txtMargir_day_count.text = "<%=UNINumClientFormat(lgObjRs("MARGIR_day"), 4,0)%>" 
             .txtMargir_time.text      = "<%=UNINumClientFormat(lgObjRs("MARGIR_TIME"), 4,0)%>"              
             .txtWk_day.text           = "<%=UNINumClientFormat(lgObjRs("WK_DAY"), lgDayPoint,0)%>"
             .txtWork_day.text           = "<%=UNINumClientFormat(lgObjRs("WORK_DAY"), lgDayPoint,0)%>"             
             .txtWk_time.text          = "<%=UNINumClientFormat(lgObjRs("WK_TIME"), lgTimePoint,0)%>"
      End With   
</Script>       
<%
	   cntEmpNo =  FilterVar(lgObjRs("emp_no"), "''", "S") & ")"
       Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
       Call SubBizQueryMulti(txtGlNo)
      
    End If
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
     	 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = UNIcdbl(Request("txtFlgMode"), 0)                                       '☜: Read Operayion Mode (CREATE, UPDATE)

	Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '☜: Update
              Call SubBizSaveSingleUpdate()
    End Select

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  HCA090T"            '☜ single
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO     = " & FilterVar(UCase(lgKeyStream(2)), "''", "S")
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)
    
    lgStrSQL = "DELETE  HCA070T"            '☜ multi 
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " DILIG_YYMM     = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO     = " & FilterVar(UCase(lgKeyStream(2)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pKey1)
    Dim iDx
    Dim iLoopMax
    Dim strWhere
	dim d_type
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    if lgSpreadFlg = "1" then
		d_type = "" & FilterVar("2", "''", "S") & ""
	else
		d_type = "" & FilterVar("1", "''", "S") & " "
	end if
	
    If lgKeyStream(2) = "" then
       strWhere =  " " & FilterVar("%", "''", "S") & ")"
    Else
       strWhere = cntEmpNo 
    End if
    
    If lgKeyStream(0) = "" then
		strWhere = strWhere & " and (b.DILIG_YYMM LIKE " & FilterVar("%", "''", "S") & " )"
		strWhere = strWhere & " and a.DILIG_TYPE = " & d_type 
		strWhere = strWhere & " ORDER BY a.DILIG_SEQ ASC "
    Else
		strWhere = strWhere & " and (b.DILIG_YYMM LIKE " & FilterVar(lgKeyStream(0), "''", "S") & " ))"  '☜ :DILIG_YYMM은 DB에 TYPE이 VARCHAAR로 되어있다..
		strWhere = strWhere & " and a.DILIG_TYPE = " & d_type
		strWhere = strWhere & " ORDER BY a.DILIG_SEQ ASC "
    End if
                        
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                                   '☆: Make sql statements
   
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        lgStrPrevKey1 = ""
    Else
     
    	if lgSpreadFlg = "1" then
			Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		elseif lgSpreadFlg = "2" then            		
			Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)
		end if

        lgstrData = ""
        
        iDx = 1
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DILIG_NM"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_CNT"),3,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_HH"), 3,0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("DILIG_MM"), 3,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DAY_TIME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BAS_MARGIR"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WK_DAY"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ATTEND_DAY"))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
            
		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
				if lgSpreadFlg = "1" then            
		               lgStrPrevKey = lgStrPrevKey + 1
		        elseif lgSpreadFlg = "2" then            
		               lgStrPrevKey1 = lgStrPrevKey1 + 1		        
		        end if
               Exit Do
            End If   
               
        Loop 
    End If

    If iDx <= C_SHEETMAXROWS_D Then
		if lgSpreadFlg = "1" then            
		       lgStrPrevKey = ""
		elseif lgSpreadFlg = "2" then            
		       lgStrPrevKey1 = ""
		end if
    End If   
	if lgSpreadFlg = "1" then            
			lgstrData1 =lgstrData
	elseif lgSpreadFlg = "2" then            
	       lgstrData2 = lgstrData
	end if

    Call SubHandleError("MR",lgObjRs,Err)
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
Sub SubBizSaveSingleCreate()
    Dim txtGlNo
    Dim intAttendDay

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    txtGlNo = FilterVar(lgKeyStream(0), "''", "S")
    intAttendDay = (UNIcdbl(Request("txtTot_day"),0) - UNIcdbl(Request("txtSun_day"),0)  -  UNIcdbl(Request("txtHol_day"),0))

    lgStrSQL = "INSERT INTO HCA090T("
    lgStrSQL = lgStrSQL & " WK_YYMM      ," 
    lgStrSQL = lgStrSQL & " EMP_NO       ," 
    lgStrSQL = lgStrSQL & " TOT_DAY      ," 
    lgStrSQL = lgStrSQL & " SUN_DAY      ,"
    lgStrSQL = lgStrSQL & " HOL_DAY      ," 
    lgStrSQL = lgStrSQL & " ATTEND_DAY   ," 
    lgStrSQL = lgStrSQL & " WEEK_DAY     ,"
    lgStrSQL = lgStrSQL & " NON_WEEK_DAY ," 
    lgStrSQL = lgStrSQL & " MARGIR_DAY   ," 
    lgStrSQL = lgStrSQL & " MARGIR_TIME  ," 
    lgStrSQL = lgStrSQL & " WK_DAY       ,"
    lgStrSQL = lgStrSQL & " WORK_DAY       ,"    
    lgStrSQL = lgStrSQL & " WK_TIME      ," 
    lgStrSQL = lgStrSQL & " WEEK_TIME    ," 
    lgStrSQL = lgStrSQL & " HOL_TIME     ,"
    lgStrSQL = lgStrSQL & " ATTEND_TIME  ,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO  ," 
    lgStrSQL = lgStrSQL & " ISRT_DT      ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO  ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")							 & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtEmp_No")), "''", "S")				 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtTot_day"),0)		                    	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtSun_day"),0)		                    	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtHol_day"),0)		                    	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(intAttendDay,0)	                    				 & ","	    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtweek_hol_day"),0)		                   	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtNon_week_day"),0)		                   	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtMargir_day_count"),0)		               	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtMargir_time"),0)		                	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtWk_day"),0)		                    	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtWork_day"),0)		                    	 & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtWk_time"),0)		                    	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtweek_hol_day"),0)		                   	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtHol_day"),0)		                     	 & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtAttend_day"),0)		                   	 & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")											 & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")									 & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")										     & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()


    Dim intAttendDay, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    intAttendDay = 0
    
    Call CommonQueryRs(" ATTEND_DAY, SUN_DAY , HOL_DAY,TOT_DAY "," HCA090T "," WK_YYMM = " & FilterVar(lgKeyStream(0), "''", "S") & " AND EMP_NO =  " & FilterVar(Request("txtEmp_no"), "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If ( UNIcdbl(Replace(lgF1,Chr(11),""),0) < UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) < UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0) - (UNIcdbl(Request("txtSun_day"),0) - UNIcdbl(Replace(lgF1,Chr(11),""),0) )  - (UNIcdbl(Request("txtHol_day"),0) - UNIcdbl(Replace(lgF2,Chr(11),""),0) )
    ElseIf ( UNIcdbl(Replace(lgF1,Chr(11),""),0) > UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) > UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0) + ( UNIcdbl(Replace(lgF1,Chr(11),""),0)  -  UNIcdbl(Request("txtSun_day"),0))  + ( UNIcdbl(Replace(lgF2,Chr(11),""),0) - UNIcdbl(Request("txtHol_day"),0))
    ElseIf ( UNIcdbl(Replace(lgF1,Chr(11),""),0) < UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) > UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0) - (UNIcdbl(Request("txtSun_day"),0)  - UNIcdbl(Replace(lgF1,Chr(11),""),0))  + ( UNIcdbl(Replace(lgF2,Chr(11),""),0) - UNIcdbl(Request("txtHol_day"),0))
    ElseIf ( UNIcdbl(Replace(lgF1,Chr(11),""),0) > UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) < UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0) + ( UNIcdbl(Replace(lgF1,Chr(11),""),0)  -  UNIcdbl(Request("txtSun_day"),0))  - (  UNIcdbl(Request("txtHol_day"),0) - UNIcdbl(Replace(lgF2,Chr(11),""),0))
    ElseIf ( UNIcdbl(Replace(lgF1,Chr(11),""),0) = UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) > UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0) + ( UNIcdbl(Replace(lgF2,Chr(11),""),0)  -  UNIcdbl(Request("txtHol_day"),0))
    ElseIf ( UNIcdbl(Replace(lgF1,Chr(11),""),0) = UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) < UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0)  - (  UNIcdbl(Request("txtHol_day"),0)  - UNIcdbl(Replace(lgF2,Chr(11),""),0))
    ElseIf ( UNIcdbl(Replace(lgF1,Chr(11),""),0) < UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) = UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0) - (UNIcdbl(Request("txtSun_day"),0)  - UNIcdbl(Replace(lgF1,Chr(11),""),0))
    ElseIf ( UNIcdbl(Replace(lgF1,Chr(11),""),0) > UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) = UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0) + ( UNIcdbl(Replace(lgF1,Chr(11),""),0)  -  UNIcdbl(Request("txtSun_day"),0)) 
    ElseIf ( UNIcdbl(Replace(lgF1,Chr(11),""),0) = UNIcdbl(Request("txtSun_day"),0) ) AND ( UNIcdbl(Replace(lgF2,Chr(11),""),0) = UNIcdbl(Request("txtHol_day"),0) ) Then
        intAttendDay = UNIcdbl(Request("txtAttend_day"),0)
    End if

    If UNIcdbl(Replace(lgF3,Chr(11),""),0) < UNIcdbl(Request("txtTot_day"),0) Then
        intAttendDay = intAttendDay + (UNIcdbl(Request("txtTot_day"),0) - UNIcdbl(Replace(lgF3,Chr(11),""),0) )
    ElseIf UNIcdbl(Replace(lgF3,Chr(11),""),0) > UNIcdbl(Request("txtTot_day")) Then
        intAttendDay = intAttendDay - (UNIcdbl(Replace(lgF3,Chr(11),""),0)  -  UNIcdbl(Request("txtTot_day"),0) )
    Else
    End if
   '근무일수(ATTEND_DAY) = 총일수 - 휴일 - 일요일 - (근태테이블(HCA070T)에서 ATTEND_DAY가 "N"인 것은 그 횟수)
    lgStrSQL = "UPDATE  HCA090T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " TOT_DAY      = " & UNIConvNum(Request("txtTot_day"),0)			 & ","
    lgStrSQL = lgStrSQL & " SUN_DAY      = " & UNIConvNum(Request("txtSun_day"),0)			 & ","
    lgStrSQL = lgStrSQL & " HOL_DAY      = " & UNIConvNum(Request("txtHol_day"),0)			 & ","
    lgStrSQL = lgStrSQL & " WEEK_DAY     = " & UNIConvNum(Request("txtweek_hol_day"),0)	 & ","
    lgStrSQL = lgStrSQL & " ATTEND_DAY   = " & UNIConvNum(intAttendDay,0)					 & ","	    
    lgStrSQL = lgStrSQL & " NON_WEEK_DAY = " & UNIConvNum(Request("txtNon_week_day"),0)	 & ","
    lgStrSQL = lgStrSQL & " MARGIR_DAY   = " & UNIConvNum(Request("txtMargir_day_count"),0) & ","
    lgStrSQL = lgStrSQL & " MARGIR_TIME  = " & UNIConvNum(Request("txtMargir_time"),0)      & ","
    lgStrSQL = lgStrSQL & " WEEK_TIME    = " & UNIConvNum(Request("txtweek_hol_day"),0)     & ","
    lgStrSQL = lgStrSQL & " HOL_TIME     = " & UNIConvNum(Request("txtHol_day"),0)			 & ","
    lgStrSQL = lgStrSQL & " ATTEND_TIME  = " & UNIConvNum(Request("txtAttend_day"),0)       & ","
    lgStrSQL = lgStrSQL & " WK_DAY       = " & UNIConvNum(Request("txtWk_day"),0)			 & ","  
    lgStrSQL = lgStrSQL & " WORK_DAY       = " & UNIConvNum(Request("txtWork_day"),0)			 & ","
    lgStrSQL = lgStrSQL & " WK_TIME      = " & UNIConvNum(Request("txtWk_time"),0)			 & "," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO  = " & FilterVar(gUsrId, "''", "S")                     & "," 
    lgStrSQL = lgStrSQL & " UPDT_DT      = " & FilterVar(lgSvrDateTime, "''", "S") 
    lgStrSQL = lgStrSQL & " WHERE  "
    lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(lgKeyStream(2), "''", "S")
     
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
   
   If  arrColVal(9) = "Y" Then                        '☜  insert 시 차감시간과 차감일수를 더해준다..........
        If  arrColVal(8) = "1" Then 
             lgStrSQL = "UPDATE  HCA090T"
             lgStrSQL = lgStrSQL & " SET " 
             lgStrSQL = lgStrSQL & " MARGIR_DAY  = MARGIR_DAY + " & UNIcdbl(arrColVal(5),0) & ","
             lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME + " & (UNIcdbl(arrColVal(5),0) * 8)
             lgStrSQL = lgStrSQL & " WHERE  "
             lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
             lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
        Elseif arrColVal(8) = "2" Then 
             lgStrSQL = "UPDATE  HCA090T"
             lgStrSQL = lgStrSQL & " SET " 
             lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME + " & UNIcdbl(arrColVal(6),0) & " + " & (UNIcdbl(arrColVal(7),0)/60)
             lgStrSQL = lgStrSQL & " WHERE  "
             lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
             lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
        Elseif arrColVal(8) = "3" Then   
             lgStrSQL = "UPDATE  HCA090T"
             lgStrSQL = lgStrSQL & " SET " 
             lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME + " & (UNIcdbl(arrColVal(5),0) * 4)
             lgStrSQL = lgStrSQL & " WHERE  "
             lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
             lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
        End if 
   Else
   End if
   If  arrColVal(10) = "N" Then           '근태코드 테이블에서 wk_day = 'N'인것 근무일수를 minus
        lgStrSQL = "UPDATE  HCA090T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " WORK_DAY  = (TOT_DAY  - " & UNIcdbl(arrColVal(5),0) & ")"
        lgStrSQL = lgStrSQL & " WHERE  "
        lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
        lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
        Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   Else
   End if
   
   If  arrColVal(11) = "N" Then           '출근일수 구분이 'N'이면 출근일수에서 minus           
    '근무일수(ATTEND_DAY) = 총일수 - 휴일 - 일요일 - (근태테이블(HCA070T)에서 ATTEND_DAY가 "N"인 것은 그 횟수)
        lgStrSQL = "UPDATE  HCA090T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " ATTEND_DAY  = " & int(arrColVal(12)) - int(arrColVal(5))
        lgStrSQL = lgStrSQL & " WHERE  "
        lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
        lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        
        lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
        Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   Else
   End if
   
    lgStrSQL = "INSERT INTO HCA070T("
    lgStrSQL = lgStrSQL & " DILIG_YYMM     ," 
    lgStrSQL = lgStrSQL & " EMP_NO         ," 
    lgStrSQL = lgStrSQL & " DILIG_CD       ," 
    lgStrSQL = lgStrSQL & " DILIG_CNT      ," 
    lgStrSQL = lgStrSQL & " DILIG_HH       ," 
    lgStrSQL = lgStrSQL & " DILIG_MM       ," 
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO    ," 
    lgStrSQL = lgStrSQL & " ISRT_DT        ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO    ," 
    lgStrSQL = lgStrSQL & " UPDT_DT        )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(arrColVal(2), "''", "S")           & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(3)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(4)), "''", "S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(5),0)                        & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(6),0)                        & ","
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)                        & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                          & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")                 & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                          & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

   If  arrColVal(9) = "Y" Then                   '☜  update시 조건에 맞게 차감시간과 차감일수를 조절해 준다.
         lgStrSQL = "SELECT  DILIG_CNT, DILIG_HH , DILIG_MM " 
         lgStrSQL = lgStrSQL & " FROM  HCA070T "
         lgStrSQL = lgStrSQL & " WHERE DILIG_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
         lgStrSQL = lgStrSQL & " AND EMP_NO     = " & FilterVar(UCase(arrColVal(3)), "''", "S")
         lgStrSQL = lgStrSQL & " AND DILIG_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        Call FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")
        
        If  arrColVal(8) = "1" Then                                       'DAY_TIME이 1,2,3인 경우 횟수에대해.....
             If  UNIcdbl(lgObjRs("DILIG_CNT"),0) < UNIcdbl(arrColVal(5),0) Then     '갱신값이 원래값보다 큰 경우 
                 lgStrSQL = "UPDATE  HCA090T"
                 lgStrSQL = lgStrSQL & " SET " 
                 lgStrSQL = lgStrSQL & " MARGIR_DAY  = MARGIR_DAY + " & (UNIcdbl(arrColVal(5),0) - UNIcdbl(lgObjRs("DILIG_CNT"),0)) & ","
                 lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME + " &( (UNIcdbl(arrColVal(5),0) - UNIcdbl(lgObjRs("DILIG_CNT"),0))*8)
                 lgStrSQL = lgStrSQL & " WHERE  "
                 lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
                 lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
                lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
             Elseif UNIcdbl(lgObjRs("DILIG_CNT")) = UNIcdbl(arrColVal(5)) Then
             Else                                                          '갱신값이 원래값보다 작 경우 
                 lgStrSQL = "UPDATE  HCA090T"
                 lgStrSQL = lgStrSQL & " SET " 
                 lgStrSQL = lgStrSQL & " MARGIR_DAY = MARGIR_DAY - " & (UNIcdbl(lgObjRs("DILIG_CNT"),0) - UNIcdbl(arrColVal(5),0)) & ","
                 lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME - " & ((UNIcdbl(lgObjRs("DILIG_CNT"),0) - UNIcdbl(arrColVal(5),0))*8)
                 lgStrSQL = lgStrSQL & " WHERE  "
                 lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
                 lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
                lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
            End if 
        Elseif arrColVal(8) = "2" Then                                      'DAY_TIME이 2(시간)인 경우에 대해............       
             If  UNIcdbl(lgObjRs("DILIG_HH")) < UNIcdbl(arrColVal(6)) Then
                 lgStrSQL = "UPDATE  HCA090T"
                 lgStrSQL = lgStrSQL & " SET " 
                 lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME +  ((" & UNIcdbl(arrColVal(6),0) & " + " & (UNIcdbl(arrColVal(7),0)/60) & ") - (" & UNIcdbl(lgObjRs("DILIG_HH"),0) & " + " & (UNIcdbl(lgObjRs("DILIG_MM"),0)/60) & "))"
                 lgStrSQL = lgStrSQL & " WHERE  "
                 lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
                 lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
                lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
             Elseif UNIcdbl(lgObjRs("DILIG_HH")) = UNIcdbl(arrColVal(6)) then
             Else
                 lgStrSQL = "UPDATE  HCA090T"
                 lgStrSQL = lgStrSQL & " SET " 
                 lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME -  ((" & UNIcdbl(lgObjRs("DILIG_HH"),0) & " + " & (UNIcdbl(lgObjRs("DILIG_MM"),0)/60) & ") - (" & UNIcdbl(arrColVal(6),0) & " + " & (UNIcdbl(arrColVal(7),0)/60) & "))"
                 lgStrSQL = lgStrSQL & " WHERE  "
                 lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
                 lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
                lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
            End if 
        Elseif arrColVal(8) = "3" Then                                     'DAY_TIME이 3(    *4해줌   )인 경우 시간에 대해............       
             If  UNIcdbl(lgObjRs("DILIG_CNT")) < UNIcdbl(arrColVal(5)) Then
                 lgStrSQL = "UPDATE  HCA090T"
                 lgStrSQL = lgStrSQL & " SET " 
                 lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME + " &( (UNIcdbl(arrColVal(5),0) - UNIcdbl(lgObjRs("DILIG_CNT"),0))*4)
                 lgStrSQL = lgStrSQL & " WHERE  "
                 lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
                 lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
                lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
             Elseif UNIcdbl(lgObjRs("DILIG_CNT")) = UNIcdbl(arrColVal(5)) then
             Else
                 lgStrSQL = "UPDATE  HCA090T"
                 lgStrSQL = lgStrSQL & " SET " 
                 lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME - " & ((UNIcdbl(lgObjRs("DILIG_CNT"),0) - UNIcdbl(arrColVal(5),0))*4)
                 lgStrSQL = lgStrSQL & " WHERE  "
                 lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
                 lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
                lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
            End if 
        End if 
   
   Else
   End if
   
   If  arrColVal(10) = "N" Then           '근태코드 테이블에서 wk_day = 'N'인것--- work day(근무일수) 계산 
         lgStrSQL = "SELECT  DILIG_CNT, DILIG_HH , DILIG_MM " 
         lgStrSQL = lgStrSQL & " FROM  HCA070T "
         lgStrSQL = lgStrSQL & " WHERE DILIG_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
         lgStrSQL = lgStrSQL & " AND EMP_NO     = " & FilterVar(UCase(arrColVal(3)), "''", "S")
         lgStrSQL = lgStrSQL & " AND DILIG_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        Call FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")

         If  UNIcdbl(lgObjRs("DILIG_CNT")) < UNIcdbl(arrColVal(5)) Then
            lgStrSQL = "UPDATE  HCA090T"
            lgStrSQL = lgStrSQL & " SET " 
            lgStrSQL = lgStrSQL & " WORK_DAY  = WORK_DAY  - (" & UNIcdbl(arrColVal(5),0) & "-" & UNIcdbl(lgObjRs("DILIG_CNT"),0) & ")"
            lgStrSQL = lgStrSQL & " WHERE  "
            lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
            lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
            
         Elseif UNIcdbl(lgObjRs("DILIG_CNT")) = UNIcdbl(arrColVal(5)) then
         Else
            lgStrSQL = "UPDATE  HCA090T"
            lgStrSQL = lgStrSQL & " SET " 
            lgStrSQL = lgStrSQL & " WORK_DAY  = WORK_DAY  + (" & UNIcdbl(lgObjRs("DILIG_CNT"),0) & "-" & UNIcdbl(arrColVal(5),0) & ")"
            lgStrSQL = lgStrSQL & " WHERE  "
            lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
            lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
            
        End if 
   Else
   End if
   
   If  arrColVal(11) = "N" Then           '출근일수 구분이 'N'이면 출근일수에서 minus
    '근무일수(ATTEND_DAY) = 총일수 - 휴일 - 일요일 - (근태테이블(HCA070T)에서 ATTEND_DAY가 "N"인 것은 그 횟수)
         lgStrSQL = "SELECT  DILIG_CNT, DILIG_HH , DILIG_MM " 
         lgStrSQL = lgStrSQL & " FROM  HCA070T "
         lgStrSQL = lgStrSQL & " WHERE DILIG_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
         lgStrSQL = lgStrSQL & " AND EMP_NO     = " & FilterVar(UCase(arrColVal(3)), "''", "S")
         lgStrSQL = lgStrSQL & " AND DILIG_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
         Call FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")
   
         If  UNIcdbl(lgObjRs("DILIG_CNT")) < UNIcdbl(arrColVal(5)) Then
            lgStrSQL = "UPDATE  HCA090T"
            lgStrSQL = lgStrSQL & " SET " 
            lgStrSQL = lgStrSQL & " ATTEND_DAY  = " & UNIcdbl(arrColVal(12),0) & " - (" & UNIcdbl(arrColVal(5),0) & "-" & UNIcdbl(lgObjRs("DILIG_CNT"),0) & ")"
            lgStrSQL = lgStrSQL & " WHERE  "
            lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
            lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
         Elseif UNIcdbl(lgObjRs("DILIG_CNT")) = UNIcdbl(arrColVal(5)) then
         Else
            lgStrSQL = "UPDATE  HCA090T"
            lgStrSQL = lgStrSQL & " SET " 
            lgStrSQL = lgStrSQL & " ATTEND_DAY  = " & UNIcdbl(arrColVal(12),0) & " + (" & UNIcdbl(lgObjRs("DILIG_CNT"),0) & "-" & UNIcdbl(arrColVal(5),0) & ")"
            lgStrSQL = lgStrSQL & " WHERE  "
            lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
            lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
        End if 
   Else
   End if
   
    lgStrSQL = "UPDATE  HCA070T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " DILIG_CNT   = " & FilterVar(Trim(UCase(arrColVal(5))),"0","D")   & ","
    lgStrSQL = lgStrSQL & " DILIG_HH    = " & FilterVar(Trim(UCase(arrColVal(6))),"0","D")   & ","
    lgStrSQL = lgStrSQL & " DILIG_MM    = " & FilterVar(Trim(UCase(arrColVal(7))),"0","D")  
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " DILIG_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO     = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DILIG_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub


'============================================================================================================
' Name : MultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

   If  arrColVal(9) = "Y" Then            '☜  삭제시 차감일수와 차감시간을 빼준다.......
        If  arrColVal(8) = "1" Then 
             lgStrSQL = "UPDATE  HCA090T"
             lgStrSQL = lgStrSQL & " SET " 
             lgStrSQL = lgStrSQL & " MARGIR_DAY  = MARGIR_DAY - " & UNIcdbl(arrColVal(5),0) & ","
             lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME - " & (UNIcdbl(arrColVal(5),0) * 8)
             lgStrSQL = lgStrSQL & " WHERE  "
             lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
             lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
        Elseif arrColVal(8) = "2" Then 
             lgStrSQL = "UPDATE  HCA090T"
             lgStrSQL = lgStrSQL & " SET " 
             lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME - (" & UNIcdbl(arrColVal(6),0) & " + " & (UNIcdbl(arrColVal(7),0)/60) & ")"
             lgStrSQL = lgStrSQL & " WHERE  "
             lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
             lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
        Elseif arrColVal(8) = "3" Then   
             lgStrSQL = "UPDATE  HCA090T"
             lgStrSQL = lgStrSQL & " SET " 
             lgStrSQL = lgStrSQL & " MARGIR_TIME = MARGIR_TIME - " & (UNIcdbl(arrColVal(5),0) * 4)
             lgStrSQL = lgStrSQL & " WHERE  "
             lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
             lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
            lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
        End if 
   Else
   End if
    
   If  arrColVal(10) = "N" Then           '근태코드 테이블에서 wk_day = 'N'인것 
        lgStrSQL = "UPDATE  HCA090T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " WORK_DAY  = (WORK_DAY  + " & UNIcdbl(arrColVal(5),0) & ")"
        lgStrSQL = lgStrSQL & " WHERE  "
        lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
        lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
   Else
   End if

   If  arrColVal(11) = "N" Then           '출근일수 구분이 'N'이면 출근일수에서 minus
    '근무일수(ATTEND_DAY) = 총일수 - 휴일 - 일요일 - (근태테이블(HCA070T)에서 ATTEND_DAY가 "N"인 것은 그 횟수)
        lgStrSQL = "UPDATE  HCA090T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " ATTEND_DAY  =  " & UNIcdbl(arrColVal(12),0) + UNIcdbl(arrColVal(5),0)
        lgStrSQL = lgStrSQL & " WHERE  "
        lgStrSQL = lgStrSQL & " WK_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
        lgStrSQL = lgStrSQL & " AND EMP_NO  = " & FilterVar(UCase(arrColVal(3)), "''", "S")
        lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords  
   Else
   End if
   
    lgStrSQL = "DELETE  HCA070T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " DILIG_YYMM     = " & FilterVar(arrColVal(2), "''", "S")
    lgStrSQL = lgStrSQL & " AND EMP_NO     = " & FilterVar(UCase(arrColVal(3)), "''", "S")
    lgStrSQL = lgStrSQL & " AND DILIG_CD   = " & FilterVar(UCase(arrColVal(4)), "''", "S")
    
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
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL = " SELECT WK_YYMM, EMP_NO, TOT_DAY, SUN_DAY, HOL_DAY, WEEK_DAY, "
                                   lgStrSQL = lgStrSQL & " NON_WEEK_DAY, MARGIR_DAY, MARGIR_TIME, WK_DAY,WORK_DAY, ATTEND_DAY, "
                                   lgStrSQL = lgStrSQL & " WORK_DAY, WK_TIME " 
                                   lgStrSQL = lgStrSQL & " FROM HCA090T "
                                   lgStrSQL = lgStrSQL & " WHERE ( EMP_NO " & pComp & " " &  pCode                           
 							 Case "P"
 							       lgStrSQL = " SELECT TOP 1 WK_YYMM, EMP_NO, TOT_DAY, SUN_DAY, HOL_DAY, WEEK_DAY, "
                                   lgStrSQL = lgStrSQL & " NON_WEEK_DAY, MARGIR_DAY, MARGIR_TIME, WK_DAY,WORK_DAY, ATTEND_DAY, "
                                   lgStrSQL = lgStrSQL & " WORK_DAY, WK_TIME " 
                                   lgStrSQL = lgStrSQL & " FROM HCA090T "
                                   lgStrSQL = lgStrSQL & " WHERE ( EMP_NO < " & pCode  
                                   lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                             Case "N"
								   lgStrSQL = " SELECT TOP 1 WK_YYMM, EMP_NO, TOT_DAY, SUN_DAY, HOL_DAY, WEEK_DAY, "
                                   lgStrSQL = lgStrSQL & " NON_WEEK_DAY, MARGIR_DAY, MARGIR_TIME, WK_DAY,WORK_DAY, ATTEND_DAY, "
                                   lgStrSQL = lgStrSQL & " WORK_DAY, WK_TIME " 
                                   lgStrSQL = lgStrSQL & " FROM HCA090T "
                                   lgStrSQL = lgStrSQL & " WHERE ( EMP_NO > " & pCode  
                                   lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
                        End Select
           End Select             
        Case "M"           '                  0               1                 2              3                4                5
			if lgSpreadFlg = "1" then
				iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
			elseif lgSpreadFlg = "2" then            		
				iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1
			end if

           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "SELECT TOP " & iSelCount 
                       lgStrSQL = lgStrSQL & " b.DILIG_YYMM, b.EMP_NO, a.DILIG_CD, a.DILIG_NM, b.DILIG_CNT, "
                       lgStrSQL = lgStrSQL & " b.DILIG_HH, b.DILIG_MM, a.DAY_TIME, a.ATTEND_DAY, a.BAS_MARGIR, a.WK_DAY"
                       lgStrSQL = lgStrSQL & "  FROM HCA010T a, HCA070T b "
                       lgStrSQL = lgStrSQL & " WHERE a.DILIG_CD = b.DILIG_CD "
                       lgStrSQL = lgStrSQL & "   AND ((b.EMP_NO " & pComp & " " &  pCode 	
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
			With Parent

				if .gSpreadFlg = "1" then
					.ggoSpread.Source     = .frm1.vspdData
					.ggoSpread.SSShowData "<%=lgstrData1%>"                               '☜ : Display data
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					if .topleftOK then
						.DBQueryOk
					else
						.gSpreadFlg = "2"						
						.DBQuery
					end if
                elseif .gSpreadFlg = "2" then
					.ggoSpread.Source     = .frm1.vspdData1
					.ggoSpread.SSShowData "<%=lgstrData2%>"          
					.lgStrPrevKey1    = "<%=lgStrPrevKey1%>"
	                .DBQueryOk
                end if
                
	         End with
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>
