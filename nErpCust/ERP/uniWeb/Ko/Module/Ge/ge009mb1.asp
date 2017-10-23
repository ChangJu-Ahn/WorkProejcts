<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    
    Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")   
                                                                            '☜: Clear Error status
    Dim lgtotSalesAmt
    Dim lgtotCostAmt     '매출원가 총계 
    Dim lgtotPorfitAmt   '매출이익 총계 
    Dim lgtotTotCostAmt  '총원가 총계 
    Dim lgtotCurProfitAmt   '경상이익 총계 
    Dim lgtotNetProfitAmt   '순이익 총계 
    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Filtervar(Request("txtKeyStream"),"","SNM"),gColSep)

    'Multi SpreadSheet

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgMaxCount        = UniConvNumStringToDouble(Request("lgMaxCount"),0)                                  '☜: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UniConvNumStringToDouble(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
   	Const C_SHEETMAXROWS_D  = 100  
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time	                 

    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
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
Dim RetKeyStream,RetTxtNm
Dim txtBizUnitnm,txtCostnm,txtSalesOrgnm,txtSalesGrpnm

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	If Trim(lgKeyStream(2)) <> "" Then		'사업부 
		Call CheckMACondition(" BIZ_UNIT_NM "," B_BIZ_UNIT "," BIZ_UNIT_CD =  " & FilterVar(lgKeyStream(2), "''", "S") & "" ,Trim(lgKeyStream(2)),RetKeyStream,RetTxtNm)
		txtBizUnitnm = RetTxtNm
		lgKeyStream(2) = RetKeyStream
	End If
	Response.Write Err.Description 
	If Trim(lgKeyStream(3)) <> "" Then	'Profit Center		
		Call CheckMACondition(" COST_NM "," B_COST_CENTER "," COST_TYPE = " & FilterVar("S", "''", "S") & "  AND COST_CD=  " & FilterVar(lgKeyStream(3), "''", "S") & "" ,Trim(lgKeyStream(3)),RetKeyStream,RetTxtNm)		
		txtCostnm = RetTxtNm		
		lgKeyStream(3) = RetKeyStream
	End If
	
	If Trim(lgKeyStream(4)) <> "" Then	'영업조직		
		Call CheckMACondition(" SALES_ORG_NM "," B_SALES_ORG "," SALES_ORG  =  " & FilterVar(lgKeyStream(4), "''", "S") & "" ,Trim(lgKeyStream(4)),RetKeyStream,RetTxtNm)
		txtSalesOrgnm = RetTxtNm		
		lgKeyStream(4) = RetKeyStream
	End If
	
	If Trim(lgKeyStream(5)) <> "" Then	'영업그룹 
		Call CheckMACondition(" SALES_GRP_NM "," B_SALES_GRP "," SALES_GRP  =  " & FilterVar(lgKeyStream(5), "''", "S") & "" ,Trim(lgKeyStream(5)),RetKeyStream,RetTxtNm)
		txtSalesGrpnm = RetTxtNm		
		lgKeyStream(5) = RetKeyStream
	End If
%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
      With Parent.Frm1
             .txtBizUnitNm.Value				= "<%=txtBizUnitNm%>"            
             .txtCostNm.Value					= "<%=txtCostNm%>"            
             .txtSalesOrgnm.Value				= "<%=txtSalesOrgnm%>"
             .txtSalesGrpnm.Value				= "<%=txtSalesGrpnm%>"
      End With          
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
</Script>       
<%        
    Call SubBizQueryMulti()
    
    If lgErrorStatus = "NO" Then
		Call SubBizQuerySingle()
	End If
    
End Sub 
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub CheckMACondition(StrField,StrTable,StrCon,StrKeyStream,RetKeyStream,RetTxtNm)
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	RetKeyStream = ""
	RetTxtNm = ""
	
	Call CommonQueryRs(StrField,StrTable,StrCon ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	If Trim(Replace(lgF0,Chr(11),"")) = "X" then
	  RetTxtNm = ""
	  RetKeyStream = ""
	Else   
	  RetTxtNm = Trim(Replace(lgF0,Chr(11),""))
	  RetKeyStream = Trim(StrKeyStream)
	End if    	    
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
    Dim iDx

    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    
    Call SubMakeSQLStatements("MR","X")                                                  '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        lgStrPrevKeyIndex = ""    
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)                  '☜: No data is found. 
        Call SetErrorStatus()
    Else
    
       If CDbl(lgStrPrevKeyIndex) > 0 Then
          lgObjRs.Move     = CDbl(lgMaxCount) * CDbl(lgStrPrevKeyIndex)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       End If
         
       lgstrData = ""
        
       iDx = 1
       
       Do While Not lgObjRs.EOF
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
          lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(2),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(3),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(4),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(5),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(6),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(7),ggAmtOfMoney.DecPoint, 0)
          lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
          lgstrData = lgstrData & Chr(11) & Chr(12)

          lgObjRs.MoveNext

          iDx =  iDx + 1
          If iDx > lgMaxCount Then
             lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
             Exit Do
         End If   
      Loop 
    End If
    
    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)	
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
	
End Sub    
'============================================================================================================
' Name : SubBizQuerySingle
' Desc : Save Data 
'============================================================================================================
Sub SubBizQuerySingle()
    Dim iDx

    On Error Resume Next                                                                 '☜: Protect system from crashing
    Err.Clear                                                                            '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------

    Call SubMakeSQLStatements("S", "X")                                                  '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        lgStrPrevKeyIndex = ""        
        Call SetErrorStatus()
		lgtotSalesAmt=0		'매출액 총계 
		lgtotCostAmt=0     '매출원가 총계 
		lgtotPorfitAmt=0   '매출이익 총계 
		lgtotTotCostAmt=0  '총원가 총계 
		lgtotCurProfitAmt =0  '경상이익 총계 
		lgtotNetProfitAmt=0   '순이익 총계 
    Else
    	lgtotSalesAmt=lgObjRs(0)		'매출액 총계 
		lgtotCostAmt=lgObjRs(1)     '매출원가 총계 
		lgtotPorfitAmt=lgObjRs(2)   '매출이익 총계 
		lgtotTotCostAmt=lgObjRs(3)  '총원가 총계 
		lgtotCurProfitAmt =lgObjRs(4)  '경상이익 총계 
		lgtotNetProfitAmt=lgObjRs(5)   '순이익 총계 
    End If
    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
    
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
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    Call SubMakeSQLStatements("MC",arrColVal)                                        '☆: Make sql statements
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
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
    Call SubMakeSQLStatements("MU",arrColVal)                                        '☆: Make sql statements
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
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
    Call SubMakeSQLStatements("MD",arrColVal)                                        '☆: Make sql statements

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case Mid(pDataType,1,1)
        Case "S"
            lgStrSQL = "Select "
            lgStrSQL = lgStrSQL & " IsNull(SUM(CASE WHEN A.GAIN_CD = " & FilterVar("B01", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0), "			'매출액 
            lgStrSQL = lgStrSQL & " IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0), "			'매출원가 
            lgStrSQL = lgStrSQL & " (IsNull(SUM(CASE WHEN A.GAIN_CD = " & FilterVar("B01", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
            lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0)), "		'매출이익 
            lgStrSQL = lgStrSQL & " (IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("E%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("F%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
            lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("H%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("I%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0)), "			'총원가 
            lgStrSQL = lgStrSQL & " (IsNull(SUM(CASE WHEN A.GAIN_CD = " & FilterVar("B01", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
            lgStrSQL = lgStrSQL & " - (IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("E%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("F%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
            lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("H%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("I%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0))), "			'경상이익 
            lgStrSQL = lgStrSQL & " ((IsNull(SUM(CASE WHEN A.GAIN_CD = " & FilterVar("B01", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
            lgStrSQL = lgStrSQL & " - (IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("E%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("F%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "                        
            lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("H%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "                        
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("I%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0))) "			                       
            lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("L%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "			
            lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("M%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0)) "	'순이익 
            lgStrSQL = lgStrSQL & " From G_ITEM_PROFIT A ,B_ITEM_GROUP B " 
            lgStrSQL = lgStrSQL & " WHERE A.ITEM_GROUP_CD *= B.ITEM_GROUP_CD "
            lgStrSQL = lgStrSQL & " AND A.YYYYMM BETWEEN " & FilterVar(lgKeyStream(0), "''", "S")
            lgStrSQL = lgStrSQL & " AND " & FilterVar(lgKeyStream(1), "''", "S")
            If Trim(lgKeyStream(2)) <> "" Then
			 lgStrSQL = lgStrSQL & " AND A.BIZ_UNIT_CD = " & FilterVar(lgKeyStream(2), "''", "S")
			End If
			If Trim(lgKeyStream(3)) <> "" Then
			 lgStrSQL = lgStrSQL & " AND A.COST_CD = " & FilterVar(lgKeyStream(3), "''", "S")
			End If
            If Trim(lgKeyStream(4)) <> "" Then
			 lgStrSQL = lgStrSQL & " AND A.SALES_ORG = " & FilterVar(lgKeyStream(4), "''", "S")
			End If
            If Trim(lgKeyStream(5)) <> "" Then
			 lgStrSQL = lgStrSQL & " AND A.SALES_GRP = " & FilterVar(lgKeyStream(5), "''", "S")
			End If                       
 
        Case "M"
           iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
           Select Case Mid(pDataType,2,1)
               Case "R"
                       lgStrSQL = "Select TOP " & iSelCount & " ISNULL(A.ITEM_GROUP_CD,''),ISNULL(B.ITEM_GROUP_NM,''),"                        
                       lgStrSQL = lgStrSQL & " IsNull(SUM(CASE WHEN A.GAIN_CD = " & FilterVar("B01", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0), "				'매출액 
                       lgStrSQL = lgStrSQL & " IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0), "			'매출원가 
                       lgStrSQL = lgStrSQL & " (IsNull(SUM(CASE WHEN A.GAIN_CD = " & FilterVar("B01", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
                       lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0)), "		'매출이익 
                       lgStrSQL = lgStrSQL & " (IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("E%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("F%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
                       lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("H%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("I%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0)), "			'총원가 
                       lgStrSQL = lgStrSQL & " (IsNull(SUM(CASE WHEN A.GAIN_CD = " & FilterVar("B01", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
                       lgStrSQL = lgStrSQL & " - (IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("E%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("F%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
                       lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("H%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("I%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0))), "			'경상이익 
                       lgStrSQL = lgStrSQL & " ((IsNull(SUM(CASE WHEN A.GAIN_CD = " & FilterVar("B01", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
                       lgStrSQL = lgStrSQL & " - (IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("C%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("E%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) " 
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("F%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "                        
                       lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("H%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "                        
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("I%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0))) "			                       
                       lgStrSQL = lgStrSQL & " + IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("L%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0) "			
                       lgStrSQL = lgStrSQL & " - IsNull(SUM(CASE WHEN A.GAIN_CD LIKE " & FilterVar("M%", "''", "S") & " THEN IsNull(A.AMOUNT,0) ELSE 0 END),0)) "			'순이익 
                       lgStrSQL = lgStrSQL & " From G_ITEM_PROFIT A ,B_ITEM_GROUP B " 
                       lgStrSQL = lgStrSQL & " WHERE A.ITEM_GROUP_CD *= B.ITEM_GROUP_CD "
                       lgStrSQL = lgStrSQL & " AND A.YYYYMM BETWEEN " & FilterVar(lgKeyStream(0), "''", "S")
                       lgStrSQL = lgStrSQL & " AND " & FilterVar(lgKeyStream(1), "''", "S")
                       If Trim(lgKeyStream(2)) <> "" Then
						lgStrSQL = lgStrSQL & " AND A.BIZ_UNIT_CD = " & FilterVar(lgKeyStream(2), "''", "S")
					   End If
					   If Trim(lgKeyStream(3)) <> "" Then
						lgStrSQL = lgStrSQL & " AND A.COST_CD = " & FilterVar(lgKeyStream(3), "''", "S")
					   End If
                       If Trim(lgKeyStream(4)) <> "" Then
						lgStrSQL = lgStrSQL & " AND A.SALES_ORG = " & FilterVar(lgKeyStream(4), "''", "S")
					   End If
                       If Trim(lgKeyStream(5)) <> "" Then
						lgStrSQL = lgStrSQL & " AND A.SALES_GRP = " & FilterVar(lgKeyStream(5), "''", "S")
					   End If
                       lgStrSQL = lgStrSQL & " GROUP BY A.ITEM_GROUP_CD,B.ITEM_GROUP_NM "
                       lgStrSQL = lgStrSQL & " ORDER BY A.ITEM_GROUP_CD"
                       
           End Select             
           
    End Select
    
    'Response.Write lgStrSQL
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
       With Parent		
                .Frm1.totSalesAmt.text     = "<%=UNINumClientFormat(lgtotSalesAmt,ggQty.DecPoint, 0)%>"      '매출액 총계 
                .Frm1.totCostAmt.text		= "<%=UNINumClientFormat(lgtotCostAmt,ggQty.DecPoint, 0)%>"      '매출원가 총계 
                .Frm1.totPorfitAmt.text    = "<%=UNINumClientFormat(lgtotPorfitAmt,ggQty.DecPoint, 0)%>"      '매출이익 총계 
                .Frm1.totTotCostAmt.text	= "<%=UNINumClientFormat(lgtotTotCostAmt,ggQty.DecPoint, 0)%>"      '총원가 총계 
                .Frm1.totCurProfitAmt.text = "<%=UNINumClientFormat(lgtotCurProfitAmt,ggQty.DecPoint, 0)%>"      '경상이익 총계 
                .Frm1.totNetProfitAmt.text = "<%=UNINumClientFormat(lgtotNetProfitAmt,ggQty.DecPoint, 0)%>"      '순이익 총계 
       End With   
 
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source         = .frm1.vspdData
                .lgStrPrevKeyIndex        = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk()       
                If <%=UNICDbl(lgKeyStream(6), 0)%> = "1" Or <%=UNICDbl(lgKeyStream(6), 0)%> = "2" Then					
					.ggoSpread.SSSort <%=UNICDbl(lgKeyStream(6), 0)%>, 1							
				Else
					.ggoSpread.SSSort <%=UNICDbl(lgKeyStream(6), 0)%>, 2							
				End If
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk()
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
