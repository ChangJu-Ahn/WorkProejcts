<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear 
    
	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")                                                                       '☜: Clear Error status
    
    Call HideStatusWnd                                                               '☜: Hide Processing message


	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim txtYYYYMM		 
	Dim txtPlantCd,txtPlantNm		 		 
	Dim txtItemAccntCd,txtItemAccntNm
	Dim txtItemCd,txtItemNm
	Dim txtFlag
	Dim lgStrPrevKey
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6	

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)


	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection



'
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
Dim iDx
Dim iLoopMax
Dim pKey1
Dim Currency_code


Dim intRetCd

On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

'---------- Developer Coding part (Start) ---------------------------------------------------------------

	txtYYYYMM		 = Trim(Request("txtYYYYMM"))
	txtPlantCd		 = Trim(Request("txtPlantCd"))
	txtItemAccntCd	 = Trim(Request("txtItemAccntCd"))
	txtItemCd		 = Trim(Request("txtItemCd"))
	lgStrPrevKey	 = Trim(Request("lgStrPrevKey"))         '☜: Next Key Value
	txtFlag			 = Trim(Request("txtFlag"))
'	lgMaxCount       = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData

   	Const C_SHEETMAXROWS_D  = 100 
    
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time

	intRetCd = CommonQueryRs("plant_nm","b_plant","plant_cd = " & FilterVar(txtPlantCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Trim(Replace(lgF0,Chr(11),"")) = "X"  then
		txtPlantNm = ""
		Call SetErrorStatus
		Call DisplayMsgBox("125010", vbInformation, "", "", I_MKSCRIPT)	
		Exit Sub
	else
		txtPlantNm = Trim(Replace(lgF0,Chr(11),""))	  
	end if


	intRetCd = CommonQueryRs("minor_nm","b_minor","major_cd = " & FilterVar("P1001", "''", "S") & "  and minor_cd = " & FilterVar(txtItemAccntCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	if Trim(Replace(lgF0,Chr(11),"")) = "X"  then
		txtItemAccntNm = ""
		Call SetErrorStatus
		Call DisplayMsgBox("169952", vbInformation, "", "", I_MKSCRIPT)	                      '☜ : 품목계정이 Minor 코드에 존재하지 않습니다 
		Exit Sub
	else
		txtItemAccntNm = Trim(Replace(lgF0,Chr(11),""))	  
	end if

	If txtItemCd <> "" Then
		intRetCd = CommonQueryRs("item_nm","b_item","item_cd = " & FilterVar(txtItemCd, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		if Trim(Replace(lgF0,Chr(11),"")) = "X" then
			txtItemNm = ""
		else
			txtItemNm = Trim(Replace(lgF0,Chr(11),""))	  
		end if
	End If

	If txtFlag = "RCPT" Then													'☜ : 입고단가 
	   Call SubMakeSQLStatements("QR","X")                                   
	Else																	'☜ : 총평균단가 
	   Call SubMakeSQLStatements("QT","X")                                   
	End If
	

    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found.
        Call SetErrorStatus
    Else
        
        lgstrData = ""

        iDx = 1

        Do While Not lgObjRs.EOF
			IF iDx < lgMaxCount + 1 Then
                lgstrData = lgstrData & Chr(11) & ""
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))		'품목코드 
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))		'품목명 
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))		'Tracking No                
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))		'Basic Unit
                lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))		'Spec
                
                IF txtFlag = "RCPT" Then
					lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(5),ggUnitCost.DecPoint,0)		'입고단가 
					lgstrData = lgstrData & Chr(11) & UNINumClientFormat(0,ggUnitCost.DecPoint,0)
				ELSE
					lgstrData = lgstrData & Chr(11) & UNINumClientFormat(0,ggUnitCost.DecPoint,0)							
					lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(5),ggUnitCost.DecPoint,0)		'총평균단가		
				END IF                
                
                lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs(6),ggUnitCost.DecPoint,0)		'재고표준단가 
                lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx+1
                lgstrData = lgstrData & Chr(11) & Chr(12)
             
				lgObjRs.MoveNext
				iDx = iDx + 1
			ELSE	
				lgStrPrevKey = Trim(ConvSPChars(lgObjRs(0))) + Trim(ConvSPChars(lgObjRs(2)))
				Exit Do
            End If
	    Loop
    End If

    If iDx < (lgMaxCount+1) Then
       lgStrPrevKey = ""
    End If
'    Call SubHandleError("Q",lgObjConn, lgObjRs,Err)

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
	Dim pCode
	Dim iWhere
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	txtYYYYMM		 = Trim(Request("hYYYYMM"))
	txtPlantCd		 = Trim(Request("hPlantCd"))
	txtItemAccntCd	 = Trim(Request("hItemAccntCd"))
	txtItemCd		 = Trim(Request("hItemCd"))

	iWhere = Trim(Request("hChecked"))
	txtFlag	= Trim(Request("hRadio"))

	IF iWhere = "S" Then		'☜: 선택반영 
		arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
		pCode = "("

		For iDx = 0 To UBOUND(arrRowVal,1) - 1
		    arrColVal = Split(arrRowVal(iDx), gColSep)                                 '☜: Split Column data

		    pCode = pCode & FilterVar(Trim(arrColVal(0)) & Trim(arrColVal(1)), "''", "S") & ","
		Next

		pCode = MID(pCode,1,len(pCode) -1) & ")"
	ELSE						'☜: 일괄반영 
		pcode = ""
	END IF
	
		
	IF txtflag = "RCPT"	Then		'☜: 입고단가 
		Call SubMakeSQLStatements("UR",pCode)
	ELSE							'☜: 총평균단가 
		Call SubMakeSQLStatements("UT",pCode)
	END IF
	    
    
    
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(iWhere,pCode)
    

    '------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case iWhere
        Case "QR"
            lgStrSQL =			  " select  top  " & CStr(lgMaxCount + 1) 
            lgStrSQL = lgStrSQL & "			a.item_cd,c.item_nm,a.tracking_no,c.basic_unit,c.spec,a.price,d.std_prc as std_prc "
            lgStrSQL = lgStrSQL & " from ("
            lgStrSQL = lgStrSQL & "			select plant_cd,item_cd,tracking_no,case when sum(rcpt_qty) <> 0 then sum(actl_rcpt_amt) / sum(rcpt_qty) else 0 end as price "
            lgStrSQL = lgStrSQL & "			from	c_bom_rcpt_by_opr_s(nolock) "
            lgStrSQL = lgStrSQL & "			where	yyyymm = " & FilterVar(txtYYYYMM, "''", "S") 				
            lgStrSQL = lgStrSQL & "			group by plant_cd,item_cd,tracking_no "		            
            lgStrSQL = lgStrSQL & "			having case when sum(rcpt_qty) <> 0 then sum(actl_rcpt_amt) / sum(rcpt_qty) else 0 end <> 0 ) a"		                        
            lgStrSQL = lgStrSQL & "		,B_ITEM_BY_PLANT b(nolock), B_ITEM c(nolock), I_MATERIAL_VALUATION d(nolock) "
            lgStrSQL = lgStrSQL & " where	a.plant_cd = " & FilterVar(txtPlantCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		b.item_acct = " & FilterVar(txtItemAccntCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		LTRIM(RTRIM(a.item_cd)) + LTRIM(RTRIM(a.tracking_no)) >= " & FilterVar(lgStrPrevKey, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.item_cd >= " & FilterVar(txtItemCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.plant_cd = b.plant_cd and a.item_cd = b.item_cd "
            lgStrSQL = lgStrSQL & " and		a.item_cd = c.item_cd "
            lgStrSQL = lgStrSQL & " and		a.plant_cd = d.plant_cd and a.item_cd = d.item_cd and a.tracking_no = d.tracking_no"
            lgStrSQL = lgStrSQL & " and		a.price > 0 "
            'lgStrSQL = lgStrSQL & " group by a.item_cd,c.item_nm,c.basic_unit,c.spec,a.price "            
            lgStrSQL = lgStrSQL & " order by a.item_cd,a.tracking_no"
            
            
        Case "QT"
            lgStrSQL =			  " select  top  " & CStr(lgMaxCount + 1) 
            lgStrSQL = lgStrSQL & "			a.item_cd,c.item_nm,a.tracking_no,c.basic_unit,c.spec,a.avg_prc,d.std_prc "
            lgStrSQL = lgStrSQL & " from	C_ITEM_BY_PLANT_S a(nolock), B_ITEM_BY_PLANT b(nolock), B_ITEM c(nolock), I_MATERIAL_VALUATION d(nolock) "
            lgStrSQL = lgStrSQL & " where	a.yyyymm = " & FilterVar(txtYYYYMM, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.plant_cd = " & FilterVar(txtPlantCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		b.item_acct = " & FilterVar(txtItemAccntCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		LTRIM(RTRIM(a.item_cd)) + LTRIM(RTRIM(a.tracking_no)) >= " & FilterVar(lgStrPrevKey, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.item_cd >= " & FilterVar(txtItemCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.plant_cd = b.plant_cd and a.item_cd = b.item_cd "
            lgStrSQL = lgStrSQL & " and		a.item_cd = c.item_cd "
            lgStrSQL = lgStrSQL & " and		a.plant_cd = d.plant_cd and a.item_cd = d.item_cd and a.tracking_no = d.tracking_no "
            lgStrSQL = lgStrSQL & " and		a.avg_prc > 0 "
            'lgStrSQL = lgStrSQL & " group by a.item_cd,c.item_nm,c.basic_unit,c.spec,avg_prc "
            lgStrSQL = lgStrSQL & " order by a.item_cd"
        
        Case "UR"
            lgStrSQL = lgStrSQL & "	update	c"
            lgStrSQL = lgStrSQL & "	set		c.std_prc = a.price ,updt_user_id =  " & FilterVar(gUsrID , "''", "S") & ",updt_dt = getdate() "
			lgStrSQL = lgStrSQL & " from	I_MATERIAL_VALUATION c"
            lgStrSQL = lgStrSQL & " join ("
            lgStrSQL = lgStrSQL & "			select plant_cd,item_cd,tracking_no,case when sum(rcpt_qty) <> 0 then sum(actl_rcpt_amt) / sum(rcpt_qty) else 0 end as price "
            lgStrSQL = lgStrSQL & "			from	c_bom_rcpt_by_opr_s "
            lgStrSQL = lgStrSQL & "			where	yyyymm = " & FilterVar(txtYYYYMM, "''", "S") 				
            lgStrSQL = lgStrSQL & "			group by plant_cd,item_cd,tracking_no "		            
            lgStrSQL = lgStrSQL & "			having case when sum(rcpt_qty) <> 0 then sum(actl_rcpt_amt) / sum(rcpt_qty) else 0 end <> 0 ) a"					
            lgStrSQL = lgStrSQL & " on a.plant_cd = c.plant_cd and a.item_cd = c.item_cd and a.tracking_no = c.tracking_no "
            lgStrSQL = lgStrSQL & " inner join  B_ITEM_BY_PLANT b on a.plant_cd = b.plant_cd and a.item_cd = b.item_cd"
            lgStrSQL = lgStrSQL & " where	a.plant_cd = " & FilterVar(txtPlantCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		b.item_acct = " & FilterVar(txtItemAccntCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.item_cd >= " & FilterVar(txtItemCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.price > 0 "

            
            IF pCode <> "" Then
				lgStrSQL = lgStrSQL & " and		LTRIM(RTRIM(a.item_cd)) + LTRIM(RTRIM(a.tracking_no)) in " & pCode
			END IF
        
        Case "UT"
            lgStrSQL = lgStrSQL & "	update	c"
            lgStrSQL = lgStrSQL & "	set		c.std_prc = a.avg_prc,"
			lgStrSQL = lgStrSQL & "			updt_user_id =  " & FilterVar(gUsrID , "''", "S") & ",updt_dt = getdate() "
            lgStrSQL = lgStrSQL & " from	I_MATERIAL_VALUATION c "
            lgStrSQL = lgStrSQL & " inner join	C_ITEM_BY_PLANT_S a on a.plant_cd = c.plant_cd and a.item_cd = c.item_cd and a.tracking_no = c.tracking_no "
            lgStrSQL = lgStrSQL & " inner join  B_ITEM_BY_PLANT b on a.plant_cd = b.plant_cd and a.item_cd = b.item_cd "
            lgStrSQL = lgStrSQL & " where	a.yyyymm = " & FilterVar(txtYYYYMM, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.plant_cd = " & FilterVar(txtPlantCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		b.item_acct = " & FilterVar(txtItemAccntCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.item_cd >= " & FilterVar(txtItemCd, "''", "S") 
            lgStrSQL = lgStrSQL & " and		a.avg_prc > 0 "
            
            IF pCode <> "" Then
				lgStrSQL = lgStrSQL & " and		LTRIM(RTRIM(a.item_cd)) + LTRIM(RTRIM(a.tracking_no)) in " & pCode
			END IF
          
    End Select

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
        Case "Q"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
          With Parent
			.frm1.txtPlantNM.value		="<%=ConvSPChars(txtPlantnm)%>" 
			.frm1.txtItemAccntNm.value	="<%=ConvSPChars(txtItemAccntNm)%>" 
			.frm1.txtItemNm.value		="<%=ConvSPChars(txtItemNm)%>" 
			
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				.frm1.hYYYYMM.Value			= "<%=txtYYYYMM%>"
                .frm1.hPlantCd.Value		= "<%=txtPlantCd%>"
                .frm1.hItemAccntCd.Value	= "<%=txtItemAccntCd%>"
                .frm1.hItemCd.Value			= "<%=txtItemCd%>"
                .frm1.hRadio.Value			= "<%=txtFlag%>"
                
                .ggoSpread.Source			= .frm1.vspdData
                .ggoSpread.SSShowData		"<%=lgstrData%>"                             '☜ : Display data
                .lgStrPrevKey				="<%=lgStrPrevKey%>"                          '☜ : Next next data tag
                .DBQueryOk
			End If
	      End with
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If
    End Select
</Script>
