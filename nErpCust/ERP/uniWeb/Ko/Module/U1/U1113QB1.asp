<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Dim lgStrColorFlag
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
	Call LoadInfTB19029B("Q", "*","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("Q", "*","NOCOOKIE", "MB")
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)					 '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
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
    Call SubBizQueryCond()
    If lgErrorStatus <> "YES" Then
		Call SubBizQueryMulti()
	End If
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
' Name : SubBizQueryCond
' Desc : Verify Condition Value from Db
'============================================================================================================
Sub SubBizQueryCond()

    On Error Resume Next
    Err.Clear

	If lgKeyStream(0) <> "" Then
   
		Call SubMakeSQLStatements("CP")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(lgObjRs("PLANT_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(1) <> "" THEN 'AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CI")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(lgObjRs("ITEM_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(2) <> "" THEN 'AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CB")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("SCM003", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBpNm.value = """ & ConvSPChars(lgObjRs("BP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtBpNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(6) <> ""  THEN 'AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CS")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("SCM016", vbInformation, "", "", I_MKSCRIPT)
		    Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """ & ConvSPChars(lgObjRs("SL_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtSlNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If
	
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

     strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR")                                 '☆ : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
			If ConvSPChars(lgObjRs("ITEM_CD")) <> "zzzzzzzzzzzzzzzzzz" Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
			Else
				lgstrData = lgstrData & Chr(11) & ""
			End IF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))	
            lgstrData = lgstrData & Chr(11) & TRIM(ConvSPChars(lgObjRs("SPEC")))
            
            If lgObjRs("grp_flg") = 0 Then 
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("tracking_no"))
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "1" & gRowSep
			ElseIf lgObjRs("grp_flg") = 4 Then 
				lgstrData = lgstrData & Chr(11) & "[총계]"
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "5" & gRowSep
			ElseIf lgObjRs("grp_flg") = 3 Then 
				lgstrData = lgstrData & Chr(11) & "[품목소계]"
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "3" & gRowSep
			ElseIf lgObjRs("grp_flg") = 2 Then 
				lgstrData = lgstrData & Chr(11) & "[공급처소계]"
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "2" & gRowSep
			End If
				
			If ConvSPChars(lgObjRs("BP_CD")) <> "zzzzzzzzzz" Then
				lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
			Else
				lgstrData = lgstrData & Chr(11) & ""
			End IF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PLAN_DVRY_QTY"),ggQty.DecPoint,0)
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
        CALL SVRMSGBOX(arrColVal ,0,1)
        
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
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
                              
					lgStrSQL =			  " SELECT TOP " & iSelCount  & "	A.grp_flg , A.BP_CD , C.BP_NM , A.ITEM_CD , B.ITEM_NM , B.SPEC , a.tracking_no , A.PLAN_DVRY_DT , PLAN_DVRY_QTY  "
					lgStrSQL = lgStrSQL & " FROM   "

					lgStrSQL = lgStrSQL & " 	(SELECT	 (grouping(A.BP_CD) + grouping(A.ITEM_CD) + grouping(A.TRACKING_NO) + grouping(A.PLAN_DVRY_DT)) grp_flg , "
					lgStrSQL = lgStrSQL & " 		A.BP_CD, A.ITEM_CD, a.tracking_no , A.PLAN_DVRY_DT, SUM(A.PLAN_DVRY_QTY) PLAN_DVRY_QTY "
					lgStrSQL = lgStrSQL & " 	FROM	( "
						
					lgStrSQL = lgStrSQL & " 		SELECT	B.PLANT_CD, A.PO_NO, B.PO_SEQ_NO, A.BP_CD, B.ITEM_CD, b.tracking_no , C.PLAN_DVRY_DT,C.PLAN_DVRY_QTY, D.PO_TYPE_NM "
					lgStrSQL = lgStrSQL & " 		 FROM	M_PUR_ORD_HDR A (NOLOCK), M_PUR_ORD_DTL B (NOLOCK),  "
						
					lgStrSQL = lgStrSQL & " 			(SELECT	A.PO_NO, A.PO_SEQ_NO, A.PLAN_DVRY_DT,  A.PLAN_DVRY_QTY  PLAN_DVRY_QTY  "
					lgStrSQL = lgStrSQL & " 			FROM	M_SCM_PLAN_PUR_RCPT A  "
					lgStrSQL = lgStrSQL & " 			WHERE	NOT EXISTS (SELECT * FROM  M_SCM_FIRM_PUR_RCPT B   WHERE A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO   ) "
					lgStrSQL = lgStrSQL & " 			AND	A.SPLIT_SEQ_NO = 0 AND A.PUR_STATUS = 'R'   "
						
					lgStrSQL = lgStrSQL & " 			 UNION ALL "
						
					lgStrSQL = lgStrSQL & " 			SELECT	A.PO_NO, A.PO_SEQ_NO, A.PLAN_DVRY_DT,  A.PLAN_DVRY_QTY - B.PLAN_DVRY_QTY "
					lgStrSQL = lgStrSQL & " 			FROM	M_SCM_PLAN_PUR_RCPT A ,  "
					lgStrSQL = lgStrSQL & " 				(SELECT B.PO_NO, B.PO_SEQ_NO,  SUM(CASE WHEN B.RCPT_QTY = 0 AND RCPT_DT IS NULL THEN B.CONFIRM_QTY ELSE B.RCPT_QTY END ) PLAN_DVRY_QTY  "
					lgStrSQL = lgStrSQL & " 				FROM	M_SCM_FIRM_PUR_RCPT B "
					lgStrSQL = lgStrSQL & " 				GROUP BY B.PO_NO, B.PO_SEQ_NO) B "
					lgStrSQL = lgStrSQL & " 			WHERE	A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO "
					lgStrSQL = lgStrSQL & " 			AND	A.SPLIT_SEQ_NO = 0 AND A.PUR_STATUS = 'R'   "
						
					lgStrSQL = lgStrSQL & " 			UNION ALL "
						
					lgStrSQL = lgStrSQL & " 			SELECT	B.PO_NO, 	B.PO_SEQ_NO, 	B.PLAN_DVRY_DT, SUM(CASE WHEN B.RCPT_QTY = 0 AND RCPT_DT IS NULL THEN B.CONFIRM_QTY ELSE 0 END ) PLAN_DVRY_QTY  "
					lgStrSQL = lgStrSQL & " 			FROM	M_SCM_PLAN_PUR_RCPT A, M_SCM_FIRM_PUR_RCPT B   "
					lgStrSQL = lgStrSQL & " 			WHERE	A.SPLIT_SEQ_NO = 0 AND A.PUR_STATUS = 'R'   "
					lgStrSQL = lgStrSQL & " 			AND	A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO    "
					lgStrSQL = lgStrSQL & " 			GROUP BY B.PO_NO, B.PO_SEQ_NO, B.PLAN_DVRY_DT ) C , "
					lgStrSQL = lgStrSQL & " 			M_CONFIG_PROCESS D  "
						
					lgStrSQL = lgStrSQL & " 		 WHERE	A.PO_NO = B.PO_NO AND B.PO_NO = C.PO_NO  "
					lgStrSQL = lgStrSQL & " 		 AND	B.CLS_FLG = 'N'  "
					lgStrSQL = lgStrSQL & " 		 AND	A.PO_TYPE_CD = D.PO_TYPE_CD  "
					lgStrSQL = lgStrSQL & " 		 AND	B.PO_SEQ_NO = C.PO_SEQ_NO AND C.PLAN_DVRY_QTY > 0 "
					
					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "		and 	B.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					End If					
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "     and 	B.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "     and 	A.bp_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "     and 	C.plan_dvry_dt >= " & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S")
					End If					
					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "     and 	C.plan_dvry_dt <= " & FilterVar(UNIConvDate(lgKeyStream(4)),"''", "S")
					End If
					
					If lgkeystream(5) = "A" Then
						lgStrSQL = lgStrSQL & "     and 	A.RET_FLG = 'N' "
					ElseIf	lgkeystream(5) = "B" Then
						lgStrSQL = lgStrSQL & "     and 	A.RET_FLG = 'Y' "
					End If
					
					If lgkeystream(6) <> "" Then
						lgStrSQL = lgStrSQL & "     and 	B.SL_cd = " & FilterVar(lgKeyStream(6),"''", "S")
					End If
					
					If lgkeystream(7) <> "" Then
						lgStrSQL = lgStrSQL & "     and 	B.tracking_no = " & FilterVar(lgKeyStream(7),"''", "S")
					End If
					
					lgStrSQL = lgStrSQL & " 		 		) A,  "
						
					lgStrSQL = lgStrSQL & " 		 B_ITEM B (NOLOCK), B_BIZ_PARTNER C (NOLOCK)  "
					lgStrSQL = lgStrSQL & " 	WHERE	A.ITEM_CD *= B.ITEM_CD AND A.BP_CD *= C.BP_CD " 
					lgStrSQL = lgStrSQL & " 	GROUP BY A.ITEM_CD, A.TRACKING_NO , A.BP_CD,  A.PLAN_DVRY_DT WITH ROLLUP ) A, "


					lgStrSQL = lgStrSQL & "  	B_ITEM B, B_BIZ_PARTNER C  "
					lgStrSQL = lgStrSQL & "  WHERE A.ITEM_CD *= B.ITEM_CD "
					lgStrSQL = lgStrSQL & "    AND A.BP_CD *= C.BP_CD  "
					lgStrSQL = lgStrSQL & "    AND A.grp_flg NOT IN ('1')  "
					lgStrSQL = lgStrSQL & " order by isNULL(A.ITEM_CD, 'zzzzzzzzzzzzzzzzzz'), isNULL(A.BP_CD,'zzzzzzzzzz'),  isNULL(A.PLAN_DVRY_DT, '2999-12-31'), A.grp_flg  "
					
           End Select

        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "P"
                    lgStrSQL =            " select plant_nm from b_plant where plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
               Case "I"
                    lgStrSQL =            " select item_nm from b_item where item_cd = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
               Case "B"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(2) & "" ,"''", "S")
               Case "S"
                    lgStrSQL =            " select sl_nm from b_storage_location where sl_cd = " & FilterVar(lgKeyStream(6) & "" ,"''", "S")
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
	Dim lsMsg
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case pOpCode
        Case "MC"
        Case "MD"
        Case "MR"
        Case "MU"
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData1
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .lgStrColorFlag = "<%=lgStrColorFlag%>"
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
      
</Script>	                                                                                                                                                                                                                                                                                                                