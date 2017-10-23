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
	Dim lgSeq
	Dim lgSumQty
	Dim lgQty, lgQty2
	Dim lgOrgQty
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

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
		    Call DisplayMsgBox("SCM002", vbInformation, "", "", I_MKSCRIPT)
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

	If lgKeyStream(1) <> "" then 'AND lgErrorStatus <> "YES" Then
   
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

	If lgKeyStream(2) <> "" then 'AND lgErrorStatus <> "YES" Then
   
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
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("Ret_flg2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))            
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PLAN_DVRY_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("D_BP_CD"))
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_UNIT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPLIT_SEQ_NO"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("REQ_DT"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("REQ_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PO_DT"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DLVY_NO")) 
			If lgObjRs("RCPT_FLG") = "N" Then 
				lgstrData = lgstrData & Chr(11) & "출고"
			Else
				lgstrData = lgstrData & Chr(11) & "입고"
			End If 
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
    Call SubCloseRs(lgObjRs)

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)

        Select Case arrColVal(0)
            Case "C"                            '☜: Create
                    Call SubBizSaveMultiCreate(arrColVal)
            Case "U"                            '☜: Update
					Call SubBizSaveCheck(arrColVal)
					If Trim(lgErrorStatus) = "NO" Then
	                    Call SubBizSaveMultiUpdate(arrColVal)
					End If
            Case "D"							'☜: Delete
					Call SubBizDelCheck(arrColVal)
					If Trim(lgErrorStatus) = "NO" Then
						Call SubBizSaveMultiDelete(arrColVal)
					End If
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveCheck
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveCheck(arrColVal)
    On Error Resume Next
    Err.Clear

'--입고량 
    lgStrSQL =            " SELECT isNULL(SUM(RCPT_QTY),0) RCPT_QTY "
	lgStrSQL = lgStrSQL & "   FROM M_PUR_ORD_DTL(NOLOCK) "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO		= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S") 
	lgStrSQL = lgStrSQL & "    AND PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D")

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgSumQty = lgObjRs("RCPT_QTY")
    End If
       
'--발주량 
	lgStrSQL =            " SELECT (CASE WHEN B.OVER_TOL = 0 THEN  A.PLAN_DVRY_QTY "
	lgStrSQL = lgStrSQL & "             ELSE ((100 + B.OVER_TOL) * A.PLAN_DVRY_QTY) / 100 "
	lgStrSQL = lgStrSQL & "        END) PLAN_DVRY_QTY "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_PLAN_PUR_RCPT A(NOLOCK),  "
	lgStrSQL = lgStrSQL & "        M_PUR_ORD_DTL B(NOLOCK) "
	lgStrSQL = lgStrSQL & "  WHERE A.PO_NO = B.PO_NO  "
	lgStrSQL = lgStrSQL & "    AND A.PO_SEQ_NO = B.PO_SEQ_NO "
	lgStrSQL = lgStrSQL & "    AND A.PO_NO		= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S")
	lgStrSQL = lgStrSQL & "    AND A.PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D")
    lgStrSQL = lgStrSQL & "    AND A.SPLIT_SEQ_NO	= 0  "

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgQty = lgObjRs("PLAN_DVRY_QTY")
    End If

'--    
    lgStrSQL =            " SELECT isNULL(SUM(PLAN_DVRY_QTY),0) PLAN_DVRY_QTY "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_FIRM_PUR_RCPT(NOLOCK) "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO		= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S") 
	lgStrSQL = lgStrSQL & "    AND PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D")
    lgStrSQL = lgStrSQL & "    AND SPLIT_SEQ_NO	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","D")
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgOrgQty = lgObjRs("PLAN_DVRY_QTY")
    End If

'--미입고량    
    lgStrSQL =            " SELECT isnull(SUM(CASE WHEN rcpt_qty = 0 AND rcpt_dt is NULL THEN ISNULL(confirm_qty,0) ELSE 0 END),0) unrcpt_firm_qty "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_FIRM_PUR_RCPT(NOLOCK) "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO		= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S") 
	lgStrSQL = lgStrSQL & "    AND PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D")

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgQty2 = lgObjRs("unrcpt_firm_qty")
    End If
    
	If Cdbl(lgQty) < (Cdbl(lgQty2) + Cdbl(lgSumQty) + (Cdbl(arrColVal(6)) - Cdbl(lgOrgQty))) Then
	    Call DisplayMsgBox("SCM001", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus
	End If

    lgStrSQL =            " SELECT (PLAN_DVRY_QTY - isNULL(RCPT_QTY,0)) REMAIN_QTY "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_FIRM_PUR_RCPT(NOLOCK) "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO		= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S") 
	lgStrSQL = lgStrSQL & "    AND PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D")
    lgStrSQL = lgStrSQL & "    AND SPLIT_SEQ_NO	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","D")
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgOrgQty = lgObjRs("REMAIN_QTY")
    End If
    
	If Cdbl(lgOrgQty) <= 0 Then
	    Call DisplayMsgBox("SCM004", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus
	End If
End Sub

'============================================================================================================
' Name : SubBizDelCheck
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizDelCheck(arrColVal)
    On Error Resume Next
    Err.Clear

	lgStrSQL =            " SELECT isNULL(RCPT_QTY,0) RCPT_QTY "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_FIRM_PUR_RCPT(NOLOCK) "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO		= " & FilterVar(Trim(UCase(arrColVal(2))),"","S") 
	lgStrSQL = lgStrSQL & "    AND PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D")
    lgStrSQL = lgStrSQL & "    AND SPLIT_SEQ_NO	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","D")
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgOrgQty = lgObjRs("RCPT_QTY")
    End If
    
	If Cdbl(lgOrgQty) > 0 Then
	    Call DisplayMsgBox("SCM005", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus
	End If
	
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
    
    lgStrSQL = "UPDATE  M_SCM_FIRM_PUR_RCPT "
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " PLAN_DVRY_DT    = " &  FilterVar(Trim(UCase(UniConvDate(arrColVal(5)))),"''","S")   & ","
    lgStrSQL = lgStrSQL & " PLAN_DVRY_QTY   = " &  FilterVar(Trim(UCase(arrColVal(6))),"","D")   & ","
    lgStrSQL = lgStrSQL & " CONFIRM_YN		= " &  FilterVar(Trim(UCase(arrColVal(7))),"''","S")   & ","
    lgStrSQL = lgStrSQL & " D_BP_CD			= " &  FilterVar(Trim(UCase(arrColVal(8))),"''","S")   & ","
    lgStrSQL = lgStrSQL & " CONFIRM_QTY     = " &  FilterVar(Trim(UCase(arrColVal(6))),"","D")   & ","
    lgStrSQL = lgStrSQL & " DLVY_NO			= null, "
    lgStrSQL = lgStrSQL & " DLVY_SEQ_NO     = 0, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID	= " &  FilterVar(gUsrId,"''","S")                      & "," 
    lgStrSQL = lgStrSQL & " UPDT_DT			= " &  FilterVar(GetSvrDateTime,"''","S")
    lgStrSQL = lgStrSQL & " WHERE PO_NO			= " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & "	AND PO_SEQ_NO		= " &  FilterVar(Trim(UCase(arrColVal(3))),"","D")	
    lgStrSQL = lgStrSQL & " AND SPLIT_SEQ_NO	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","D")
        
        Response.Write lgStrSQL
        
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

    lgStrSQL = "DELETE  M_SCM_FIRM_PUR_RCPT "
    lgStrSQL = lgStrSQL & " WHERE PO_NO			= " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")
    lgStrSQL = lgStrSQL & "	AND PO_SEQ_NO		= " &  FilterVar(Trim(UCase(arrColVal(3))),"","D")	
    lgStrSQL = lgStrSQL & " AND SPLIT_SEQ_NO	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","D")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
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
                    lgStrSQL	= " SELECT	TOP " & iSelCount  & " d.item_cd, e.item_nm, e.spec, a.plan_dvry_dt, a.plan_dvry_qty,  A.D_BP_CD , SL_NM ,  d.po_unit, " _
								& "	a.lot_no, a.po_no, a.po_seq_no, a.split_seq_no, b.plan_dvry_dt req_dt, b.plan_dvry_qty req_qty, " _
								& "	c.po_dt, b.ret_flg, c.remark,  g.lot_flg, a.rcpt_qty, d.plant_cd, h.plant_nm , a.dlvy_no ," _
								& "	CASE WHEN b.ret_flg = " & FilterVar("Y","''","S") & " THEN '반품' ELSE '정상' END ret_flg2, " _
								& "	J.rcpt_flg " _
								& " FROM	M_SCM_FIRM_PUR_RCPT a(NOLOCK), M_SCM_PLAN_PUR_RCPT b(NOLOCK), " _
								& "			m_pur_ord_hdr c(NOLOCK), m_pur_ord_dtl d(NOLOCK), b_item e(NOLOCK), b_biz_partner f(NOLOCK), b_item_by_plant g(NOLOCK), " _
								& "			b_plant h(NOLOCK) , B_STORAGE_LOCATION I(NOLOCK), " _
								& "			m_mvmt_type J(NOLOCK) " _
								& " WHERE	a.po_no = b.po_no and a.po_seq_no = b.po_seq_no and d.plant_cd = h.plant_cd " _
								& " and		b.split_seq_no = 0 and a.po_no = c.po_no and a.po_no = d.po_no " _
								& " and		a.po_seq_no = d.po_seq_no and c.bp_cd = f.bp_cd and d.item_cd = e.item_cd " _
								& " and		d.plant_cd = g.plant_cd and d.item_cd = g.item_cd " _
								& " and		A.D_BP_CD *= I.SL_CD "
					
					'--추가내용(1. 입고가완료된건 전체를 입고하지 않고 부분입고되어 잔량이 남은 건은 나타나지 말아야됨)
					lgStrSQL = lgStrSQL & "  and    c.cls_flg = " & FilterVar("N","''","S") _
										& "  and    a.rcpt_qty = 0 " _
										& "  and	a.rcpt_dt is NULL "

					lgStrSQL = lgStrSQL & "  and    c.rcpt_type = J.io_type_cd "
					
					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "  and d.plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
					End If
					
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "  and d.item_cd = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
					End If

					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "  and c.bp_cd = " & FilterVar(lgKeyStream(2) & "" ,"''", "S")
					End If
					
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "  and a.plan_dvry_dt >= " & FilterVar(lgKeyStream(3),"''", "S")
					End If
					
					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "  and a.plan_dvry_dt <= " & FilterVar(lgKeyStream(4),"''", "S")
					End If
					
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "  and c.tracking_no = " & FilterVar(lgKeyStream(5) & "" ,"''", "S")
					End If
					
           End Select
           
        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "P"
                    lgStrSQL = " select plant_nm from b_plant where plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
               Case "I"
                    lgStrSQL = " select item_nm from b_item where item_cd = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")
               Case "B"
                    lgStrSQL = " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(2) & "" ,"''", "S")
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
                .DBQueryOk("<%=lgStrPrevKey%>")
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         
