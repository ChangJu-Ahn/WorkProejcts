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
	Dim lgQty,lgQty2
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '��: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgStrPrevKey	  = UNICInt(Trim(Request("lgStrPrevKey")),0)					 '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
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
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
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

	If lgKeyStream(7) <> "" then 'AND lgErrorStatus <> "YES" Then
   
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
    On Error Resume Next
    Err.Clear

    strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR")
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF

			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PO_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ret_flg2"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PLANT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_UNIT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_SEQ_NO"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("PLAN_DVRY_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("PLAN_DVRY_DT"))
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("UNRCPT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("INSPECT_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("FIRM_DVRY_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("REMAIN_QTY"),ggQty.DecPoint,0)
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("DT"))
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_CD"))
			lgstrData = lgstrData & Chr(11) & ""
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CLS_FLG"))
			If lgObjRs("RCPT_FLG") = "N" Then 
				lgstrData = lgstrData & Chr(11) & "���"
			Else
				lgstrData = lgstrData & Chr(11) & "�԰�"
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
            Case "C"                            '��: Create
					Call SubBizSaveCheck(arrColVal)
					Call SubBizQuerySeq(arrColVal)
                    Call SubBizSaveMultiCreate(arrColVal)
            Case "U"                            '��: Update
					Call SubBizSaveCheck(arrColVal)
                    Call SubBizSaveMultiUpdate(arrColVal)
            Case "D"							'��: Delete
                    Call SubBizSaveMultiDelete(arrColVal)
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
Sub SubBizQuerySeq(arrColVal)
    On Error Resume Next
    Err.Clear

    lgStrSQL =            " SELECT isNULL(MAX(SPLIT_SEQ_NO),0) SPLIT_SEQ_NO "
	lgStrSQL = lgStrSQL & "   FROM M_SCM_FIRM_PUR_RCPT(NOLOCK) "
	lgStrSQL = lgStrSQL & "  WHERE PO_NO = " & FilterVar(Trim(UCase(arrColVal(2))),"","S") 
	lgStrSQL = lgStrSQL & "    AND PO_SEQ_NO = " & FilterVar(Trim(UCase(arrColVal(3))),"","D")

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		If lgObjRs("SPLIT_SEQ_NO") = "" OR lgObjRs("SPLIT_SEQ_NO") = 0 Then
			lgSeq = 1
		Else
			lgSeq = lgObjRs("SPLIT_SEQ_NO") + 1
		End If
    End If
    
End Sub

'============================================================================================================
' Name : SubBizSaveCheck
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveCheck(arrColVal)
    On Error Resume Next
    Err.Clear
	'-- �԰��� üũ(2005-09-10 LSW)
    lgStrSQL	= " SELECT ISNULL(SUM(rcpt_qty),0) RCPT_QTY, CLS_FLG " _
				& " FROM M_PUR_ORD_DTL(NOLOCK) " _
				& " WHERE PO_NO	= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S") _
				& " AND PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D") _
				& " GROUP BY CLS_FLG "

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		If lgObjRs("CLS_FLG") = "Y" Then
			Call DisplayMsgBox("179033", vbInformation, "", "", I_MKSCRIPT)
			Call SetErrorStatus()
			Response.End
			Call SubCloseDB(lgObjConn)
		Else
			lgSumQty = lgObjRs("RCPT_QTY")
		End If
    End If
    
    '-- ���԰��� üũ(2005-09-10 LSW)
    
	lgStrSQL	= " SELECT ISNULL(SUM(CASE WHEN rcpt_qty = 0 AND rcpt_dt is NULL THEN ISNULL(confirm_qty,0) ELSE 0 END),0) unrcpt_firm_qty " _
			 	& " FROM M_SCM_FIRM_PUR_RCPT(NOLOCK) " _
				& " WHERE PO_NO	= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S") _
				& " AND PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D")
   
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgQty2 = lgObjRs("unrcpt_firm_qty")
    End If
	
	'-- ���ַ� üũ(2005-09-10 LSW)
	lgStrSQL	= " SELECT (CASE WHEN B.OVER_TOL = 0 THEN  A.PLAN_DVRY_QTY " _
				& " ELSE ((100 + B.OVER_TOL) * A.PLAN_DVRY_QTY) / 100 " _
				& " END) PLAN_DVRY_QTY " _
				& " FROM M_SCM_PLAN_PUR_RCPT A(NOLOCK), M_PUR_ORD_DTL B(NOLOCK) " _
				& " WHERE A.PO_NO = B.PO_NO " _
				& " AND A.PO_SEQ_NO = B.PO_SEQ_NO " _
				& " AND A.PO_NO	= " & FilterVar(Trim(UCase(arrColVal(2))),"''","S") _
				& " AND A.PO_SEQ_NO	= " & FilterVar(Trim(UCase(arrColVal(3))),"","D") _
				& " AND A.SPLIT_SEQ_NO	= 0 "
   
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        Call SetErrorStatus()
    Else
		lgQty = lgObjRs("PLAN_DVRY_QTY")
    End If
    
    '--���ַ� - �԰��� - ���԰��� <  �Է°�  : ���� , ���ַ� < �԰��� + ���԰��� + �԰���(2005-09-10 LSW)
    If Cdbl(lgQty) < Cdbl(lgQty2) + Cdbl(lgSumQty) + Cdbl(arrColVal(5)) Then
	    Call DisplayMsgBox("SCM001", vbInformation, "", "", I_MKSCRIPT)
		ObjectContext.SetAbort
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

    lgStrSQL = "INSERT INTO M_SCM_FIRM_PUR_RCPT ( "
    lgStrSQL = lgStrSQL & vbCrLf & " PO_NO	, "
    lgStrSQL = lgStrSQL & vbCrLf & " PO_SEQ_NO	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " SPLIT_SEQ_NO	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " PLAN_DVRY_DT	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " PLAN_DVRY_QTY	, " 
    lgStrSQL = lgStrSQL & vbCrLf & " D_BP_CD	, "
    lgStrSQL = lgStrSQL & vbCrLf & " LOT_NO	, "  
    lgStrSQL = lgStrSQL & vbCrLf & " CONFIRM_QTY	, "  
    lgStrSQL = lgStrSQL & vbCrLf & " INSRT_USER_ID	, "
    lgStrSQL = lgStrSQL & vbCrLf & " INSRT_DT		, "  
    lgStrSQL = lgStrSQL & vbCrLf & " UPDT_USER_ID	, "  
    lgStrSQL = lgStrSQL & vbCrLf & " UPDT_DT			) "    
    lgStrSQL = lgStrSQL & vbCrLf & " VALUES			( "
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(2))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(3))),"","D")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & lgSeq  & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UniConvDate(arrColVal(4))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(5))),"","D")   & ","
    
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(6))),"","S")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(7))),"","S")   & ","
'    lgStrSQL = lgStrSQL & vbCrLf & "'*' ,"
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(Trim(UCase(arrColVal(5))),"","D")   & ","
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(gUsrId,"","S")                      & "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(GetSvrDateTime,"''","S")			& "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(gUsrId,"","S")                      & "," 
    lgStrSQL = lgStrSQL & vbCrLf & FilterVar(GetSvrDateTime,"''","S")			& ")"     
    
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next
    Err.Clear
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
                    lgStrSQL	= " select	TOP " & iSelCount  & " c.po_dt, d.item_cd, e.item_nm, e.spec, d.po_unit, a.po_no, a.po_seq_no, " _
								& "			a.plan_dvry_qty, a.plan_dvry_dt ,  d.rcpt_qty, (a.plan_dvry_qty - d.rcpt_qty) unrcpt_qty, D.INSPECT_QTY, " _
								& "			isNULL(b.unrcpt_firm_qty,0) firm_dvry_qty, d.plant_cd, h.plant_nm, " _
								& "			CASE WHEN C.RET_FLG = " & FilterVar("Y","''","S") & " THEN '��ǰ' ELSE '����' END RET_FLG2 , " _
								& "			(a.plan_dvry_qty - d.rcpt_qty - (CASE WHEN ISNULL(b.unrcpt_firm_qty,0) = 0 THEN D.INSPECT_QTY ELSE ISNULL(b.unrcpt_firm_qty,0) END)) REMAIN_QTY, " _
								& "			a.ret_flg, c.remark , GetDate() AS DT , " _
								& "			d.SL_CD , I.SL_NM , " _ 
								& "			' ' LOT_NO, " _ 
								& "			d.cls_flg, J.RCPT_FLG " _ 
								& "   from	M_SCM_PLAN_PUR_RCPT a(NOLOCK) " _
								& "			left outer join (select	po_no, po_seq_no, " _
								& "			SUM(CASE WHEN RCPT_QTY = 0 AND RCPT_DT IS NULL THEN CONFIRM_QTY ELSE 0 END) UNRCPT_FIRM_QTY " _
								& "			FROM	M_SCM_FIRM_PUR_RCPT(NOLOCK) " _
								& "			WHERE 	(PLAN_DVRY_QTY - RCPT_QTY) > 0 " _
								& "			GROUP BY PO_NO, PO_SEQ_NO) b ON A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO " _
								& "			inner join M_PUR_ORD_HDR C(NOLOCK) ON A.PO_NO = C.PO_NO " _
								& "			inner join M_PUR_ORD_DTL D(NOLOCK) ON A.PO_NO = D.PO_NO AND A.PO_SEQ_NO = D.PO_SEQ_NO " _
								& "			inner join B_ITEM E(NOLOCK) ON D.ITEM_CD = E.ITEM_CD " _
								& "			inner join B_BIZ_PARTNER F(NOLOCK) ON C.BP_CD = F.BP_CD " _
								& "			inner join B_PLANT H(NOLOCK) ON D.PLANT_CD = H.PLANT_CD " _
								& "			inner join B_STORAGE_LOCATION I(NOLOCK) ON D.SL_CD = I.SL_CD " _
								& "			inner join M_MVMT_TYPE J(NOLOCK) ON C.RCPT_TYPE = J.IO_TYPE_CD " _
								& "  where	a.split_seq_no = 0 " _
								& "  AND	c.release_flg = " & FilterVar("Y","''","S") _
								& "	 AND	(d.cls_flg = " & FilterVar("N","''","S") & "OR (d.cls_flg = " & FilterVar("Y","''","S") _
								& "  AND	d.rcpt_qty > 0)) " _
								& "	 AND 	(a.plan_dvry_qty - d.rcpt_qty - (CASE WHEN ISNULL(b.unrcpt_firm_qty,0) = 0 THEN D.INSPECT_QTY ELSE ISNULL(b.unrcpt_firm_qty,0) END)) > 0 "

					'--2005-08-02 ���ϰ˻�� ��û���� (����Ȯ���� �Ǹ� ��ȸ�ǰ� ���� )
					lgStrSQL = lgStrSQL & "	 and    c.release_flg = " & FilterVar("Y","''","S")
					'--	
										
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
						lgStrSQL = lgStrSQL & "  and a.plan_dvry_dt >= " & FilterVar(UniConvDate(lgKeyStream(3)),"''", "S")
					End If
					
					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "  and a.plan_dvry_dt <= " & FilterVar(UniConvDate(lgKeyStream(4)),"''", "S")
					End If
					
					lgStrSQL = lgStrSQL & "  and a.plan_dvry_qty > d.rcpt_qty "
					
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "  and c.po_dt >= " & FilterVar(UNIConvDate(lgKeyStream(5)),"''", "S")
					End If
					
					If lgkeystream(6) <> "" Then
						lgStrSQL = lgStrSQL & "  and c.po_dt <= " & FilterVar(UNIConvDate(lgKeyStream(6)),"''", "S")
					End If

					If lgkeystream(7) <> "" Then
						lgStrSQL = lgStrSQL & "  and d.SL_CD Like " & FilterVar(lgKeyStream(7) & "%","''", "S")
					End If
					
					If lgkeystream(8) <> "" Then
						lgStrSQL = lgStrSQL & "  and d.tracking_no = " & FilterVar(lgKeyStream(8) ,"''", "S")
					End If
					
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
                    lgStrSQL =            " select sl_nm from b_storage_location where sl_cd = " & FilterVar(lgKeyStream(7) & "" ,"''", "S")                    
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
	Dim lsMsg
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData1
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"                
                .DBQueryOk("<%=lgStrPrevKey%>")
	         End with
	      Else
				Parent.DBQueryNotOk()
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '�� : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               