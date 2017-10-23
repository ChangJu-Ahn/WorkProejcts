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
	Dim lgOnhandQty, lgBadOnhandQty, lgRcptQty, lgBadRcptQty, lgIssueQty, lgBadIssueQty
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)							 '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

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

	If lgKeyStream(1) <> "" AND lgErrorStatus <> "YES" Then
  
		If lgKeyStream(6) = "NORM" Then
			Call SubMakeSQLStatements("CI")
		ElseIf lgKeyStream(6) = "NEXT" Then
			Call SubMakeSQLStatements("CJ")
		Else
			Call SubMakeSQLStatements("CK")
		End If
     
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf

			If lgKeyStream(6) = "NORM" Then
			    Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			ElseIf lgKeyStream(6) = "NEXT" Then
			    Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
			Else
			    Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)
			End If

		    Call SetErrorStatus()
		Else
			lgKeyStream(1) = lgObjRs("ITEM_CD")

			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtItemCd.value = """ & ConvSPChars(lgObjRs("ITEM_CD")) & """" & vbCrLf
				Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(lgObjRs("ITEM_NM")) & """" & vbCrLf
				Response.Write "parent.frm1.txtSpec.value = """ & ConvSPChars(lgObjRs("SPEC")) & """" & vbCrLf
				Response.Write "parent.frm1.txtUnit.value = """ & ConvSPChars(lgObjRs("BASIC_UNIT")) & """" & vbCrLf
				Response.Write "parent.frm1.txtItemAcct.value = """ & ConvSPChars(lgObjRs("ITEM_ACCT")) & """" & vbCrLf
				Response.Write "parent.frm1.txtProcureType.value = """ & ConvSPChars(lgObjRs("PROCUR_TYPE")) & """" & vbCrLf
				Response.Write "parent.frm1.txtLocation.value = """ & ConvSPChars(lgObjRs("LOCATION")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

	If lgKeyStream(2) <> "" AND lgErrorStatus <> "YES" Then
   
		Call SubMakeSQLStatements("CS")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtSlNm.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    Call DisplayMsgBox("125700", vbInformation, "", "", I_MKSCRIPT)
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
    
     lgOnhandQty = Trim(lgKeyStream(7))
     lgBadOnhandQty = Trim(lgKeyStream(8))
     lgRcptQty = Trim(lgKeyStream(9))
     lgBadRcptQty = Trim(lgKeyStream(10))
     lgIssueQty = Trim(lgKeyStream(11))
     lgBadIssueQty = Trim(lgKeyStream(12))

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

			If lgObjRs("debit_credit_flag") = "X" Then 
				lgOnhandQty		= Cdbl(lgObjRs("GOOD_ON_HAND_QTY"))
				lgBadOnhandQty	= Cdbl(lgObjRs("BAD_ON_HAND_QTY"))
			End If
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SPEC"))
			If lgObjRs("debit_credit_flag") = "X" Then
				lgstrData = lgstrData & Chr(11) & "이월재고"
			ElseIf lgObjRs("debit_credit_flag") = "0" Then
				lgstrData = lgstrData & Chr(11) & "현재고"
            Else
				lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("DOCUMENT_DT"))
			End If
            
			lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("BAS_INV_QTY"),ggQty.DecPoint,0)
			If Cdbl(lgObjRs("RCPT_QTY")) <> 0 Then
				lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("RCPT_QTY"),ggQty.DecPoint,0)
				lgRcptQty = lgRcptQty + Cdbl(lgObjRs("RCPT_QTY"))
			Else
				If lgObjRs("debit_credit_flag") = "0" Then
					lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgRcptQty,ggQty.DecPoint,0)				
				Else
					lgstrData = lgstrData & Chr(11) & 0
				End IF
			End If
			
			If Cdbl(lgObjRs("ISSUE_QTY")) <> 0 Then
				lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("ISSUE_QTY"),ggQty.DecPoint,0)
				lgIssueQty = lgIssueQty + Cdbl(lgObjRs("ISSUE_QTY"))
			Else
				If lgObjRs("debit_credit_flag") = "0" Then
					lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgIssueQty,ggQty.DecPoint,0)
				Else
					lgstrData = lgstrData & Chr(11) & 0
				End IF
			End If
			
			If lgObjRs("debit_credit_flag") = "X" or lgObjRs("debit_credit_flag") = "0" Then
				lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("GOOD_ON_HAND_QTY"),ggQty.DecPoint,0)
			Else
				lgOnhandQty = lgOnhandQty + Cdbl(lgObjRs("RCPT_QTY")) - Cdbl(lgObjRs("ISSUE_QTY"))
				lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgOnhandQty,ggQty.DecPoint,0)				
			End If
			
			
			
			If Cdbl(lgObjRs("BAD_RCPT_QTY")) <> 0 Then
				lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("BAD_RCPT_QTY"),ggQty.DecPoint,0)
				lgBadRcptQty = lgBadRcptQty + Cdbl(lgObjRs("BAD_RCPT_QTY"))
			Else
				If lgObjRs("debit_credit_flag") = "0" Then
					lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgBadRcptQty,ggQty.DecPoint,0)				
				Else
					lgstrData = lgstrData & Chr(11) & 0
				End IF
			End If
			
			If Cdbl(lgObjRs("BAD_ISS_QTY")) <> 0 Then
				lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("BAD_ISS_QTY"),ggQty.DecPoint,0)
				lgBadIssueQty = lgBadIssueQty + Cdbl(lgObjRs("BAD_ISS_QTY"))
			Else
				If lgObjRs("debit_credit_flag") = "0" Then
					lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgBadIssueQty,ggQty.DecPoint,0)
				Else
					lgstrData = lgstrData & Chr(11) & 0
				End IF
			End If
			
			If lgObjRs("debit_credit_flag") = "X" or lgObjRs("debit_credit_flag") = "0" Then
				lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs("BAD_ON_HAND_QTY"),ggQty.DecPoint,0)
			Else
				lgBadOnhandQty = lgBadOnhandQty + Cdbl(lgObjRs("BAD_RCPT_QTY")) - Cdbl(lgObjRs("BAD_ISS_QTY"))
				lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgBadOnhandQty,ggQty.DecPoint,0)				
			End If
			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRNS_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MOV_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_DOCUMENT_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WC_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("WC_NM"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRNS_SL_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TRNS_SL_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOCUMENT_TEXT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PO_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MVMT_RCPT_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PRODT_ORDER_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DN_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("LOT_NO"))
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("PRICE"), 0)
            lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs("AMOUNT"), 0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("INSRT_USER_ID"))
           
            If lgObjRs("debit_credit_flag") = "X" Then 
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "1" & gRowSep
			ElseIf lgObjRs("debit_credit_flag") = "0" Then 
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "3" & gRowSep
			Else
				lgStrColorFlag = lgStrColorFlag & CStr(lgLngMaxRow + iDx) & gColSep & "2" & gRowSep
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
					lgStrSQL =                   "	select	TOP " & iSelCount  & " a.item_cd, b.item_nm, b.spec, a.document_dt, "
					lgStrSQL = lgStrSQL & "             	case when a.debit_credit_flag = 'X' then a.bas_inv_qty + a.rcpt_qty - a.issue_qty + a.bad_rcpt_qty - a.bad_iss_qty else 0 end bas_inv_qty, "
					lgStrSQL = lgStrSQL & "             	case when a.debit_credit_flag = 'D' or  a.debit_credit_flag = 'C' then a.rcpt_qty else 0 end rcpt_qty, "
					lgStrSQL = lgStrSQL & "             	case when a.debit_credit_flag = 'D' or  a.debit_credit_flag = 'C' then a.issue_qty else 0 end issue_qty, "
					lgStrSQL = lgStrSQL & "             	case a.debit_credit_flag when '0' then a.good_on_hand_qty when 'X' then a.rcpt_qty - a.issue_qty else 0 end good_on_hand_qty, "
					
					lgStrSQL = lgStrSQL & "             	case when a.debit_credit_flag = 'D' or  a.debit_credit_flag = 'C' then a.bad_rcpt_qty else 0 end bad_rcpt_qty, "
					lgStrSQL = lgStrSQL & "             	case when a.debit_credit_flag = 'D' or  a.debit_credit_flag = 'C' then a.bad_iss_qty else 0 end bad_iss_qty, "
					lgStrSQL = lgStrSQL & "             	case a.debit_credit_flag when '0' then a.bad_on_hand_qty when 'X' then a.bad_rcpt_qty - a.bad_iss_qty else 0 end bad_on_hand_qty, "
					
					lgStrSQL = lgStrSQL & "             	a.debit_credit_flag, e.minor_nm trns_type, f.minor_nm mov_type, d.prodt_order_no, c.document_text, "
					lgStrSQL = lgStrSQL & "             	d.item_document_no, d.lot_no, d.lot_sub_no, d.sl_cd, g.sl_nm, d.wc_cd, i.wc_nm, a.trns_sl_cd, a.mvmt_rcpt_no, "
					lgStrSQL = lgStrSQL & "             	gg.sl_nm trns_sl_nm, d.so_no, d.bp_cd, h.bp_nm, d.po_no, d.dn_no, j.usr_nm insrt_user_id, a.remark, a.price, a.amount "
					lgStrSQL = lgStrSQL & "			from	(select b.item_cd, 0 bas_inv_qty, "
					lgStrSQL = lgStrSQL & "							sum(case when debit_credit_flag = 'D' and item_status = 'G' then qty else 0 end) rcpt_qty, "
					lgStrSQL = lgStrSQL & "							sum(case when debit_credit_flag = 'C' and item_status = 'G' then qty else 0 end) issue_qty, "
					lgStrSQL = lgStrSQL & "							sum(case when debit_credit_flag = 'D' and item_status = 'B' then qty else 0 end) bad_rcpt_qty, "
					lgStrSQL = lgStrSQL & "							sum(case when debit_credit_flag = 'C' and item_status = 'B' then qty else 0 end) bad_iss_qty, "
					lgStrSQL = lgStrSQL & "							0 good_on_hand_qty, 0 bad_on_hand_qty, "
					lgStrSQL = lgStrSQL & "							" & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S") & " document_dt, 'X' debit_credit_flag, "
					lgStrSQL = lgStrSQL & "							'' trns_type, '' mov_type, '' item_document_no, '' document_year, 0 seq_no, 0 sub_seq_no, "
					lgStrSQL = lgStrSQL & "							'' bp_cd, '' sl_cd, '' wc_cd, '' trns_sl_cd, '' insrt_user_id, '' remark, 0 price, 0 amount, '' mvmt_rcpt_no "
					lgStrSQL = lgStrSQL & "					from 	i_goods_movement_header a, i_goods_movement_detail b, b_storage_location c "
					lgStrSQL = lgStrSQL & "					where 	b.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					lgStrSQL = lgStrSQL & "					and 	a.item_document_no = b.item_document_no and	a.document_year = b.document_year and b.delete_flag = 'N' "
					lgStrSQL = lgStrSQL & "					and 	b.sl_cd = c.sl_cd "
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.sl_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.document_dt < " & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S")
					End If
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	c.sl_group_cd = " & FilterVar(lgKeyStream(5),"''", "S")
					End If
					lgStrSQL = lgStrSQL & "					group by b.item_cd "
					lgStrSQL = lgStrSQL & "					union all " '------------- 기초재고
					lgStrSQL = lgStrSQL & "					select item_cd, 0 bas_inv_qty, 0 rcpt_qty, 0 issue_qty, 0 bad_rcpt_qty, 0 bad_iss_qty, 0 good_on_hand_qty, 0 bad_on_hand_qty,"
					lgStrSQL = lgStrSQL & "							" & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S") & " document_dt, 'X' debit_credit_flag, "
					lgStrSQL = lgStrSQL & "							'' trns_type, '' mov_type, '' item_document_no, '' document_year, 0 seq_no, 0 sub_seq_no, "
					lgStrSQL = lgStrSQL & "							'' bp_cd, '' sl_cd, '' wc_cd, '' trns_sl_cd, '' insrt_user_id, '' remark, 0 price, 0 amount, '' mvmt_rcpt_no "
					lgStrSQL = lgStrSQL & "					from 	i_material_valuation "
					lgStrSQL = lgStrSQL & "					where 	plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					lgStrSQL = lgStrSQL & "					and	item_cd not in (	select	distinct b.item_cd "
					lgStrSQL = lgStrSQL & "											from	i_goods_movement_header a, i_goods_movement_detail b, b_storage_location c "
					lgStrSQL = lgStrSQL & "											where	b.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					lgStrSQL = lgStrSQL & "											and		a.item_document_no = b.item_document_no and	a.document_year = b.document_year "
					lgStrSQL = lgStrSQL & "											and		b.delete_flag = 'N' "
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            							and 	b.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            							and 	b.sl_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "            							and 	a.document_dt < " & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S")
					End If
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "            							and 	c.sl_group_cd = " & FilterVar(lgKeyStream(5),"''", "S")
					End If
					
					lgStrSQL = lgStrSQL & "											and		b.sl_cd = c.sl_cd ) "
					lgStrSQL = lgStrSQL & "					union all "	'------------- 수불
					
					lgStrSQL = lgStrSQL & "					select	b.item_cd, 0 bas_inv_qty,  "
					lgStrSQL = lgStrSQL & "						case when debit_credit_flag = 'D' and item_status = 'G' then qty else 0 end rcpt_qty,  "
					lgStrSQL = lgStrSQL & "						case when debit_credit_flag = 'C' and item_status = 'G' then qty else 0 end issue_qty,  "
					lgStrSQL = lgStrSQL & "						case when debit_credit_flag = 'D' and item_status = 'B' then qty else 0 end bad_rcpt_qty,  "
					lgStrSQL = lgStrSQL & "						case when debit_credit_flag = 'C' and item_status = 'B' then qty else 0 end bad_iss_qty,  "
					lgStrSQL = lgStrSQL & "						0 good_on_hand_qty, 0 bad_on_hand_qty, a.document_dt, "
					lgStrSQL = lgStrSQL & "						b.debit_credit_flag, b.trns_type, b.mov_type, b.item_document_no, b.document_year, b.seq_no, b.sub_seq_no, "
					lgStrSQL = lgStrSQL & "						b.bp_cd, b.sl_cd, b.wc_cd, b.trns_sl_cd, b.insrt_user_id, c.remark, b.price, b.amount, e.mvmt_rcpt_no "
					lgStrSQL = lgStrSQL & "					from 	i_goods_movement_header a, i_goods_movement_detail b, s_dn_dtl c, b_storage_location d, m_pur_goods_mvmt e "
					lgStrSQL = lgStrSQL & "					where 	b.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					lgStrSQL = lgStrSQL & "					and 	a.item_document_no = b.item_document_no and	a.document_year = b.document_year and b.delete_flag = 'N' "
					lgStrSQL = lgStrSQL & "					and 	b.dn_no *= c.dn_no and b.dn_seq *= c.dn_seq "
					lgStrSQL = lgStrSQL & "					and 	b.item_document_no *= e.gm_no and b.document_year *= e.gm_year and b.seq_no *= e.gm_seq_no and b.sub_seq_no *= e.gm_sub_seq_no "
					lgStrSQL = lgStrSQL & "					and 	b.sl_cd = d.sl_cd "
'					lgStrSQL = lgStrSQL & "					and 	(a.plant_cd + a.item_document_no + a.document_year) not in (select (b.plant_cd + b.item_document_no + b.document_year) from I_MOVETYPE_CONFIGURATION a, i_goods_movement_detail b where a.gui_control_flag2 = 'Y' and a.gui_control_flag3 = 'N' and a.mov_type = b.mov_type and b.TRNS_SL_CD like 'WH%' and b.SL_CD like 'WH%') "
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.sl_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If
					If lgkeystream(3) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.document_dt >= " & FilterVar(UNIConvDate(lgKeyStream(3)),"''", "S")
					End If
					If lgkeystream(4) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.document_dt <= " & FilterVar(UNIConvDate(lgKeyStream(4)),"''", "S")
					End If
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	d.sl_group_cd = " & FilterVar(lgKeyStream(5),"''", "S")
					End If
					lgStrSQL = lgStrSQL & "					union all "	'------------- 재고
					lgStrSQL = lgStrSQL & "					select	a.item_cd, 0 bas_inv_qty, 0 rcpt_qty, 0 issue_qty, 0 bad_rcpt_qty, 0 bad_iss_qty, "
					lgStrSQL = lgStrSQL & "					sum(a.good_on_hand_qty) good_on_hand_qty, sum(a.bad_on_hand_qty) bad_on_hand_qty, "
					lgStrSQL = lgStrSQL & "							getdate() document_dt, '0' debit_credit_flag, '' trns_type, '' mov_type, '' item_document_no, "
					lgStrSQL = lgStrSQL & "							'' document_year, 0 seq_no, 0 sub_seq_no, '' bp_cd, '' sl_cd, '' wc_cd, '' trns_sl_cd, '' insrt_user_id, "
					lgStrSQL = lgStrSQL & "							'' remark, 0 price, 0 amount, '' mvmt_rcpt_no "
					lgStrSQL = lgStrSQL & "					from 	i_onhand_stock a, b_storage_location b "
					lgStrSQL = lgStrSQL & "					where 	a.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					lgStrSQL = lgStrSQL & "					and		a.sl_cd = b.sl_cd "
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.sl_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	b.sl_group_cd = " & FilterVar(lgKeyStream(5),"''", "S")
					End If
					lgStrSQL = lgStrSQL & "					group by item_cd "
					lgStrSQL = lgStrSQL & "					union all "	'------------- 재고
					lgStrSQL = lgStrSQL & "					select	item_cd, 0 bas_inv_qty, 0 rcpt_qty, 0 issue_qty,  0 bad_rcpt_qty, 0 bad_iss_qty, "
					lgStrSQL = lgStrSQL & "							0 good_on_hand_qty, 0 bad_on_hand_qty, getdate() document_dt, '0' debit_credit_flag, "
					lgStrSQL = lgStrSQL & "							'' trns_type, '' mov_type, '' item_document_no, '' document_year, 0 seq_no, 0 sub_seq_no, "
					lgStrSQL = lgStrSQL & "							'' bp_cd, '' sl_cd, '' wc_cd, '' trns_sl_cd, '' insrt_user_id, '' remark, 0 price, 0 amount, '' mvmt_rcpt_no "
					lgStrSQL = lgStrSQL & "					from 	i_material_valuation "
					lgStrSQL = lgStrSQL & "					where 	plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					lgStrSQL = lgStrSQL & "					and	item_cd not in (	select	distinct a.item_cd "
					lgStrSQL = lgStrSQL & "											from	i_onhand_stock a, b_storage_location b "
					lgStrSQL = lgStrSQL & "											where	a.plant_cd = " & FilterVar(lgKeyStream(0),"''", "S")
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            							and 	a.item_cd = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            							and 	a.sl_cd = " & FilterVar(lgKeyStream(2),"''", "S")
					End If
					If lgkeystream(5) <> "" Then
						lgStrSQL = lgStrSQL & "            							and 	b.sl_group_cd = " & FilterVar(lgKeyStream(5),"''", "S")
					End If
					lgStrSQL = lgStrSQL & "											and		a.sl_cd = b.sl_cd ) ) a, "
					lgStrSQL = lgStrSQL & "					b_item b, i_goods_movement_header c, i_goods_movement_detail d, b_minor e, b_minor f, "
					lgStrSQL = lgStrSQL & "					b_storage_location g, b_storage_location gg, b_biz_partner h, p_work_center i, z_usr_mast_rec j "
					lgStrSQL = lgStrSQL & "			where	a.item_cd = b.item_cd and a.item_document_no *= c.item_document_no and	a.document_year *= c.document_year "


					lgStrSQL = lgStrSQL & "			and		a.item_document_no *= d.item_document_no and a.document_year *= d.document_year "
					lgStrSQL = lgStrSQL & "			and		a.seq_no *= d.seq_no and a.sub_seq_no *= d.sub_seq_no "
					lgStrSQL = lgStrSQL & "			and		e.major_cd = 'I0002' and a.trns_type *= e.minor_cd "
					lgStrSQL = lgStrSQL & "			and		f.major_cd = 'I0001' and a.mov_type *= f.minor_cd "
					lgStrSQL = lgStrSQL & "			and		a.sl_cd *= g.sl_cd "
					lgStrSQL = lgStrSQL & "			and		a.trns_sl_cd *= gg.sl_cd "
					lgStrSQL = lgStrSQL & "			and		a.bp_cd *= h.bp_cd "
					lgStrSQL = lgStrSQL & "			and		a.wc_cd *= i.wc_cd "
					lgStrSQL = lgStrSQL & "			and		a.insrt_user_id *= j.usr_id "
					lgStrSQL = lgStrSQL & "			order by a.item_cd, a.document_dt, a.debit_credit_flag desc "
				
           End Select     
           
        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "P"
                    lgStrSQL =            " select plant_nm from b_plant where plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
               Case "I"
                    lgStrSQL =            " select	a.item_cd, a.item_nm, a.spec, a.basic_unit, c.minor_nm item_acct, d.minor_nm procur_type, b.location "
                    lgStrSQL = lgStrSQL & " from	b_item a, b_item_by_plant b, b_minor c, b_minor d "
                    lgStrSQL = lgStrSQL & " where	a.item_cd = b.item_cd and b.item_acct = c.minor_cd and c.major_cd = 'P1001' "
                    lgStrSQL = lgStrSQL & " and		b.procur_type = d.minor_cd and d.major_cd = 'P1003' "
                    lgStrSQL = lgStrSQL & " and		b.plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
                    lgStrSQL = lgStrSQL & " and		a.item_cd = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")

'response.write lgStrSQL
'response.end 

               Case "J" ' Next
                    lgStrSQL =            " select	a.item_cd, a.item_nm, a.spec, a.basic_unit, c.minor_nm item_acct, d.minor_nm procur_type, b.location "
                    lgStrSQL = lgStrSQL & " from	b_item a, b_item_by_plant b, b_minor c, b_minor d "
                    lgStrSQL = lgStrSQL & " where	a.item_cd = b.item_cd and b.item_acct = c.minor_cd and c.major_cd = 'P1001' "
                    lgStrSQL = lgStrSQL & " and		b.procur_type = d.minor_cd and d.major_cd = 'P1003' "
                    lgStrSQL = lgStrSQL & " and		b.plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
                    lgStrSQL = lgStrSQL & " and		a.item_cd = (select Min(item_cd) from b_item where item_cd > " & FilterVar(lgKeyStream(1) & "" ,"''", "S") & ")" 

               Case "K" ' Previous
                    lgStrSQL =            " select	a.item_cd, a.item_nm, a.spec, a.basic_unit, c.minor_nm item_acct, d.minor_nm procur_type, b.location "
                    lgStrSQL = lgStrSQL & " from	b_item a, b_item_by_plant b, b_minor c, b_minor d "
                    lgStrSQL = lgStrSQL & " where	a.item_cd = b.item_cd and b.item_acct = c.minor_cd and c.major_cd = 'P1001' "
                    lgStrSQL = lgStrSQL & " and		b.procur_type = d.minor_cd and d.major_cd = 'P1003' "
                    lgStrSQL = lgStrSQL & " and		b.plant_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
                    lgStrSQL = lgStrSQL & " and		a.item_cd = (select Max(item_cd) from b_item where item_cd < " & FilterVar(lgKeyStream(1) & "" ,"''", "S") & ")" 
               Case "S"
                    lgStrSQL =            " select sl_nm from b_storage_location where sl_cd = " & FilterVar(lgKeyStream(2) & "" ,"''", "S")
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
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .lgStrColorFlag = "<%=lgStrColorFlag%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"

		' 2005-02-28 Begin
		.lgOnhandQty	= "<%=lgOnhandQty%>"
		.lgBadOnhandQty	= "<%=lgBadOnhandQty%>"
		.lgRcptQty	= "<%=lgRcptQty%>"
		.lgBadRcptQty	= "<%=lgBadRcptQty%>"
		.lgIssueQty	= "<%=lgIssueQty%>"
		.lgBadIssueQty	= "<%=lgBadIssueQty%>"
		' 2005-02-28 End

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