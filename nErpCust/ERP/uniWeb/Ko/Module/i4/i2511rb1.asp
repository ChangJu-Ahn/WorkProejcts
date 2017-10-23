<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'********************************************************************************************************
'*  1. Module Name          : Inventory           *
'*  2. Function Name        : LOT Related Ord 
'*  3. Program ID           : i2511rb1.asp            *
'*  4. Program Name         :               *
'*  5. Program Desc         : 
'*  7. Modified date(First) : 2000/10/09     
'*  8. Modified date(Last)  : 2000/10/09     
'*  9. Modifier (First)     : Kim Nam Hoon
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :                   *
'* 12. Common Coding Guide  : this mark(¢Ð) means that "Do not change"         *
'*                            this mark(¢Á) Means that "may  change"         *
'*                            this mark(¡Ù) Means that "must change"         *
'* 13. History              :                   *
'*                            2000/10/09 : 4th Iteration
'********************************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%               
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")   
Call HideStatusWnd 

On Error Resume Next
Dim i25138

Dim strData
Dim LngRow
Dim LngMaxRow
Dim intGroupCount 
Dim lgStrPrevKey
Dim PvArr

	Const C_SHEETMAXROWS_D = 100

'-----------------------
'IMPORTS View
'-----------------------
DIM I1_i_lot_history_seq_no
DIM I2_i_lot_master
	Const I421_I2_lot_no = 0
	Const I421_I2_lot_sub_no = 1
ReDIM I2_i_lot_master(I421_I2_lot_sub_no)
DIM I3_b_plant_cd
DIM I4_b_item_cd
'-----------------------
'EXPORTS View
'-----------------------
DIM E1_i_lot_history_seq_no
DIM EG1_group_export
	Const I421_EG1_E1_i_lot_history_seq_no = 0
	Const I421_EG1_E1_i_lot_history_prnt_prodt_ord_no = 1
	Const I421_EG1_E1_i_lot_history_mov_type = 2
	Const I421_EG1_E1_i_lot_history_debit_credit_flag = 3
	Const I421_EG1_E1_i_lot_history_qty = 4
	Const I421_EG1_E1_i_lot_history_so_no = 5
	Const I421_EG1_E1_i_lot_history_so_seq_no = 6
	Const I421_EG1_E1_i_lot_history_trns_type = 7
	Const I421_EG1_E1_i_lot_history_transaction_dt = 8
	Const I421_EG1_E1_i_lot_history_bp_cd = 9
	Const I421_EG1_E1_i_lot_history_wc_cd = 10
	Const I421_EG1_E1_i_lot_history_sl_cd = 11
	Const I421_EG1_E1_i_lot_history_trns_sl_cd = 12
	Const I421_EG1_E1_i_lot_history_item_document_no = 13
	Const I421_EG1_E1_i_lot_history_document_year = 14
	Const I421_EG1_E1_i_lot_history_document_seq_no = 15
	Const I421_EG1_E1_i_lot_history_sub_seq_no = 16
	Const I421_EG1_E1_i_lot_history_delete_flag = 17
	Const I421_EG1_E1_i_lot_history_po_no = 18
	Const I421_EG1_E1_i_lot_history_po_seq_no = 19


	lgStrPrevKey     = Request("lgStrPrevKey")
	 
	I3_b_plant_cd                       = Request("txtPlantCd")
	I4_b_item_cd                        = Request("txtItemCd")
	I2_i_lot_master(I421_I2_lot_no)     = Request("txtLotNo")
	I2_i_lot_master(I421_I2_lot_sub_no) = Request("txtLotSubNo")
	I1_i_lot_history_seq_no             = Request("txtSeq")

	if lgStrPrevKey <> "" then I1_i_lot_history_seq_no = lgStrPrevKey
	 
	If CheckSYSTEMError(Err, True) = True Then
		Response.End            
	End If    
	 
	Set i25138 = Server.CreateObject("PI4G070.cILstLotRelateOrdSvr")
	    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End            
	End If    
	 
	Call i25138.I_LIST_LOT_HISTORY(gStrGlobalCollection, C_SHEETMAXROWS_D, _
								I1_i_lot_history_seq_no, _
								I2_i_lot_master, _
								I3_b_plant_cd, _
								I4_b_item_cd, _
								E1_i_lot_history_seq_no, _
								EG1_group_export)

	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Set i25138 = Nothing          
		Response.End              
	End If

	Set i25138 = Nothing

	if isEmpty(EG1_group_export) then
		Response.End            
	END IF

	'Group Count
	intGroupCount = Ubound(EG1_group_export,1)
	LngMaxRow     = Request("txtMaxRows")
	strData       = ""
	ReDim PvArr(Ubound(EG1_group_export,1))
	 
	For LngRow = 0 To intGroupCount
		strData =	Chr(11) & ConvSPChars(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_mov_type)) & _
		Chr(11) & UniConvNumberDBToCompany(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
		Chr(11) & ConvSPChars(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_so_no)) & _
		Chr(11) & ConvSPChars(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_so_seq_no)) & _
		Chr(11) & ConvSPChars(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_prnt_prodt_ord_no)) & _
		Chr(11) & ConvSPChars(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_po_no)) & _
		Chr(11) & ConvSPChars(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_po_seq_no)) & _
		Chr(11) & ConvSPChars(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_transaction_dt)) & _
		Chr(11) & ConvSPChars(EG1_group_export(LngRow, I421_EG1_E1_i_lot_history_seq_no)) & _
		Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
	PvArr(LngRow) = strData
	Next
	
	strData = Join(PvArr, "")
	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If

	If EG1_group_export(intGroupCount, I421_EG1_E1_i_lot_history_seq_no) = E1_i_lot_history_seq_no then    
		lgStrPrevKey = "" 
	Else
		lgStrPrevKey = E1_i_lot_history_seq_no 
	End If

	Response.Write "<Script Language=vbscript> " & vbCr
	Response.Write "With parent "                & vbCr
	    
	Response.Write " .ggoSpread.Source = .vspdData "             & vbCr
	Response.Write " .ggoSpread.SSShowData """ &  strData & """" & vbCr
	Response.Write " .vspdData.focus "                           & vbCr
	    
	Response.Write " .lgStrPrevKey = """ & ConvSPChars(lgStrPrevKey) & """"                  & vbCr
	Response.Write " If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> """" Then " & vbCr
	Response.Write "   .DbQuery "                                                            & vbCr
	Response.Write " Else "                                                                  & vbCr
	Response.Write "   .DbQueryOk "                                                          & vbCr
	Response.Write "    End If "                                                             & vbCr
	    
	Response.Write "End with " & vbCr
	Response.Write "</Script> " & vbCr

%>
