<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : VMI 출고등록 조회 
'*  3. Program ID           : I1523MB1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/06
'*  8. Modified date(Last)  : 2003/04/25
'*  9. Modifier (First)     : Choi Sung Jae
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              : VB Conversion
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")  
Call HideStatusWnd 

Err.Clear
On Error Resume Next
	
Dim subPI5S110								

Const C_SHEETMAXROWS_D = 100

Dim	I1_VMI_goods_mvmt_hdr
	Const I152_I1_item_document_no = 0
	Const I152_I1_document_year = 1
	Const I152_I1_trns_type = 2
	Const I152_I1_plant_cd = 3
ReDim I1_VMI_goods_mvmt_hdr(I152_I1_plant_cd)	
Dim	I2_i_goods_movement_date
	Const I152_I2_seq_no = 0
	Const I152_I2_sub_seq_no = 1
ReDim I2_i_goods_movement_date(I152_I2_sub_seq_no)	

Dim	E1_i_VMI_storage_location
	Const I152_E1_sl_cd = 0
	Const I152_E1_sl_nm = 1
Dim	E2_b_biz_partner
	Const I152_E2_bp_cd = 0
	Const I152_E2_bp_nm = 1
Dim	E3_i_VMI_goods_mvmt_hdr
	Const I152_E3_item_document_no = 0
	Const I152_E3_document_year = 1
	Const I152_E3_document_dt = 1
	Const I152_E3_document_text = 2
Dim	E4_i_VMI_goods_mvmt_dtl
	Const I152_E4_seq_no = 0
	Const I152_E4_sub_seq_no = 1
Dim	EG1_i_VMI_goods_mvmt_dtl
	Const I152_EG1_seq_no = 0
	Const I152_EG1_sub_seq_no = 1
	Const I152_EG1_b_item_item_cd = 2
	Const I152_EG1_b_item_item_nm = 3
	Const I152_EG1_b_item_spec = 4
	Const I152_EG1_lot_no = 5
	Const I152_EG1_lot_sub_no = 6
	Const I152_EG1_entry_qty = 7
	Const I152_EG1_base_unit = 8
	Const I152_EG1_entry_unit = 9
	Const I152_EG1_tracking_no = 10
	Const I152_EG1_insp_flag = 11
	Const I152_EG1_insp_req_no = 12
	Const I152_EG1_insp_status = 13

Dim lgStrPrevKey		
Dim lgStrPrevKey1	

Dim LngMaxRow		
Dim LngRow
Dim GroupCount          
Dim strData
Dim PvArr	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	I1_VMI_goods_mvmt_hdr(I152_I1_item_document_no) = Request("txtItemDocumentNo")
	I1_VMI_goods_mvmt_hdr(I152_I1_document_year)    = Request("txtDocumentYear")
	I1_VMI_goods_mvmt_hdr(I152_I1_trns_type)        = "VI"
	I1_VMI_goods_mvmt_hdr(I152_I1_plant_cd)         = Request("txtPlantCd")

	I2_i_goods_movement_date(I152_I2_seq_no)        = Request("lgStrPrevKey")
	I2_i_goods_movement_date(I152_I2_sub_seq_no)    = Request("lgStrPrevKey1")
	    
    LngMaxRow = CLng(Request("txtMaxRows"))

    Set subPI5S110 = Server.CreateObject("PI5S110.clVMILookUpGoodsMvmt")
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End												
	End If    

    '-----------------------
    'Com Action Area
    '-----------------------
    Call subPI5S110.I_VMI_LOOK_UP_GOODS_MVMT(gStrGlobalCollection, C_SHEETMAXROWS_D, _
                           I1_VMI_goods_mvmt_hdr, _
                           I2_i_goods_movement_date, _
                           E1_i_VMI_storage_location, _
                           E2_b_biz_partner, _
                           E3_i_VMI_goods_mvmt_hdr, _
                           E4_i_VMI_goods_mvmt_dtl, _
                           EG1_i_VMI_goods_mvmt_dtl) 	
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err, True) = True Then
    	Set subPI5S110 = Nothing											
		Response.End														
	End If

  	Set subPI5S110 = Nothing												

	if isEmpty(EG1_i_VMI_goods_mvmt_dtl) then
		Response.End
	end if 
	
	lgStrPrevKey  = E4_i_VMI_goods_mvmt_dtl(I152_E4_seq_no)
	lgStrPrevKey1 = E4_i_VMI_goods_mvmt_dtl(I152_E4_sub_seq_no)	

	GroupCount = ubound(EG1_i_VMI_goods_mvmt_dtl,1)
	ReDim PvArr(GroupCount)		
	
    If EG1_i_VMI_goods_mvmt_dtl(GroupCount, I152_EG1_seq_no)     = lgStrPrevKey and _
       EG1_i_VMI_goods_mvmt_dtl(GroupCount, I152_EG1_sub_seq_no) = lgStrPrevKey1 then

    	lgStrPrevKey  = ""
    	lgStrPrevKey1 = ""
	End If 

	For LngRow = 0 To GroupCount

	    strData = Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_b_item_item_cd)) & _
				  Chr(11) & "" & _                                                                    
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_b_item_item_nm)) & _
				  Chr(11) & UniConvNumberDBToCompany(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_entry_qty), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0) & _
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_entry_unit)) & _
				  Chr(11) & "" & _                                                               
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_tracking_no)) & _
				  Chr(11) & "" & _                                                                 
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_lot_no)) & _     
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_lot_sub_no)) & _ 
				  Chr(11) & "" & _                                                                 
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_b_item_spec)) & _
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_base_unit)) & _  
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_seq_no)) & _     
				  Chr(11) & ConvSPChars(EG1_i_VMI_goods_mvmt_dtl(LngRow, I152_EG1_sub_seq_no)) & _ 
				  Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12)
		PvArr(LngRow) = strData
	Next
	strData = Join(PvArr,"")
	
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write "With Parent" & vbcr
	Response.Write "	.frm1.txtSlCd.value = """ & ConvSPChars(E1_i_VMI_storage_location(I152_E1_sl_cd)) & """" & vbcr
	Response.Write "	.frm1.txtSlNm.value = """ & ConvSPChars(E1_i_VMI_storage_location(I152_E1_sl_nm)) & """" & vbcr

	Response.Write "	.frm1.txtBpCd.value = """ & ConvSPChars(E2_b_biz_partner(I152_E2_bp_cd)) & """" & vbcr
	Response.Write "	.frm1.txtBpNm.value = """ & ConvSPChars(E2_b_biz_partner(I152_E2_bp_nm)) & """" & vbcr

	Response.Write "	.frm1.txtDocumentDt.Text       = """ & UNIDateClientFormat(E3_i_VMI_goods_mvmt_hdr(I152_E3_document_dt)) & """" & vbcr
	Response.Write "	.frm1.txtItemDocumentNo2.value = """ & ConvSPChars(E3_i_VMI_goods_mvmt_hdr(I152_E3_item_document_no))    & """" & vbcr
	Response.Write "	.frm1.txtDocumentText.value    = """ & ConvSPChars(E3_i_VMI_goods_mvmt_hdr(I152_E3_document_text))       & """" & vbcr

	Response.Write "    .ggoSpread.Source = .frm1.vspdData" & vbcr
	Response.Write "    .ggoSpread.SSShowData """ & strData & """" & vbcr

	Response.Write "    .lgStrPrevKey  = """ & ConvSPChars(lgStrPrevKey)  & """" & vbcr
	Response.Write "    .lgStrPrevKey1 = """ & ConvSPChars(lgStrPrevKey1) & """" & vbcr
	Response.Write "	If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
	Response.Write "		.DbQuery								"				& vbCr
	Response.Write "    Else								"				& vbCr
	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"				& vbCr
	Response.Write "End With"       & vbcr
	Response.Write "</Script>"      & vbcr

   	Response.End	

%>

