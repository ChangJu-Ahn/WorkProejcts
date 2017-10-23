<!--'**********************************************************************************************
'*  1. Module Name          : Inventory List onhand stock detail
'*  2. Function Name        : 
'*  3. Program ID           : I2212rb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 현 재고 상세 조회 
'*  6. Comproxy List        : 
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2002/07/06
'*  9. Modifier (First)     : Nam hoon kim
'* 10. Modifier (Last)      : Jung Je Ahn
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->


<%													
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")  
Call HideStatusWnd 

Err.Clear
On Error Resume Next
Dim pPI3S020													
Dim strData
Dim strMode		
Dim lgObjConn ,lgStrSQL, lgObjRs, lgstrData      '2008-05-19 4:33오후 :: hanc										


Dim StrNextKey		
Dim StrNextSubKey      

Dim LngMaxRow		
Dim LngRow
Dim GroupCount
Dim PvArr          
	Const C_SHEETMAXROWS_D	= 100	   
	   
    '-----------------------
    'IMPORTS View
    '-----------------------
	Dim I1_next_i_onhand_stock_tracking_no
	Dim I2_i_onhand_stock_tracking_no
	Dim I3_ief_supplied_select_char
	Dim I4_query_type_action_entry
	Dim I5_b_storage_location_sl_cd
	Dim I6_b_item_item_cd
	Dim I7_b_plant_plant_cd
	
	Dim I8_i_onhand_stock_detail
		Const I303_I8_lot_no = 0
		Const I303_I8_lot_sub_no = 1
	ReDim I8_i_onhand_stock_detail(1)
	
	Dim I9_b_unit_of_measure
		Const I303_I9_unit = 0   
		Const I303_I9_unit_nm = 1
	ReDim I9_b_unit_of_measure(I303_I9_unit_nm)
	
	Dim I10_qty_flag
		Const I303_I10_qty_flag = 0
		Const I303_I10_valid_flag = 1
	ReDim I10_qty_flag(I303_I10_valid_flag)
	
	'-----------------------
	'EXPORTS View
	'-----------------------
	Dim E1_b_plant
		Const I303_E1_plant_cd = 0
		Const I303_E1_plant_nm = 1
	Dim E1_nextt_i_onhand_stock_tracking_no
	Dim E2_next_b_item_item_cd
	Dim E3_next_b_storage_location_sl_cd
	Dim E4_next_i_onhand_stock_detail
		Const I303_E5_lot_no = 0
		Const I303_E5_lot_sub_no = 1
	Dim EG1_export_group
		Const I303_EG1_E1_tracking_no = 0
		Const I303_EG1_E2_sl_cd = 1
		Const I303_EG1_E2_sl_nm = 2
		Const I303_EG1_E3_item_cd = 3
		Const I303_EG1_E3_item_nm = 4
		Const I303_EG1_E3_spec = 5
		Const I303_EG1_E3_basic_unit = 6
		Const I303_EG1_E4_lot_no = 7
		Const I303_EG1_E4_lot_sub_no = 8
		Const I303_EG1_E4_good_on_hand_qty = 9
		Const I303_EG1_E4_bad_on_hand_qty = 10
		Const I303_EG1_E4_stk_on_insp_qty = 11
		Const I303_EG1_E4_stk_on_trns_qty = 12
		Const I303_EG1_E4_picking_qty = 13
		Const I303_EG1_E4_prev_good_qty = 14
		Const I303_EG1_E4_prev_bad_qty = 15
		Const I303_EG1_E4_prev_stk_on_insp_qty = 16
		Const I303_EG1_E4_prev_stk_in_trns_qty = 17
		Const I303_EG1_E5_abc_flg = 18
	    Const I303_EG1_E6_good_on_hand_qty = 19    


	StrNextKey    = Request("lgStrPrevKey")
	StrNextSubKey = Request("lgStrPrevSubKey")
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I7_b_plant_plant_cd								= Request("txtPlant_Cd")
    I5_b_storage_location_sl_cd						= Request("txtSL_Cd")
    I6_b_item_item_cd								= Request("txtItem_Cd")
    I2_i_onhand_stock_tracking_no					= Request("txtTracking_No")
    I4_query_type_action_entry						= "YY"   'N:Item Y:Storage Location 창고별 품목재고정보 
    I3_ief_supplied_select_char						= Request("lgStrUserFlag")    'I:현재고조회 J:LOT조회'
    I9_b_unit_of_measure(I303_I9_unit)				= Request("txtTrns_Unit")
    I8_i_onhand_stock_detail(I303_I8_lot_no)		= Request("txtLotNo")
    I8_i_onhand_stock_detail(I303_I8_lot_sub_no)	= Request("lgStrePrevSubKey")
    I10_qty_flag(I303_I10_qty_flag)					= Request("txtFlag")

    if Request("lgStrPrevKey") <> "" then
		I8_i_onhand_stock_detail(I303_I8_lot_no)     = StrNextKey
		I8_i_onhand_stock_detail(I303_I8_lot_sub_no) = StrNextSubKey
    end if
    
	    
	Set pPI3S020 = Server.CreateObject("PI3S020.cILstOnhandStkDtl")    	
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError(Err, True) = True Then
		Response.End												
	End If    

    Call pPI3S020.I_LIST_ONHAND_STOCK_DETAIL(gStrGlobalCollection, C_SHEETMAXROWS_D, _
                             I1_next_i_onhand_stock_tracking_no, _
                             I2_i_onhand_stock_tracking_no, _
                             I3_ief_supplied_select_char, _
                             I4_query_type_action_entry, _
                             I5_b_storage_location_sl_cd, _
                             I6_b_item_item_cd, _
                             I7_b_plant_plant_cd, _
                             I8_i_onhand_stock_detail, _
							 I9_b_unit_of_measure, _
							 I10_qty_flag, _
                             E1_nextt_i_onhand_stock_tracking_no, _
                             E2_next_b_item_item_cd, _
                             E3_next_b_storage_location_sl_cd, _
                             E4_next_i_onhand_stock_detail, _
                             EG1_export_group)
                                    
	If CheckSYSTEMError(Err, True) = True Then
		Set pPI3S020 = Nothing
		Response.End
	End If

	Set pPI3S020 = Nothing    

	if IsEmpty(EG1_export_group) then
		Response.End
	end if
	
	strData = ""
	LngMaxRow = CLng(Request("txtMaxRows"))
	GroupCount = ubound(EG1_export_group,1)
	Redim PvArr(ubound(EG1_export_group,1))
	

    Call SubOpenDB(lgObjConn)  '2008-05-19 4:25오후 :: hanc                                                      '☜: Make a DB Connection
	
	For LngRow = 0 To ubound(EG1_export_group,1)

    ''2008-05-19 4:21오후 :: hanc-------------------------------------------------------------------------------
        Dim d_EXT2_CD
        Dim d_BP_CD
        Dim d_BP_NM
        
        lgStrSQL = " SELECT	EXT2_CD,	BP_CD, (SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD = A.BP_CD) BP_NM	"	
    	lgStrSQL = lgStrSQL & " FROM	M_PUR_GOODS_MVMT A                                                      "
    	lgStrSQL = lgStrSQL & " WHERE	LOT_NO		=	" &  FilterVar(EG1_export_group(LngRow, I303_EG1_E4_lot_no),"''", "S")
    	lgStrSQL = lgStrSQL & " AND		ITEM_CD		=	" &  FilterVar(EG1_export_group(LngRow, I303_EG1_E3_item_cd),"''", "S")
    	
        If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                   'R(Read) X(CursorType) X(LockType) 
        Else    
          
'    	   If CDbl(lgStrPrevKey) > 0 Then
'    		  lgObjRs.Move     = CDbl(C_SHEETMAXROWS_D) * CDbl(lgStrPrevKey)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
'    	   End If   
           iDx = 1		
           
           lgstrData = ""
'           lgLngMaxRow       = CLng(Request("txtMaxRows"))
    
           Do While Not lgObjRs.EOF
              lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EXT2_CD"))
              lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
              lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
              
              d_EXT2_CD =   ConvSPChars(lgObjRs("EXT2_CD"))
              d_BP_CD   =   ConvSPChars(lgObjRs("BP_CD"))
              d_BP_NM   =   ConvSPChars(lgObjRs("BP_NM"))
    
              lgObjRs.MoveNext
'    
'              iDx =  iDx + 1
'             If iDx > C_SHEETMAXROWS_D Then
'    			 lgStrPrevKey = lgStrPrevKey + 1
'                 Exit Do
'             End If        
          Loop 
        End If

    
    '-------------------------------------------------------------------------------

        strData =	Chr(11) & ConvSPChars(EG1_export_group(LngRow, I303_EG1_E3_item_cd)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I303_EG1_E3_item_nm)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I303_EG1_E1_tracking_no)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I303_EG1_E4_lot_no)) & _
					Chr(11) & ConvSPChars(EG1_export_group(LngRow, I303_EG1_E4_lot_sub_no)) & _
					Chr(11) & ConvSPChars(d_EXT2_CD) & _
					Chr(11) & ConvSPChars(d_BP_CD) & _
					Chr(11) & ConvSPChars(d_BP_NM) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_good_on_hand_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E6_good_on_hand_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_bad_on_hand_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_stk_on_insp_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_stk_on_trns_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_picking_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_prev_good_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_prev_bad_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_prev_stk_on_insp_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & UniConvNumberDBToCompany(EG1_export_group(LngRow, I303_EG1_E4_prev_stk_in_trns_qty),ggQty.DecPoint,ggQty.RndPolicy,ggQty.RndUnit,0) & _
					Chr(11) & LngMaxRow + LngRow & Chr(11) & Chr(12) 
		PvArr(LngRow) = strData    
	Next
		strData = Join(PvArr, "")

    Call SubCloseDB(lgObjConn)  '2008-05-19 4:26오후 :: hanc
    
	If CheckSYSTEMError(Err, True) = True Then
		Response.End
	End If

    If EG1_export_group(GroupCount, I303_EG1_E4_lot_no) = E4_next_i_onhand_stock_detail(I303_E5_lot_no) AND _
       EG1_export_group(GroupCount, I303_EG1_E4_lot_sub_no) = E4_next_i_onhand_stock_detail(I303_E5_lot_sub_no) then

		StrNextKey = ""
		StrNextSubKey = ""
	else
		StrNextKey     = E4_next_i_onhand_stock_detail(I303_E5_lot_no)
		StrNextSubKey  = E4_next_i_onhand_stock_detail(I303_E5_lot_sub_no)
    End If   	


    Response.Write "<Script Language=vbscript> " & vbcr
    Response.Write "With parent "                & vbcr			
    
    Response.Write "	.hTrns_Unit.value = """ & Request("txtTrns_Unit")                               & """ " & vbcr
	
    Response.Write "	.ggoSpread.Source = .vspdData "             & vbcr
    Response.Write "	.ggoSpread.SSShowData """ & strData & """ " & vbcr
			
    Response.Write "	.lgStrPrevKey =    """ & ConvSPChars(StrNextKey)    & """ " & vbcr
    Response.Write "	.lgStrPrevSubKey = """ & ConvSPChars(StrNextSubKey) & """ " & vbcr
		
    Response.Write "	If .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
  	Response.Write "		.DbQuery								"				& vbCr
  	Response.Write "    Else								"						& vbCr
  	Response.Write "		.DbQueryOK								"				& vbCr
	Response.Write "    End If								"						& vbCr
   
    Response.Write "End With "       & vbcr
    Response.Write "</Script> "      & vbcr

%>

