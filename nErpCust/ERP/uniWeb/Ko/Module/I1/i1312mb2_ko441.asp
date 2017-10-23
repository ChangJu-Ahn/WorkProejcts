<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : 재고이동등록  저장 업무 처리 ASP
'*  2. Function Name        : 
'*  3. Program ID           : i1311mb2.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : PI0C180.cIStockTransfer

'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2003/05/13
'*  9. Modifier (First)     : Mr  Kim Nam Hoon
'* 10. Modifier (Last)      : Mr  Ahn Jung Je
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next											
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
    Call HideStatusWnd											

	Dim iPI0C180									
	
	Dim lgIntFlgMode
	DIm LngRow
	Dim LngMaxRow
	
	Dim arrRowVal			
	Dim arrColVal			
		Const C_ItemCd			= 2
		Const C_TrackingNo		= 3
		Const C_LotNo			= 4
		Const C_LotSubNo		= 5
		Const C_EntryQty		= 6
		Const C_EntryUnit		= 7
		Const C_TrnsItemCd		= 8
		Const C_TrnsPlantCd		= 9
		Const C_TrnsSLCd		= 10
		Const C_TrnsLotNo		= 11
		Const C_TrnsLotSubNo	= 12
		Const C_TrnsTrackingNo	= 13
		Const C_EntryQtyAfter	= 14

		Const D_SeqNo  = 2
		Const D_ItemCd = 3
	Dim strStatus							
    
	Dim IntLotSubNo			
	Dim IntTrnsLotSubNo
	
	Dim I2_good_mvmt_workset 	
		Const C_I2_item_document_no		= 0
		Const C_I2_trns_type			= 2
		Const C_I2_mov_type				= 3
		Const C_I2_document_dt			= 4
		Const C_I2_pos_dt				= 5
		Const C_I2_document_text		= 6
		Const C_I2_plant				= 7
		Const C_I2_cost_cd				= 8
		Const C_I2_biz_ared_cd			= 9
    Redim I2_good_mvmt_workset(C_I2_biz_ared_cd)
    
    Dim DI2_good_mvmt_workset 	
		Const C_DI2_item_document_no	= 0
		Const C_DI2_year				= 1
		Const C_DI2_trns_type			= 2
		Const C_DI2_mov_type			= 3
		Const C_DI2_document_dt			= 4
		Const C_DI2_pos_dt				= 5
		Const C_DI2_plant				= 6
		Const C_DI2_document_text		= 7
    Redim DI2_good_mvmt_workset(C_DI2_document_text)

    Dim IG1_import_group
		Const C_IG1_I1_sl_cd				= 0
		Const C_IG1_I2_lot_no				= 1
		Const C_IG1_I2_lot_sub_no			= 2
		Const C_IG1_I2_item_status			= 3
		Const C_IG1_I2_entry_qty			= 4
		Const C_IG1_I2_entry_unit			= 5
		Const C_IG1_I2_trns_lot_no			= 6
		Const C_IG1_I2_trns_lot_sub_no		= 7
		Const C_IG1_I2_trns_plant_cd		= 8
		Const C_IG1_I2_trns_sl_cd			= 9
		Const C_IG1_I2_trns_item_cd			= 10
		Const C_IG1_I2_tracking_no			= 12
		Const C_IG1_I2_trns_tracking_no		= 13
		Const C_IG1_I3_plant_cd				= 18
		Const C_IG1_I4_item_cd				= 19
		Const C_IG1_I5_cost_cd				= 20
		Const C_IG1_I2_return_flag			= 32
		Const C_IG1_I2_entry_qty_after  	= 33    '20080313::hanc::이동후수량
	
	Dim dIG1_import_group 		
		Const D_IG1_sl_cd		= 0
		Const D_IG1_seq_no		= 2
		Const D_IG1_plant_cd	= 4
		Const D_IG1_item_cd		= 5
			
	Dim E1_good_mvmt_workset 
		Const C_E1_item_document_no		= 0
		Const C_E1_year					= 1
    
    Dim iErrorPosition
    Redim iErrorPosition(0)

 	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim i
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For i = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(i)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")
    
       
	'-----------------------
	'Data manipulate area
	'-----------------------
	I2_good_mvmt_workset(C_I2_item_document_no) = Trim(Request("txtDocumentNo2"))	
	I2_good_mvmt_workset(C_I2_trns_type) 	    = "ST"
	I2_good_mvmt_workset(C_I2_mov_type) 	    = Request("txtMovType")
	I2_good_mvmt_workset(C_I2_document_dt)	    = UNIConvDate(Request("txtDocumentDt"))
	I2_good_mvmt_workset(C_I2_pos_dt)  		    = UNIConvDate(Request("txtPostingDt"))
	I2_good_mvmt_workset(C_I2_document_text)    = Request("txtDocumentText")
	I2_good_mvmt_workset(C_I2_plant) 		    = Request("txtPlantCd1")
	I2_good_mvmt_workset(C_I2_cost_cd)		    = Request("txtCostCd1")	


	DI2_good_mvmt_workset(C_DI2_item_document_no) = Trim(Request("txtDocumentNo1"))
    DI2_good_mvmt_workset(C_DI2_year)             = Request("hYear")
    DI2_good_mvmt_workset(C_DI2_trns_type)        = "ST"
    DI2_good_mvmt_workset(C_DI2_mov_type)         = Request("txtMovType")
    DI2_good_mvmt_workset(C_DI2_document_dt)      = UNIConvDate(Request("txtDocumentDt"))
    DI2_good_mvmt_workset(C_DI2_pos_dt)           = UNIConvDate(Request("txtPostingDt"))
    DI2_good_mvmt_workset(C_DI2_plant)            = Request("txtPlantCd1")
    DI2_good_mvmt_workset(C_DI2_document_text)    = Request("txtDocumentText")


    If itxtSpread <> "" Then
	
	   arrRowVal = Split(itxtSpread, gRowSep)
	   LngMaxRow = UBound(arrRowVal) - 1

		Redim IG1_import_group(LngMaxRow, C_IG1_I2_entry_qty_after) '20080313::hanc::기존->C_IG1_I2_return_flag    
		Redim dIG1_import_group(LngMaxRow, D_IG1_item_cd)
	   
	   For LngRow = 0 To LngMaxRow
	  
		arrColVal = Split(arrRowVal(LngRow), gColSep)
		strStatus = arrColVal(0)  		

		Select Case strStatus
		Case "C"			
	
			If arrColVal(C_TrnsLotNo) <> "" Then 
			     IG1_import_group(LngRow,C_IG1_I2_tracking_no)	= arrColVal(C_TrackingNo)
			Else
			     IG1_import_group(LngRow,C_IG1_I2_tracking_no)	= "*"
			End if
			
			if arrColVal(C_LotSubNo) = "" then
			   IntLotSubNo = 0
			else
			   IntLotSubNo = arrColVal(C_LotSubNo)
			End if
			
			If arrColVal(C_TrnsItemCd) <> "" then	
				IG1_import_group(LngRow,C_IG1_I2_trns_item_cd)	= arrColVal(C_TrnsItemCd)
			Else	
				IG1_import_group(LngRow,C_IG1_I2_trns_item_cd)	= arrColVal(C_ItemCd)
			End if
			
			If arrColVal(C_TrnsPlantCd) <> "" then
				IG1_import_group(LngRow,C_IG1_I2_trns_plant_cd)	= arrColVal(C_TrnsPlantCd)
			Else
				IG1_import_group(LngRow,C_IG1_I2_trns_plant_cd)	= Request("txtPlantCd1")
			End if

			If arrColVal(C_TrnsSLCd) <> "" then
				IG1_import_group(LngRow,C_IG1_I2_trns_sl_cd)	= arrColVal(C_TrnsSLCd)
			Else	
				IG1_import_group(LngRow,C_IG1_I2_trns_sl_cd)	= Request("txtSLCd1")
			End if

			if arrColVal(C_TrnsLotSubNo) = "" then
			   IntTrnsLotSubNo = 0
			else
			   IntTrnsLotSubNo = arrColVal(C_TrnsLotSubNo)
			End if
			
			If arrColVal(C_TrnsTrackingNo) <> "" then	
				IG1_import_group(LngRow,C_IG1_I2_trns_tracking_no)	= arrColVal(C_TrnsTrackingNo)
			Else	
				IG1_import_group(LngRow,C_IG1_I2_trns_tracking_no)	= arrColVal(C_TrackingNo)
			End if
			IG1_import_group(LngRow,C_IG1_I2_entry_qty_after)	= arrColVal(C_EntryQtyAfter)   '20080313::hanc::이동후수량

				IG1_import_group(LngRow,C_IG1_I2_item_status)	= "G"

			IG1_import_group(LngRow,C_IG1_I1_sl_cd)           = Request("txtSLCd1")			
			IG1_import_group(LngRow,C_IG1_I4_item_cd)		  = arrColVal(C_ItemCd)
			IG1_import_group(LngRow,C_IG1_I2_lot_no)		  = arrColVal(C_LotNo)
			IG1_import_group(LngRow,C_IG1_I2_lot_sub_no)	  = CInt(IntLotSubNo)
			IG1_import_group(LngRow,C_IG1_I2_entry_qty)	      = UNIConvNum(arrColVal(C_EntryQty), 0)
			IG1_import_group(LngRow,C_IG1_I2_entry_unit)	  = arrColVal(C_EntryUnit)										
    		IG1_import_group(LngRow,C_IG1_I3_plant_cd)	      = Request("txtPlantCd1")
			IG1_import_group(LngRow,C_IG1_I2_trns_lot_sub_no) = CInt(IntTrnsLotSubNo)
			IG1_import_group(LngRow,C_IG1_I2_trns_lot_no)	  = arrColVal(C_TrnsLotNo)
			IG1_import_group(LngRow,C_IG1_I5_cost_cd)		  = Request("txtCostCd2")


		case "D"
			dIG1_import_group(LngRow,D_IG1_sl_cd)		= Request("txtSLCd1")
			dIG1_import_group(LngRow,D_IG1_seq_no) 		= CInt(arrColVal(D_SeqNo))
			dIG1_import_group(LngRow,D_IG1_item_cd)		= arrColVal(D_ItemCd)	
			dIG1_import_group(LngRow,D_IG1_plant_cd)	= Request("txtPlantCd1")


		End Select
	   Next
	End If

	If CheckSYSTEMError(Err,True) = True Then
	   Response.End
    End If

	Set iPI0C180 = Server.CreateObject("PI0C180_KO441.cIStockTransfer")     '20080313::HANC::공장간이동
		
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript> " & vbCr   
		Response.Write " Parent.RemovedivTextArea	"	& vbCr
		Response.Write "</Script>	"	& vbCr
	   Response.End
    End If
	    
	Select Case strStatus	
	Case "C"
        

		Call iPI0C180.I_CREATE_STOCK_TRANSFER(gStrGlobalCollection, _
												, _
												I2_good_mvmt_workset, _
												IG1_import_group, _
												E1_good_mvmt_workset, _
												, _
												iErrorPosition)
		

    Case "D"
            
		Call iPI0C180.I_DELETE_STOCK_TRANSFER(gStrGlobalCollection, _
											DI2_good_mvmt_workset, _
											dIG1_import_group) 
            
    End select
        
    If CheckSYSTEMError(Err,True) = True  Then
       Set iPI0C180 = Nothing

       If iErrorPosition(0) <> 0 Then
			Response.Write "<Script Language=VBScript>" & vbCrLF
			Response.Write " Parent.RemovedivTextArea	"	& vbCr
			Response.Write "Call parent.SheetFocus(" & iErrorPosition(0) & ", 1)" & vbCrLF
			Response.Write "</Script>" & vbCrLF
       End If
       Response.End
    End If

	Set iPI0C180 = Nothing                         

	If	strStatus  = "C" Then		
       Response.Write "<Script Language=vbscript> " & vbCrlf
	   Response.Write "With parent.frm1 " & vbCrlf
       Response.Write ".txtDocumentNo1.Value 	= """ & ConvSPChars(E1_good_mvmt_workset(C_E1_item_document_no)) & """" & vbCrlf
       Response.Write ".txtYear.Text 		    = """ & E1_good_mvmt_workset(C_E1_year) & """" & vbCrlf
       Response.Write ".txtDocumentNo2.Value	= """ & ConvSPChars(E1_good_mvmt_workset(C_E1_item_document_no)) & """" & vbCrlf
       Response.Write "End With" & vbCrlf
       Response.Write "</Script>" & vbCrlf
	End If 

    Response.Write "<Script Language=vbscript> " & vbCrlf
	Response.Write " Parent.RemovedivTextArea	"	& vbCr
	Response.Write "parent.DbSaveOk" & vbCrlf
    Response.Write "</Script>" & vbCrlf

    Response.End
%>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        