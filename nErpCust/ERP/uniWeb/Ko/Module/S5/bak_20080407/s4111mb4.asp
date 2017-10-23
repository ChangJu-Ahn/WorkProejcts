<% Option Explicit %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5114MA2
'*  4. Program Name         : 매출채권일괄확정 
'*  5. Program Desc         :
'*  6. Comproxy List        : S51115BatchArProcessSvr
'*  7. Modified date(First) : 2000/04/20
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Cho song hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/20 : 3rd 화면 layout & ASP Coding
'*                            -2000/08/11 : 4th 화면 layout
'*                            -2001/12/19 : Date 표준적용 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

On Error Resume Next									

Call HideStatusWnd

Dim iStrMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim iObjPS5G137
Dim iArrHdrInfo
Dim pvCB			

iStrMode = Request("txtMode")												'☜ : 현재 상태를 받음 

Select Case iStrMode

Case CStr(UID_M0001)														'☜: 현재 조회/Prev/Next 요청을 받음 
	
	Dim iStrNextKey							' 다음 값 
	Dim iArrNextKey
	Dim lgStrPrevKey						' 이전 값 
	Dim iLngLastRow							' 현재 그리드의 최대Row

	Dim iLngRow
	Dim iLngSheetMaxRows
	Dim iArrCols
	Dim iArrRows

	Dim	iObjPS5G138

    Dim l1_query_con
    
    const S412_l1_plant			=	0
    const S412_l1_req_date_from	=	1
    const S412_l1_req_date_to	=	2
    const S412_l1_dn_type		=	3
    const S412_l1_ship_to_party	=	4
    const S412_l1_sales_grp		=	5
    Redim l1_query_con(S412_l1_sales_grp)

	Dim l2_next_key		
	const S412_l2_so_no			=	0
	const S412_l2_so_seq		=	1
	const S412_l2_so_schd_no	=	2
	Redim l2_next_key(S412_l2_so_schd_no)
	
	Dim E1_get_name		
	const S412_E1_plant_nm			=	0
	const S412_E1_dn_type_nm		=	1
	const S412_E1_ship_to_party_nm	=	2
	const S412_E1_sales_grp_nm		=	3
	
	Dim EG1_exp_grp
	Const S412_EG1_promise_dt = 0
    Const S412_EG1_ship_to_party = 1
    Const S412_EG1_bp_nm = 2
    Const S412_EG1_item_cd = 3
    Const S412_EG1_item_nm = 4
    Const S412_EG1_remain_qty = 5
    Const S412_EG1_bonus_remain_qty = 6
    Const S412_EG1_so_unit = 7
    Const S412_EG1_gi_qty = 8
    Const S412_EG1_gi_bonus_qty = 9
    Const S412_EG1_plant_cd = 10
    Const S412_EG1_plant_nm = 11
    Const S412_EG1_sl_cd = 12
    Const S412_EG1_sl_nm = 13
    Const S412_EG1_on_hand_qty = 14
    Const S412_EG1_su_on_hand_qty = 15
    Const S412_EG1_basic_unit = 16
    Const S412_EG1_so_no = 17
    Const S412_EG1_so_seq = 18
    Const S412_EG1_so_schd_no = 19
    Const S412_EG1_tracking_no = 20
    Const S412_EG1_spec = 21
    Const S412_EG1_dn_type = 22
    Const S412_EG1_dn_type_nm = 23
    Const S412_EG1_so_type = 24
    Const S412_EG1_sales_grp = 25
    Const S412_EG1_remark = 26

	Dim C_SHEETMAXROWS_D				' 한번에 Query할 Row수 

	If Request("txtBatchQuery") = "Y" Then
		C_SHEETMAXROWS_D = -1			' 조회조건에 해당되는 모든 Row를 반환한다.
	Else
		C_SHEETMAXROWS_D = 100
	End If
	'---------------------------------------------
    'next key값을 넘겨준다.
    '---------------------------------------------
	lgStrPrevKey = Trim(Request("lgStrPrevKey"))
	If lgStrPrevKey <> "" Then	
		iArrNextKey = Split(lgStrPrevKey, gColSep)		
		l2_next_key(S412_l2_so_no) = Trim(iArrNextKey(0))		
		l2_next_key(S412_l2_so_seq) = Trim(iArrNextKey(1))
		l2_next_key(S412_l2_so_schd_no) = Trim(iArrNextKey(2))
	Else
		l2_next_key(S412_l2_so_no) = ""
		l2_next_key(S412_l2_so_seq) = 0
		l2_next_key(S412_l2_so_schd_no) = 0		
	End if	    
		    
    '---------------------------------------------
    'Data manipulate  area(import view match)
    '---------------------------------------------
	l1_query_con(S412_l1_plant)				= Trim(Request("txtConPlant"))
	l1_query_con(S412_l1_req_date_from)		= UNIConvDate(Request("txtConReqDateFrom"))
	l1_query_con(S412_l1_req_date_to)		= UNIConvDate(Request("txtConReqDateTo"))
	l1_query_con(S412_l1_dn_type)			= Trim(Request("txtConDnType"))
	l1_query_con(S412_l1_ship_to_party)		= Trim(Request("txtConShipToParty"))
	l1_query_con(S412_l1_sales_grp)			= Trim(Request("txtConSalesGrp"))
	    
	Set iObjPS5G138 = Server.CreateObject("PS5G138.cSListSchdForGiSvr2")

	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Call parent.SetToolbar(""11100000000011"")" & vbCr
		Response.Write "parent.frm1.txtConPlant.focus" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'☜: Process End
       Response.End       
    End If
  
    Call iObjPS5G138.ListRows2(gStrGlobalCollection, C_SHEETMAXROWS_D, l1_query_con, l2_next_key, _
							E1_get_name, EG1_exp_grp)

	If CheckSYSTEMError(Err,True) = True Then
		Set iObjPS5G138 = Nothing
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "Call parent.SetToolbar(""11100000000011"")" & vbCr
		Response.Write "parent.frm1.txtConPlant.focus" & vbCr
		Response.Write "</Script>" & vbCr
		Response.End																				'☜: Process End
	End If
	
	Set iObjPS5G138 = Nothing

	' Client(MA)의 현재 조회된 마직막 Row
	iLngLastRow = CLng(Request("txtLastRow")) + 1
	
	' Set Next key
	If C_SHEETMAXROWS_D > 0 And Ubound(EG1_exp_grp,2) = C_SHEETMAXROWS_D Then
		'수주번호, 수주순번, 납품순번 
		iStrNextKey = EG1_exp_grp(S412_EG1_so_no, C_SHEETMAXROWS_D) & gColSep & EG1_exp_grp(S412_EG1_so_seq, C_SHEETMAXROWS_D) & gColSep & EG1_exp_grp(S412_EG1_so_schd_no, C_SHEETMAXROWS_D)
		iLngSheetMaxRows  = C_SHEETMAXROWS_D - 1
	Else
		iStrNextKey = ""
		iLngSheetMaxRows = Ubound(EG1_exp_grp,2)
	End If

	ReDim iArrCols(34)						' Column 수 
	Redim iArrRows(iLngSheetMaxRows)		' 조회된 Row 수만큼 배열 재정의 

	iArrCols(0) = ""
   	iArrCols(15) = ""						' 창고 Popup
		
   	For iLngRow = 0 To iLngSheetMaxRows
   		iArrCols(1) = "0"
   		iArrCols(2) = UNIDateClientFormat(EG1_exp_grp(S412_EG1_promise_dt, iLngRow))							' 출고예정일 
   		iArrCols(3) = ConvSPChars(EG1_exp_grp(S412_EG1_ship_to_party, iLngRow))									' 납품처 
   		iArrCols(4) = ConvSPChars(EG1_exp_grp(S412_EG1_bp_nm, iLngRow)) 										' 납품처명 
   		iArrCols(5) = ConvSPChars(EG1_exp_grp(S412_EG1_item_cd, iLngRow)) 										' 품목 
   		iArrCols(6) = ConvSPChars(EG1_exp_grp(S412_EG1_item_nm, iLngRow)) 										' 품목명 
   		iArrCols(7) = UNINumClientFormat(EG1_exp_grp(S412_EG1_remain_qty, iLngRow), ggQty.DecPoint, 0)			' 잔량 
   		iArrCols(8) = UNINumClientFormat(EG1_exp_grp(S412_EG1_bonus_remain_qty, iLngRow), ggQty.DecPoint, 0)	' 덤잔량 
   		iArrCols(9) = ConvSPChars(EG1_exp_grp(S412_EG1_so_unit, iLngRow)) 										' 단위 
   		iArrCols(10) = UNINumClientFormat(EG1_exp_grp(S412_EG1_gi_qty, iLngRow), ggQty.DecPoint, 0)				' 출고가능량 
   		iArrCols(11) = UNINumClientFormat(EG1_exp_grp(S412_EG1_gi_bonus_qty, iLngRow), ggQty.DecPoint, 0)		' 덤출고가능량 
   		iArrCols(12) = ConvSPChars(EG1_exp_grp(S412_EG1_plant_cd, iLngRow)) 									' 공장 
   		iArrCols(13) = ConvSPChars(EG1_exp_grp(S412_EG1_plant_nm, iLngRow)) 									' 공장명 
   		iArrCols(14) = ConvSPChars(EG1_exp_grp(S412_EG1_sl_cd, iLngRow)) 										' 창고 
   		iArrCols(16) = ConvSPChars(EG1_exp_grp(S412_EG1_sl_nm, iLngRow)) 										' 창고명 
   		iArrCols(17) = UNINumClientFormat(EG1_exp_grp(S412_EG1_su_on_hand_qty, iLngRow), ggQty.DecPoint, 0)		' 수주단위재고량 
   		iArrCols(18) = UNINumClientFormat(EG1_exp_grp(S412_EG1_on_hand_qty, iLngRow), ggQty.DecPoint, 0)		' 현재고량 
   		iArrCols(19) = ConvSPChars(EG1_exp_grp(S412_EG1_basic_unit, iLngRow)) 									' 재고단위 
   		iArrCols(20) = ConvSPChars(EG1_exp_grp(S412_EG1_so_no, iLngRow)) 										' 수주번호 
   		iArrCols(21) = ConvSPChars(EG1_exp_grp(S412_EG1_so_seq, iLngRow)) 										' 수주순번 
   		iArrCols(22) = ConvSPChars(EG1_exp_grp(S412_EG1_so_schd_no, iLngRow)) 									' 납품순번 
   		iArrCols(23) = ConvSPChars(EG1_exp_grp(S412_EG1_tracking_no, iLngRow)) 									' Tracking No
   		iArrCols(24) = ConvSPChars(EG1_exp_grp(S412_EG1_spec, iLngRow))	 										' 규격 
   		iArrCols(25) = ConvSPChars(EG1_exp_grp(S412_EG1_dn_type, iLngRow)) 										' 출하형태 
   		iArrCols(26) = ConvSPChars(EG1_exp_grp(S412_EG1_dn_type_nm, iLngRow)) 									' 출하형태명 
   		iArrCols(27) = ConvSPChars(EG1_exp_grp(S412_EG1_remark, iLngRow)) 										' 비고 
   		iArrCols(28) = ConvSPChars(EG1_exp_grp(S412_EG1_so_type, iLngRow)) 										' 수주유형 
   		iArrCols(29) = ConvSPChars(EG1_exp_grp(S412_EG1_sales_grp, iLngRow)) 									' 영업그룹 
   		iArrCols(30) = iArrCols(14)			' 창고 
   		iArrCols(31) = iArrCols(16)			' 창고명 
   		iArrCols(32) = iArrCols(10)			' 출고가능수량 
   		iArrCols(33) = iArrCols(11)			' 출고가능 덤수량 
   		iArrCols(34) = iLngLastRow + iLngRow 
   		
   		iArrRows(iLngRow) = Join(iArrCols, gColSep)
	Next
	
	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT> " & vbCr   
	Response.Write "With parent " & vbCr   
	
	' 조건절 명 Display(처음 조회시만 조건절의 명을 Display한다)
	If lgStrPrevKey = "" Then	
		Response.Write ".frm1.txtConPlantNm.Value			= """ & ConvSPChars(E1_get_name(S412_E1_plant_nm)) & """" & vbCr
		Response.Write ".frm1.txtConDnTypeNm.Value			= """ & ConvSPChars(E1_get_name(S412_E1_dn_type_nm)) & """" & vbCr
		Response.Write ".frm1.txtConShipToPartyNm.Value		= """ & ConvSPChars(E1_get_name(S412_E1_ship_to_party_nm)) & """" & vbCr
		Response.Write ".frm1.txtConSalesGrpNm.Value		= """ & ConvSPChars(E1_get_name(S412_E1_sales_grp_nm)) & """" & vbCr
	End If
	
	' 내역 Display
    Response.Write ".ggoSpread.Source = .frm1.vspdData " & vbCr
    Response.Write ".frm1.vspdData.Redraw = False  "      & vbCr      
    Response.Write ".ggoSpread.SSShowDataByClip   """ & Join(iArrRows, gColSep & gRowSep) & gColSep & gRowSep & """ ,""F""" & vbCr
    Response.Write ".lgStrPrevKey = """ & ConvSPChars(iStrNextKey) & """" & vbCr  
    Response.Write ".DbQueryOk" & vbCr   
	Response.Write ".frm1.vspdData.Redraw = True  "       & vbCr
	
	' 다음 Query를 위한 조회조건 설정 
	If iStrNextKey <> "" Then
		Response.Write ".frm1.txtHConPlant.value		= """ & Request("txtConPlant") & """" & vbCr
		Response.Write ".frm1.txtHConReqDateFrom.value	= """ & Request("txtConReqDateFrom") & """" & vbCr
		Response.Write ".frm1.txtHConReqDateTo.value	= """ & Request("txtConReqDateTo") & """" & vbCr
		Response.Write ".frm1.txtHConDnType.value		= """ & Request("txtConDnType") & """" & vbCr
		Response.Write ".frm1.txtHConShipToParty.value	= """ & Request("txtConShipToParty") & """" & vbCr
		Response.Write ".frm1.txtHConSalesGrp.value		= """ & Request("txtConSalesGrp") & """" & vbCr
	End If
	Response.Write "End With " & vbCr   
	Response.Write "</SCRIPT> " & vbCr      	

	Response.End 
    
Case CStr(UID_M0002)						'☜: 저장 요청을 받음 

	Dim iArrDnNo						' 추가된 출고번호 배열 (Output)
	Dim iErrorPosition
	Dim iStrFirstDnNo, iStrLastDnNo		' 추가된 출고번호 
	Dim iIntDnNoCount					' 추가된 출고정보 수 
    Dim itxtSpreadIns, itxtSpreadArr
    Dim iIntIndex, iCCount
    
    Redim iArrHdrInfo(4)
    
    iArrHdrInfo(0) = UNIConvDate(Request("txtActualGIDt"))	' 실제 출고일 
    iArrHdrInfo(1) = UCase(Trim(Request("txtTransMeth")))	' 운송방법 
    iArrHdrInfo(2) = Trim(Request("txtHArFlag"))			' 매출생성여부 
    iArrHdrInfo(3) = Trim(Request("txtHVatFlag"))			' 세금계산서 생성여부 
    iArrHdrInfo(4) = UCase(Trim(Request("txtInvMgr")))		' 재고담당자 - 2003.08.26(Hwang Seongbae)
	
	pvCB = "F" 	   
	iIntDnNoCount = 0
	
    iCCount = Request.Form("txtCSpread").Count

    ' 추가 
    ReDim itxtSpreadArr(iCCount)
    For iIntIndex = 1 To iCCount
        itxtSpreadArr(iIntIndex) = Request.Form("txtCSpread")(iIntIndex)
    Next
    itxtSpreadIns = Join(itxtSpreadArr,"")
	
	Set iObjPS5G137 = Server.CreateObject("PS5G137.cSCollectivelyGiSvr2")

	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write "Call parent.RemovedivTextArea " & vbCr   
		Response.Write "</Script> "																				         & vbCr          
       Response.End       
    End If

    Call iObjPS5G137.S_MAINT_COLLECTIVELY_GI_SVR2(pvCB, gStrGlobalCollection, iArrHdrInfo, itxtSpreadIns, _
												  iArrDnNo, iErrorPosition)
    Set iObjPS5G137 = Nothing
	
	If Trim(iErrorPosition) <> "" Then
		If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") Then		
			Set iObjPS5G137 = Nothing
			Response.Write "<Script language=vbs> " & vbCr   
			Response.Write " Call parent.RemovedivTextArea " & vbCr   
			Response.Write " Call Parent.SubSetErrPos(" & iErrorPosition & ")" & vbCr
			Response.Write "</Script> "																				         & vbCr          
			Response.End
		End If
	Else
		If CheckSYSTEMError(Err,True) = True Then
			Set iObjPS5G137 = Nothing
			Response.Write "<Script language=vbs> " & vbCr   
			Response.Write " Call parent.RemovedivTextArea " & vbCr
			Response.Write " Call parent.frm1.txtConPlant.focus " & vbCr
			Response.Write "</Script> "																				         & vbCr          
			Response.End
		End If
	End If
	
	iIntDnNoCount = UBound(iArrDnNo)
	iStrFirstDnNo = iArrDnNo(0)
	iStrLastDnNo = iArrDnNo(iIntDnNoCount)
		
	iIntDnNoCount = iIntDnNoCount + 1		' 추가된 출고정보 수 

	Call DisplayMsgBox("204262", vbOKOnly, iStrFirstDnNo & "~" & iStrLastDnNo & " (" & iIntDnNoCount & ")", "", I_MKSCRIPT)

	Response.Write "<Script language=vbs> " & vbCr   
	Response.Write "Call parent.DbSaveOk " & vbCr   
	Response.Write "</Script> "	& vbCr          

End Select
%>

