<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  b1b12mb1.asp
'*  4. Program Name         :  Lot Control 조회 
'*  5. Program Desc         :
'*  6. Component List       : +PB3G112.cBLkUpLotCtlSvr.B_LOOK_UP_LOT_CONTROL_Svr
'*  7. Modified date(First) : 2000/05/03
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************

													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 
Err.Clear


'[CONVERSION INFORMATION]  View Name : export b_plant
Const P037_E1_plant_cd = 0
Const P037_E1_plant_nm = 1

'[CONVERSION INFORMATION]  EXPORTS View 상수 

'[CONVERSION INFORMATION]  View Name : export b_item
Const P037_E2_item_cd = 0
Const P037_E2_item_nm = 1
Const P037_E2_formal_nm = 2
Const P037_E2_spec = 3

'[CONVERSION INFORMATION]  EXPORTS View 상수 

'[CONVERSION INFORMATION]  View Name : export b_lot_control
Const P037_E3_lot_gen_mthd = 0
Const P037_E3_last_lot_no = 1
Const P037_E3_lot_prefix = 2
Const P037_E3_increment = 3
Const P037_E3_unit_of_perd = 4
Const P037_E3_effective_flg = 5
Const P037_E3_effective_perd = 6
Const P037_E3_valid_from_dt = 7
Const P037_E3_valid_to_dt = 8

Call LoadBasisGlobalInf() 
Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")
Call HideStatusWnd                                                     '☜: Hide Processing message

Dim strMode																	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim pPB3G112
Dim I2_b_plant_cd
Dim I3_b_item_cd
Dim I1_select_char
Dim E1_b_plant
Dim E2_b_item
Dim E3_b_lot_control
Dim iStatusCodeOfPrevNext
'strMode = Request("txtMode")												'☜ : 현재 상태를 받음 

    Err.Clear                                                               '☜: Protect system from crashing

    If Request("txtPlantCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	End If
    
    If Request("txtItemCd") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)
		Response.End 
	End If
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I2_b_plant_cd	 = UCase(Trim(Request("txtPlantCd")))
    I3_b_item_cd  = UCase(Trim(Request("txtItemCd")))
    I1_select_char = Request("PrevNextFlg")
    
  
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    Set pPB3G112 = Server.CreateObject("PB3G112.cBLkUpLotCtlSvr")    
	
    If CheckSYSTEMError(Err, True)= True Then
		Response.End																		'☜: Process End
	End If

    Call pPB3G112.B_LOOK_UP_LOT_CONTROL_Svr (gStrGlobalCollection, I1_select_char, I2_b_plant_cd, _
				I3_b_item_cd, E1_b_plant, E2_b_item, E3_b_lot_control, iStatusCodeOfPrevNext)
	
	If CheckSYSTEMError(Err, True) = True Then
		Set pPB3G112 = Nothing															'☜: Unload Component
		Response.Write "<Script Language=VBScript>" & vbCrLf
		Response.Write "With parent.frm1" & vbCrLf
		Response.Write ".txtPlantNm.value = """ & ConvSPChars(E1_b_plant(P037_E1_plant_nm)) & """" & vbCrLf
		Response.Write ".txtItemNm.value = """ & ConvSPChars(E2_b_item(P037_E2_item_nm)) & """" & vbCrLf
		Response.Write ".txtItemCd.Focus()" & vbCrLf
		Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Response.End
	End If

    Set pPB3G112= Nothing

	If Trim(iStatusCodeOfPrevNext) = "900011" Or  Trim(iStatusCodeOfPrevNext) = "900012" Then
		Call DisplayMsgBox(iStatusCodeOfPrevNext, vbOkOnly, "", "", I_MKSCRIPT)
	End If

	
Response.Write "<Script Language=vbscript> " & vbCr
Response.Write "	With parent.frm1 " & vbCr
		
Response.Write "		.txtPlantCd.value		= """ & ConvSPChars(UCase(E1_b_plant(P037_E1_plant_cd))) & """" & vbCr	 
Response.Write "		.txtPlantNm.value		= """ & ConvSPChars(E1_b_plant(P037_E1_plant_nm)) & """" & vbCr
Response.Write "		.txtItemCd.value		= """ & ConvSPChars(UCase(E2_b_item(P037_E2_item_cd))) & """" & vbCr
Response.Write "		.txtItemNm.value		= """ & ConvSPChars(E2_b_item(P037_E2_item_nm)) & """" & vbCr 
Response.Write "		.txtItemCd1.value		= """ & ConvSPChars(UCase(E2_b_item(P037_E2_item_cd))) & """" & vbCr
Response.Write "		.txtItemNm1.value		= """ & ConvSPChars(E2_b_item(P037_E2_item_nm)) & """" & vbCr
Response.Write "		.cboLotType.value		= """ & E3_b_lot_control(P037_E3_lot_gen_mthd) & """" & vbCr
Response.Write "		.txtNewLotNo.value		= """ & ConvSPChars(E3_b_lot_control(P037_E3_last_lot_no)) & """" & vbCr	   
Response.Write "		.txtLotStartChar.value	= """ & ConvSPChars(E3_b_lot_control(P037_E3_lot_prefix)) & """" & vbCr
Response.Write "		.txtLotInc.Text		= """ & UniConvNumDBToCompanyWithOutChange(E3_b_lot_control(P037_E3_increment), 0) & """" & vbCr
				
		If Trim(UCase(E3_b_lot_control(P037_E3_effective_flg))) = "Y" Then
			Response.Write "	.rdoValidPerdFlg(0).checked = True " & vbCr
			Response.Write "	parent.lgRdoOldVal1 = 1 " & vbCr
		Else
			Response.Write "	.rdoValidPerdFlg(1).checked = True " & vbCr
			Response.Write "	parent.lgRdoOldVal1 = 2 " & vbCr
		End If		

		
Response.Write "		.txtValidPerd.Text		= """ & UniConvNumDBToCompanyWithOutChange(E3_b_lot_control(P037_E3_effective_perd), 0) & """" & vbCr
Response.Write "		.txtValidFromDt.text	= """ & UniDateClientFormat(E3_b_lot_control(P037_E3_valid_from_dt)) & """" & vbCr
Response.Write "		.txtValidToDt.text		= """ & UniDateClientFormat(E3_b_lot_control(P037_E3_valid_to_dt)) & """" & vbCr
		 		
		
Response.Write "		parent.lgNextNo = """"" & vbCr		' 다음 키 값 넘겨줌 
Response.Write "		parent.lgPrevNo = """"" & vbCr		' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 
		
Response.Write "		parent.DbQueryOk " & vbCr																'☜: 조화가 성공 
Response.Write "	End With			 " & vbCr

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================

Response.Write "</Script> " & vbCr

Response.End																	'☜: Process End
%>
