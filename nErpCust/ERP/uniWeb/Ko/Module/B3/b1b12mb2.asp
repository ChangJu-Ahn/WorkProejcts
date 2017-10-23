<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrnumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b12mb2.asp	
'*  4. Program Name         : Lot Control Entry
'*  5. Program Desc         :
'*  6. Component List       : +PB3G112.cBLkUpLotCtlSvr.B_MANAGE_LOT_CONTROL
'*  7. Modified date(First) : 2000/05/3
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : Lee Hwa Jung
'* 11. Comment              :
'**********************************************************************************************
												'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 
    Err.Clear																		'☜: Protect system from crashing


    '[CONVERSION INFORMATION]  View Name : import b_lot_control
     Const P042_I3_lot_gen_mthd = 0
     Const P042_I3_last_lot_no = 1
     Const P042_I3_lot_prefix = 2
     Const P042_I3_increment = 3
     Const P042_I3_unit_of_perd = 4
     Const P042_I3_effective_flg = 5
     Const P042_I3_effective_perd = 6
     Const P042_I3_valid_from_dt = 7
     Const P042_I3_valid_to_dt = 8

	Call LoadBasisGlobalInf() 
	Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")
    Call HideStatusWnd                                                     '☜: Hide Processing message

	Dim pPB3S113
	Dim iCommandSent																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
    Dim I3_b_lot_control
    Dim I1_b_item_cd
    Dim I2_b_plant_cd
    
    If Request("txtPlantCd") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)					 '⊙: 에러메세지는 DB화 한다.           
		Response.End 
	End If
    
    If Request("txtItemCd1") = "" Then												'⊙: 저장을 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("900010", vbOKOnly, "", "", I_MKSCRIPT)						   '⊙: 에러메세지는 DB화 한다.           
		Response.End 
	End If
	
	
	Redim I3_b_lot_control(P042_I3_valid_to_dt)
    '-----------------------
    'Data manipulate area
    '-----------------------
	I2_b_plant_cd									= UCase(Request("txtPlantCd"))
	I1_b_item_cd									= UCase(Request("txtItemCd1"))
	I3_b_lot_control(P042_I3_lot_gen_mthd)			= UCase(Request("cboLotType"))
	I3_b_lot_control(P042_I3_last_lot_no)			= UniConvNum(Request("txtNewLotNo"),0)
	I3_b_lot_control(P042_I3_lot_prefix)			= Trim(Request("txtLotStartChar"))
	I3_b_lot_control(P042_I3_increment)				= UniCInt(Request("txtLotInc"),0)
	I3_b_lot_control(P042_I3_effective_flg)			= UCase(Request("rdoValidPerdFlg"))
	I3_b_lot_control(P042_I3_effective_perd)		= UniConvNum(Request("txtValidPerd"),0)
	
	If Len(Trim(Request("txtValidFromDt"))) Then
		If UniConvDate(Request("txtValidFromDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtValidFromDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			I3_b_lot_control(P042_I3_valid_from_dt)	= UniConvDate(Request("txtValidFromDt"))
		End If
	End If
	
	If Len(Trim(Request("txtValidToDt"))) Then
		If UniConvDate(Request("txtValidToDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtValidToDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			I3_b_lot_control(P042_I3_valid_to_dt)		= UniConvDate(Request("txtValidToDt"))
		End If
	End If

    If CInt(Request("txtFlgMode")) = OPMD_CMODE Then				'☜: 저장시 Create/Update 판별 
		iCommandSent = "CREATE"
    ElseIf CInt(Request("txtFlgMode")) = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If
    
	Set pPB3S113 = Server.CreateObject("PB3S113.cBMngLotCtl")
	
	If CheckSYSTEMError(Err, True) = True Then
		Response.End 
	End if
	
	Call pPB3S113.B_MANAGE_LOT_CONTROL (gStrGlobalCollection, I1_b_item_cd, I2_b_plant_cd, _
			I3_b_lot_control, iCommandSent)
			
	If CheckSYSTEMError(Err, True) = True Then
		Set pPB3S113 = Nothing															'☜: Unload Component
		Response.End
	End If

	Set pPB3S113 = Nothing															'☜: Unload Component

	'-----------------------
	'Result data display area
	'----------------------- 
%>
<Script Language=vbscript>
	With parent																			
		.DbSaveOk
	End With
</Script>
<%					

	Response.End																				'☜: Process End

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
%>
<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>

</SCRIPT>