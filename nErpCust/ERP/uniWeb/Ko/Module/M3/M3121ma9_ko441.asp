<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
Response.Expires = -1
%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        :
'*  3. Program ID           : m3121ma1
'*  4. Program Name         : 발주등록(통합)-테스트
'*  5. Program Desc         :
'*  6. Modified date(First) : 2004/11/10
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Byun Jee Hyun
'*  9. Modifier (Last)      :
'* 10. Comment              :
'* 11. History              :
'*
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

Const TAB1 = 1
Const TAB2 = 2
Const TAB3 = 3

'==========================================================================================================
Dim C_PlantCd
Dim C_Popup1
Dim C_PlantNm
Dim C_itemCd
Dim C_Popup2
Dim C_itemNm
Dim C_SpplSpec
Dim C_OrderQty
Dim C_OrderUnit
Dim C_Popup3
Dim C_Cost
Dim C_Check
Dim C_CostCon
Dim C_CostConCd
Dim C_OrderAmt
Dim C_NetAmt
Dim C_OrgNetAmt
Dim C_IOFlg
Dim C_IOFlgCd
Dim C_VatType
Dim C_Popup7
Dim C_VatNm
Dim C_VatRate
Dim C_VatAmt
Dim C_OrgVatAmt
Dim C_DlvyDT
Dim C_HSCd
Dim C_Popup5
Dim C_HSNm
Dim C_SLCd
Dim C_Popup6
Dim C_SLNm
Dim C_TrackingNo
Dim C_TrackingNoPop
Dim C_Lot_No
Dim C_Lot_Seq
Dim C_RetCd
Dim C_Popup8
Dim C_RetNm
Dim C_Over
Dim C_Under
Dim C_Bal_Qty
Dim C_Bal_Doc_Amt
Dim C_Bal_Loc_Amt
Dim C_ExRate
Dim C_SeqNo
Dim C_PrNo
Dim C_MvmtNo
Dim C_PoNo
Dim C_PoSeqNo
Dim C_MaintSeq
Dim C_SoNo
Dim C_SoSeqNo
Dim C_OrgNetAmt1
Dim C_reference
Dim C_Stateflg
Dim C_Remrk

Dim StrTime
Dim EndTime
Dim DifferTime


<!-- #Include file="../../inc/lgvariables.inc" -->
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Const BIZ_PGM_ID 					= "m3121mb1_ko441.asp" '>>air
Const BIZ_OnLine_ID 				= "m3111ab1.asp"
Const BIZ_PGM_JUMP_ID_PO_DTL 		= "M3112MA1"
Const BIZ_PGM_JUMP_ID_PUR_CHARGE	= "M6111MA2"
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim lgMpsFirmDate, lgLlcGivenDt
Dim gSelframeFlg
Dim lgIntFlgMode_Dtl
Dim cboOldVal
Dim IsOpenPop
Dim lblnWinEvent
Dim lgCboKeyPress
Dim lgOldIndex
Dim lgOldIndex2
Dim lgOpenFlag
Dim lgTabClickFlag
Dim arrCollectVatType
Dim StartDate, EndDate
Dim iDBSYSDate
Dim lgReqRefChk


iDBSYSDate = "<%=GetSvrDate%>"

	'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
	EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
    'StartDate = UniDateAdd("m", -1, iDBSYSDate,gServerDateFormat)    '☆: 초기화면에 뿌려지는 시작 날짜 -----
    'StartDate = UniConvDateAToB(StartDate,gServerDateFormat,gDateFormat)
'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함)
'#########################################################################################################

'========================================================================================
' Function Name : OnLineQueryOK
' Function Desc : fi
'========================================================================================
Function OnLineQueryOK()

	If Trim(frm1.txtSupplierCd.value) <> "" Then Call SupplierLookUp()
	'======================== 추후에 수정=======================
	if Trim(frm1.txtPotypeCd.Value) <> "" then Call ChangePotype()
	'======================== 추후에 수정=======================
End Function

'==========================================================================================
'   Event Name : SupplierLookUp
'   Event Desc :
'==========================================================================================
Function SupplierLookup()

    Err.Clear

    If CheckRunningBizProcess = True Then
		Exit Function
	End If

    Dim strVal

    if Trim(frm1.txtSupplierCd.Value) = "" then
    	Exit Function
    End if

	lgBlnFlgChgValue = true

    strVal = BIZ_PGM_ID & "?txtMode=" & "SupplierLookupAfterOnline"
    strVal = strVal & "&txtSupplierCd=" & Trim(frm1.txtSupplierCd.value)

    If LayerShowHide(1) = False Then Exit Function

	Call RunMyBizASP(MyBizASP, strVal)

End Function
'==========================================================================================
'   Event Name : ChangePotype
'   Event Desc : txtPotypeCd Chagne Event
'==========================================================================================
Sub ChangePotype()

	If gLookUpEnable = False Then
		Exit Sub
	End If

	Call PotypeRef()

End Sub

'==========================================================================================
'   Event Name : ChangeSupplier
'   Event Desc : txtSupplierCd Chagne Event
'==========================================================================================
 Sub ChangeSupplier()

	If gLookUpEnable = False Then
		Exit Sub
	End If

	Call SpplRef()
End Sub

'==========================================   PotypeRef()  ======================================
'	Name : PotypeRef()
'	Description :
'=========================================================================================================

 Sub PotypeRef()

    Err.Clear

    If CheckRunningBizProcess = True Then
		Exit Sub
	End If

    Dim strVal

    if Trim(frm1.txtPotypeCd.Value) = "" then
		Call DisplayMsgBox("205152", "X", "발주형태", "X")
		frm1.txtPotypeCd.focus
    	Exit Sub
    End if

	if lgIntFlgMode <> Parent.OPMD_UMODE Then
		lgBlnFlgChgValue = true
	end if

    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpPoType"
    strVal = strVal & "&txtPoTypeCd=" & Trim(frm1.txtPoTypeCd.value)
    strVal = strVal & "&txtTabClickFlag=" & lgTabClickFlag
	strVal = strVal & "&txtgSelframeFlg=" & gSelframeFlg
'	msgbox strVal
    If LayerShowHide(1) = False Then Exit Sub

	Call RunMyBizASP(MyBizASP, strVal)

End Sub

'==========================================   SpplRef()  ======================================
'	Name : SpplRef()
'	Description : It is Call at txtSupplier Change Event
'=========================================================================================================

 Sub SpplRef()

    Err.Clear

    Dim strVal

    if Trim(frm1.txtSupplierCd.Value) = "" then
    	Exit Sub
    End if

	lgBlnFlgChgValue = true

    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpSupplier"
    strVal = strVal & "&txtCurr=" & Parent.gCurrency
    strVal = strVal & "&txtSupplierCd=" & Trim(frm1.txtSupplierCd.value)
    strVal = strVal & "&txtGroupCd=" & Trim(frm1.txtGroupCd.value)
    strVal = strVal & "&lgPGCd=" & lgPGCd

    If LayerShowHide(1) = False Then Exit Sub

	Call RunMyBizASP(MyBizASP, strVal)

End Sub
'==========================================   Cfm()  ======================================
'	Name : Cfm()
'	Description : 확정버튼,확정취소버튼의 Event 합수 
'=========================================================================================================
 Sub Cfm()
    Dim IntRetCD

    Err.Clear

    if lgBlnFlgChgValue = True	then
		Call DisplayMsgBox("189217", "X", "X", "X")
		Exit sub
	End if

	if frm1.rdoRelease(0).checked = True then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		Call DbSave("Cfm")

	elseif frm1.rdoRelease(1).checked = True then

		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		Call DbSave("UnCfm")
	End if

End Sub

'-------------------------------------------------------------------
'		확정여부에 따라 Field의 속성을 Protect로 전환,복구 시키는 함수 
'--------------------------------------------------------------------

Function ChangeTag(Byval Changeflg)

	with frm1

		if Changeflg = true then

			'첫번째 Tab
			'ggoOper.SetReqAttr	.txtPoTypeCd, "Q"
			ggoOper.SetReqAttr	.txtPoDt, "Q"
			ggoOper.SetReqAttr	.txtGroupCd, "Q"
			ggoOper.SetReqAttr	.txtSupplierCd, "Q"
			ggoOper.SetReqAttr	.txtCurr, "Q"
			ggoOper.SetReqAttr	.txtXch, "Q"
			ggoOper.SetReqAttr	.cboXchop,"Q"
			ggoOper.SetReqAttr	.txtVatType, "Q"
			ggoOper.SetReqAttr	.txtPayTermCd, "Q"
			ggoOper.SetReqAttr	.txtPayDur, "Q"
			ggoOper.SetReqAttr	.txtPayTermstxt, "Q"
			ggoOper.SetReqAttr	.txtPayTypeCd, "Q"
			ggoOper.SetReqAttr	.txtSuppSalePrsn, "Q"
			ggoOper.SetReqAttr	.txtTel, "Q"
			ggoOper.SetReqAttr	.txtRemark, "Q"
			ggoOper.SetReqAttr	.rdoMergPurFlg1, "Q"
			ggoOper.SetReqAttr	.rdoMergPurFlg2, "Q"
			ggoOper.SetReqAttr  .rdoVatFlg1,"Q"
            ggoOper.SetReqAttr  .rdoVatFlg2,"Q"


			'두번째 Tab
			ggoOper.SetReqAttr	.txtOffDt, "Q"
			ggoOper.SetReqAttr	.txtDvryDt, "Q"
			ggoOper.SetReqAttr	.txtExpiryDt, "Q"
			ggoOper.SetReqAttr	.txtInvNo, "Q"
			ggoOper.SetReqAttr	.txtIncotermsCd, "Q"
			ggoOper.SetReqAttr	.txtTransCd, "Q"
			ggoOper.SetReqAttr	.txtBankCd, "Q"
			ggoOper.SetReqAttr	.txtDvryPlce, "Q"
			ggoOper.SetReqAttr	.txtApplicantCd, "Q"
			ggoOper.SetReqAttr	.txtManuCd, "Q"
			ggoOper.SetReqAttr	.txtAgentCd, "Q"
			ggoOper.SetReqAttr	.txtOrigin, "Q"
			ggoOper.SetReqAttr	.txtPackingCd, "Q"
			ggoOper.SetReqAttr	.txtInspectCd, "Q"
			ggoOper.SetReqAttr	.txtDisCity, "Q"
			ggoOper.SetReqAttr	.txtDisPort, "Q"
			ggoOper.SetReqAttr	.txtLoadPort, "Q"
			ggoOper.SetReqAttr	.txtShipment, "Q"



		else
			'첫번째 Tab
			ggoOper.SetReqAttr	.txtPoNo2, "D"
			ggoOper.SetReqAttr	.txtPoDt, "N"
			ggoOper.SetReqAttr	.txtGroupCd, "N"
			ggoOper.SetReqAttr	.txtSupplierCd, "N"
			ggoOper.SetReqAttr	.txtCurr, "N"
			ggoOper.SetReqAttr	.txtXch, "D"
			ggoOper.SetReqAttr	.txtVatType, "D"
			ggoOper.SetReqAttr	.txtPayTermCd, "N"
			ggoOper.SetReqAttr	.txtPayDur, "D"
			ggoOper.SetReqAttr	.txtPayTermstxt, "D"
			ggoOper.SetReqAttr	.txtPayTypeCd, "D"
			ggoOper.SetReqAttr	.txtSuppSalePrsn, "D"
			ggoOper.SetReqAttr	.txtTel, "D"
			ggoOper.SetReqAttr	.txtRemark, "D"
			ggoOper.SetReqAttr	.rdoMergPurFlg1, "D"
			ggoOper.SetReqAttr	.rdoMergPurFlg2, "D"

			if .hdnImportflg.value = "Y" then
			    ggoOper.SetReqAttr	.txtDvryDt, "N"
			    ggoOper.SetReqAttr	.txtOffDt, "N"
			    ggoOper.SetReqAttr	.txtApplicantCd, "N"
			    ggoOper.SetReqAttr	.txtIncotermsCd, "N"
			    ggoOper.SetReqAttr	.txtTransCd, "N"
			else
			    ggoOper.SetReqAttr	.txtDvryDt, "D"
			    ggoOper.SetReqAttr	.txtOffDt, "Q"
			    ggoOper.SetReqAttr	.txtApplicantCd, "Q"
			    ggoOper.SetReqAttr	.txtIncotermsCd, "Q"
			    ggoOper.SetReqAttr	.txtTransCd, "Q"
			end if

			'두번째 Tab
			'ggoOper.SetReqAttr	.txtOffDt, "N"
			'ggoOper.SetReqAttr	.txtDvryDt, "N"
			ggoOper.SetReqAttr	.txtExpiryDt, "D"
			ggoOper.SetReqAttr	.txtInvNo, "D"
			'ggoOper.SetReqAttr	.txtIncotermsCd, "N"
			'ggoOper.SetReqAttr	.txtTransCd, "N"
			ggoOper.SetReqAttr	.txtBankCd, "D"
			ggoOper.SetReqAttr	.txtDvryPlce, "D"
			'ggoOper.SetReqAttr	.txtApplicantCd, "N"
			ggoOper.SetReqAttr	.txtManuCd, "D"
			ggoOper.SetReqAttr	.txtAgentCd, "D"
			ggoOper.SetReqAttr	.txtOrigin, "D"
			ggoOper.SetReqAttr	.txtPackingCd, "D"
			ggoOper.SetReqAttr	.txtInspectCd, "D"
			ggoOper.SetReqAttr	.txtDisCity, "D"
			ggoOper.SetReqAttr	.txtDisPort, "D"
			ggoOper.SetReqAttr	.txtLoadPort, "D"
			ggoOper.SetReqAttr	.txtShipment, "D"


			if UCase(Trim(frm1.txtCurr.value)) = UCase(Parent.gCurrency) then

				Call ggoOper.SetReqAttr(frm1.txtXch,"Q")
				Call ggoOper.SetReqAttr(frm1.cboXchop,"Q")
			else
				Call ggoOper.SetReqAttr(frm1.txtXch,"D")
			end if

			If lgPGCd <> "" then 
				Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
			End If

		End if

	end with

End Function


'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------

Function CookiePage(Byval Kubun)

	Dim strTemp, arrVal
	Dim IntRetCD


	If Kubun = 0 Then

		strTemp = ReadCookie("PoNo")

		If strTemp = "" then Exit Function

		frm1.txtPoNo.value = strTemp

		WriteCookie "PoNo" , ""

		Call dbQuery()

	elseIf Kubun = 1 Then

	    If lgIntFlgMode <> Parent.OPMD_UMODE Then
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If

	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

		WriteCookie "PoNo" , frm1.txtPoNo.value

		Call PgmJump(BIZ_PGM_JUMP_ID_PO_DTL)

	elseIf Kubun = 2 Then

	    If lgIntFlgMode <> Parent.OPMD_UMODE Then
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If

	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

	    WriteCookie "Process_Step" , "PO"
		WriteCookie "Po_No" , Trim(frm1.txtPoNo.value)
		WriteCookie "Pur_Grp", Trim(frm1.txtGroupCd.Value)
		WriteCookie "Po_Cur", Trim(frm1.txtCurr.Value)
		WriteCookie "Po_Xch", Trim(frm1.txtXch.Value)

		Call PgmJump(BIZ_PGM_JUMP_ID_PUR_CHARGE)

	End IF

End Function
'------------------------------------------------------------------------------------------
'Radio에서 Click을 할 경우 flag를 Setting
'------------------------------------------------------------------------------------------
Sub Setchangeflg()
	lgBlnFlgChgValue = True
End Sub
'------------------------------------------------------------------------------------------
'사용자가 Radio Button을 Click할 때 마다 숨겨진 hdnRelease를 Setting
'------------------------------------------------------------------------------------------
Sub Changeflg()

	if frm1.rdoRelease(0).checked = true then
		frm1.hdnRelease.value= "N"
	else
		frm1.hdnRelease.value= "Y"
	end if

End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_PlantCd 		= 1 
	C_Popup1		= 2 
	C_PlantNm 		= 3 
	C_itemCd 		= 4 
	C_Popup2 		= 5 
	C_itemNm 		= 6 
	C_SpplSpec      = 7 
	C_OrderQty		= 8 
	C_OrderUnit		= 9 
	C_Popup3		= 10
	C_Cost			= 11
	C_Check			= 12
	C_CostCon		= 13
	C_CostConCd		= 14
	C_OrderAmt		= 15
	C_NetAmt        = 16
	C_OrgNetAmt     = 17
	C_IOFlg		    = 18
	C_IOFlgCd	    = 19
	C_VatType       = 21
	C_Popup7        = 21
	C_VatNm         = 22
	C_VatRate       = 23
	C_VatAmt        = 24
	C_OrgVatAmt		= 25
	C_DlvyDT		= 26
	C_HSCd			= 27
	C_Popup5		= 28
	C_HSNm			= 29
	C_SLCd			= 30
	C_Popup6		= 31
	C_SLNm			= 32
	C_TrackingNo	= 33
	C_TrackingNoPop	= 34
	C_Lot_No        = 35
	C_Lot_Seq       = 36
	C_RetCd         = 37
	C_Popup8        = 38
	C_RetNm         = 39
	C_Over			= 40
	C_Under			= 41
	C_Bal_Qty		= 42
	C_Bal_Doc_Amt	= 43
	C_Bal_Loc_Amt	= 44
	C_ExRate		= 45
	C_SeqNo 		= 46
	C_PrNo			= 47
	C_MvmtNo		= 48
	C_PoNo			= 49
	C_PoSeqNo		= 50
	C_MaintSeq		= 51
	C_SoNo			= 52
	C_SoSeqNo		= 53
	C_OrgNetAmt1    = 54
	C_reference     = 55
	C_Stateflg		= 56
	C_Remrk			= 57

End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables

	With frm1.vspdData

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021118",,Parent.gAllowDragDropSpread

	.ReDraw = false

    .MaxCols = C_Remrk+1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0

    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetEdit 	C_PlantCd, "공장", 7,,,4,2
    ggoSpread.SSSetButton 	C_Popup1
    ggoSpread.SSSetEdit 	C_PlantNm, "공장명", 20
    ggoSpread.SSSetEdit 	C_ItemCd, "품목", 18,,,18,2
    ggoSpread.SSSetButton 	C_Popup2
    ggoSpread.SSSetEdit 	C_ItemNm, "품목명", 20
    ggoSpread.SSSetEdit		C_SpplSpec, "품목규격", 20        '품목규격 추가 
    SetSpreadFloatLocal		C_OrderQty, "발주수량",15,1,3
    ggoSpread.SSSetEdit 	C_OrderUnit, "단위", 6,,,3,2
    ggoSpread.sssetButton 	C_Popup3
    SetSpreadFloatLocal		C_Cost, "단가",15,1,4
    ggoSpread.sssetButton	C_Check
    ggoSpread.SSSetCombo 	C_CostCon, "단가구분", 10,0,False
    ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
    ggoSpread.SSSetCombo 	C_CostConCd, "단가구분코드", 10,0,False
    ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
    SetSpreadFloatLocal		C_OrderAmt, "금액",15,1,2
    SetSpreadFloatLocal		C_NetAmt, "발주순금액",15,1,2
    SetSpreadFloatLocal		C_OrgNetAmt, "C_OrgNetAmt",15,1,2
    SetSpreadFloatLocal		C_OrgNetAmt1, "C_OrgNetAmt1",15,1,2
    ggoSpread.SSSetDate 	C_DlvyDt, "납기일", 10, 2, Parent.gDateFormat
    ggoSpread.SSSetEdit 	C_HSCd, "HS부호", 15,,,20,2
    ggoSpread.sssetButton 	C_Popup5
    ggoSpread.SSSetEdit 	C_HSNm, "HS명", 20
    ggoSpread.SSSetEdit 	C_SLCd, "창고", 10,,,7,2
    ggoSpread.SSSetButton 	C_Popup6
    ggoSpread.SSSetEdit 	C_SLNm, "창고명", 20
    ggoSpread.SSSetEdit 	C_TrackingNo, "Tracking No.",  15,,,25,2
    ggoSpread.SSSetButton 	C_TrackingNoPop
    ggoSpread.SSSetEdit 	C_Lot_No, "Lot No.",  15,,,9,2           '13 차 추가 
    ggoSpread.SSSetEdit 	C_Lot_Seq, "Lot No.순번",  15,,,15,2      '13 차 추가 
    SetSpreadFloatLocal 	C_Over, "과부족허용율(+)(%)",20,1,6
    SetSpreadFloatLocal 	C_Under,"과부족허용율(-)(%)",20,1,6
    ggoSpread.SSSetCombo	C_IOFlg,"VAT포함여부", 15,2,False               '13 차 추가 
    ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
    ggoSpread.SSSetCombo 	C_IOFlgCd, "VAT포함여부코드", 15,2,False
    ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd
    ggoSpread.SSSetEdit 	C_VatType, "VAT", 7,,,4,2
    ggoSpread.SSSetButton 	C_Popup7
    ggoSpread.SSSetEdit 	C_VatNm, "VAT명", 20
    SetSpreadFloatLocal		C_VatRate, "VAT율(%)",15,1,5
    SetSpreadFloatLocal		C_VatAmt, "VAT금액",15,1,2
    SetSpreadFloatLocal		C_OrgVatAmt, "OrgVatAmt",15,1,2
    ggoSpread.SSSetEdit 	C_RetCd , "반품유형", 10,,,4,2
    ggoSpread.SSSetButton 	C_Popup8
    ggoSpread.SSSetEdit 	C_RetNm , "반품유형명", 20
    SetSpreadFloatLocal		C_Bal_Qty, "Bal. Qty.",15,1,3
    SetSpreadFloatLocal		C_Bal_Doc_Amt, "Bal. Doc. Amt.",15,1,2
    SetSpreadFloatLocal		C_Bal_Loc_Amt, "Bal. Loc. Amt.",15,1,2
    SetSpreadFloatLocal		C_ExRate, "Xch. Rate",15,1,5
    ggoSpread.SSSetEdit 	C_SeqNo, "순번", 10
    ggoSpread.SSSetEdit 	C_PrNo, "구매요청번호", 20
    ggoSpread.SSSetEdit 	C_MvmtNo, "구매입고번호", 20
    ggoSpread.SSSetEdit 	C_PoNo, "발주번호", 20
    ggoSpread.SSSetEdit 	C_PoSeqNo, "발주SEQNO", 20
    ggoSpread.SSSetEdit 	C_MaintSeq, "maintseq", 10
	ggoSpread.SSSetEdit 	C_SoNo, "", 10
	ggoSpread.SSSetEdit 	C_SoSeqNo, "", 10
    ggoSpread.SSSetEdit 	C_Stateflg, "stateflg", 10
    ggoSpread.SSSetEdit 	C_reference, "reference", 10
    ggoSpread.SSSetEdit 	C_Remrk, "비고", 20,,,120,2

	Call ggoSpread.MakePairsColumn(C_PlantCd,C_Popup1)
	Call ggoSpread.MakePairsColumn(C_ItemCd,C_Popup2)
	Call ggoSpread.MakePairsColumn(C_OrderUnit,C_Popup3)
	Call ggoSpread.MakePairsColumn(C_HSCd,C_Popup5)
	Call ggoSpread.MakePairsColumn(C_SLCd,C_Popup6)
	Call ggoSpread.MakePairsColumn(C_TrackingNo,C_TrackingNoPop)
	Call ggoSpread.MakePairsColumn(C_VatType,C_Popup7)
	Call ggoSpread.MakePairsColumn(C_RetCd,C_Popup8)

	Call ggoSpread.SSSetColHidden(C_SeqNo,C_SeqNo,True)
	Call ggoSpread.SSSetColHidden(C_Lot_Seq,C_Lot_Seq,True)
	Call ggoSpread.SSSetColHidden(C_Lot_No,C_Lot_No,True)
	Call ggoSpread.SSSetColHidden(C_IOFlgCd,C_IOFlgCd,True)
	Call ggoSpread.SSSetColHidden(C_Bal_Qty,C_Bal_Qty,True)
	Call ggoSpread.SSSetColHidden(C_Bal_Doc_Amt,C_Bal_Doc_Amt,True)
	Call ggoSpread.SSSetColHidden(C_Bal_Loc_Amt,C_Bal_Loc_Amt,True)
	Call ggoSpread.SSSetColHidden(C_ExRate,C_ExRate,True)
	Call ggoSpread.SSSetColHidden(C_CostConCd,C_CostConCd,True)
	'Call ggoSpread.SSSetColHidden(C_PrNo,C_PrNo,True)
	Call ggoSpread.SSSetColHidden(C_MvmtNo,C_MvmtNo,True)
	Call ggoSpread.SSSetColHidden(C_PoNo,C_PoNo,True)
	Call ggoSpread.SSSetColHidden(C_PoSeqNo,C_PoSeqNo,True)
	Call ggoSpread.SSSetColHidden(C_MaintSeq,C_MaintSeq,True)
	Call ggoSpread.SSSetColHidden(C_SoNo,C_SoNo,True)
	Call ggoSpread.SSSetColHidden(C_SoSeqNo,C_SoSeqNo,True)
	Call ggoSpread.SSSetColHidden(C_Stateflg,C_Stateflg,True)
	Call ggoSpread.SSSetColHidden(C_RetCd,C_RetCd,True)
	Call ggoSpread.SSSetColHidden(C_Popup8,C_Popup8,True)
	Call ggoSpread.SSSetColHidden(C_RetNm,C_RetNm,True)
	Call ggoSpread.SSSetColHidden(C_OrgNetAmt,C_OrgNetAmt,True)
	Call ggoSpread.SSSetColHidden(C_OrgNetAmt1,C_OrgNetAmt1,True)
	Call ggoSpread.SSSetColHidden(C_reference,C_reference,True)
	Call ggoSpread.SSSetColHidden(C_OrgVatAmt,C_OrgVatAmt,True)


    ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
    ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
    ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
    ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd

    Call SetSpreadLock

	.ReDraw = true

    End With

End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    With frm1
    ggoSpread.SpreadLock frm1.vspddata.maxcols,-1
    ggoSpread.SpreadLock C_SeqNo , -1
    ggoSpread.SpreadLock C_PlantCd , -1
    'ggoSpread.sssetrequired C_PlantCd, -1
    ggoSpread.SpreadLock C_Popup1 , -1
    ggoSpread.spreadlock C_PlantNm , -1
    'ggoSpread.sssetrequired C_ItemCd, -1
    ggoSpread.SpreadLock C_ItemCd, -1
    'ggoSpread.sssetrequired C_ItemCd, -1
    ggoSpread.spreadlock C_SpplSpec,-1         '품목규격 추가 
    ggoSpread.SpreadLock C_Popup2 , -1
    ggoSpread.spreadlock C_ItemNm , -1
    ggoSpread.SpreadUnLock C_OrderQty, -1
    ggoSpread.sssetrequired C_OrderQty, -1
    ggoSpread.SpreadUnLock C_OrderUnit , -1
    ggoSpread.sssetrequired C_OrderUnit, -1
    ggoSpread.SpreadUnLock C_Popup3 , -1
    ggoSpread.SpreadUnLock C_Cost , -1
    ggoSpread.sssetrequired C_Cost, -1
    ggoSpread.SpreadUnLock C_CostCon, -1
    ggoSpread.sssetrequired C_CostCon, -1
    'ggoSpread.SpreadLock C_CostCon, -1

    'If Trim(.hdnreference.value) = "N" then
    '    ggoSpread.spreadlock C_OrderAmt, -1
    'else
    '    ggoSpread.SpreadUnLock C_OrderAmt, -1
    '    ggoSpread.sssetrequired C_OrderAmt, -1
    'end if

    ggoSpread.spreadlock C_NetAmt, -1
    ggoSpread.SpreadUnLock C_DlvyDT, -1
    ggoSpread.sssetrequired C_DlvyDT, -1
    ggoSpread.spreadlock C_HSCd, -1
    ggoSpread.spreadlock C_Popup5, -1
    ggoSpread.spreadlock C_HSNm, -1
    ggoSpread.SpreadUnLock C_SLCd , -1
    ggoSpread.sssetrequired C_SLCd, -1
    ggoSpread.SpreadUnLock C_Popup6 , -1
    ggoSpread.spreadlock C_SLNm, -1

    ggoSpread.SpreadUnLock C_VatType , -1
    ggoSpread.SpreadUnLock C_Popup7 , -1
    ggoSpread.SpreadUnLock C_VatNm , -1
    ggoSpread.SpreadUnLock C_VatRate , -1
    ggoSpread.SpreadUnLock C_VatAmt , -1

    ggoSpread.SpreadUnLock C_Popup8 , -1
    ggoSpread.spreadlock C_RetNm , -1
    ggoSpread.spreadUnLock C_IOFlg , -1    '13차추가 
    ggoSpread.sssetrequired C_IOFlg, -1
    ggoSpread.SpreadLock C_IOFlgCd, -1
    ggoSpread.spreadlock C_Lot_No , -1     '13차추가 
    ggoSpread.spreadlock C_Lot_Seq , -1    '13차추가 
    ggoSpread.spreadlock C_TrackingNo , -1
	ggoSpread.spreadUnlock C_Under, -1
	ggoSpread.spreadUnlock C_Over, -1
	
	ggoSpread.spreadlock C_PrNo, -1       '2006-09

    End With


End Sub

Sub SetSpreadLockAfterQuery()

	Dim index,Count,index1 , strReqChk

    With frm1

   .vspdData.ReDraw = False

    if .vspdData.MaxRows < 1 then
		if .hdnRelease.Value <> "Y" then
			'Call SetToolbar("1110111111101")
		End if
		Exit sub
	end if

	'index1 = Cint(.hdnmaxrow.value) + 1

    if .hdnRelease.Value = "Y" then
		For index = C_PlantCd to C_Stateflg
			ggoSpread.SpreadLock index , -1
		Next
	Else

		For index1 = Cint(.hdnmaxrows.value) + 1 to .vspdData.MaxRows
		    ggoSpread.SpreadLock frm1.vspddata.maxcols, index1, frm1.vspddata.maxcols, index1
			ggoSpread.SpreadLock C_SeqNo , index1,C_SeqNo,index1
			ggoSpread.SpreadLock C_PlantCd ,index1,C_PlantCd,index1
			ggoSpread.SpreadLock C_Popup1 , index1,C_Popup1,index1
			ggoSpread.spreadlock C_PlantNm , index1,C_PlantNm,index1
			ggoSpread.SpreadLock C_ItemCd, index1,C_ItemCd,index1
			ggoSpread.SpreadLock C_Popup2 , index1,C_Popup2,index1
			ggoSpread.spreadlock C_ItemNm , index1,C_ItemNm,index1
			ggoSpread.spreadlock C_SpplSpec,index1,C_SpplSpec,index1         '품목규격 추가 
			ggoSpread.SpreadUnLock C_OrderQty,index1,C_OrderQty,index1
			ggoSpread.sssetrequired C_OrderQty, index1,index1

			if UCase(frm1.hdnRetflg.Value) = "N" then
				ggoSpread.SpreadUnLock C_OrderUnit , index1,C_OrderUnit,index1
				ggoSpread.sssetrequired C_OrderUnit, index1,index1
				ggoSpread.SpreadUnLock C_Popup3 , index1,C_Popup3,index1
				ggoSpread.SpreadUnLock C_Cost , index1,C_Cost,index1
				ggoSpread.sssetrequired C_Cost, index1,index1
			else
				ggoSpread.SpreadLock C_OrderUnit , index1,C_OrderUnit,index1
				ggoSpread.SpreadLock C_Popup3 , index1,C_Popup3,index1
				ggoSpread.SpreadLock C_Cost , index1,C_Cost,index1
			end if

			ggoSpread.SpreadUnLock C_CostCon, index1,C_CostCon,index1
			ggoSpread.sssetrequired C_CostCon, index1,index1
			ggoSpread.spreadlock C_NetAmt, index1,C_NetAmt,index1

			if .hdnImportflg.value = "Y" then
				ggoSpread.spreadUnlock C_HSCd , index1,C_HSCd,index1
				ggoSpread.sssetrequired C_HSCd, index1,index1
				ggoSpread.spreadUnlock C_Popup5 , index1,C_Popup5,index1
				ggoSpread.spreadlock C_HSNm , index1,C_HSNm,index1
			else
				ggoSpread.spreadlock C_HSCd, index1,C_HSCd,index1
				ggoSpread.spreadlock C_Popup5, index1,C_Popup5,index1
				ggoSpread.spreadlock C_HSNm, index1,C_HSNm,index1
			End if

'			If Trim(.hdnreference.value) = "N" then
'			     ggoSpread.SSSetProtected	C_OrderAmt, index1, index1
'			else
			    ggoSpread.SSSetRequired  C_OrderAmt, index1, index1
'			end if

			ggoSpread.spreadlock C_TrackingNo , index1,C_TrackingNo,index1
			ggoSpread.SpreadUnLock C_IOFlg, index1,C_IOFlgCd,index1
			ggoSpread.SSSetRequired	C_IOFlg, index1,index1
			ggoSpread.SSSetRequired	C_IOFlgCd, index1,index1

			ggoSpread.SpreadUnLock C_VatType , index1,C_VatType,index1
			ggoSpread.SpreadUnLock C_Popup7 , index1,C_Popup7,index1
			ggoSpread.spreadlock C_VatNm , index1,C_VatNm,index1
			ggoSpread.spreadlock C_VatRate , index1,C_VatRate,index1
			ggoSpread.spreadlock C_VatAmt , index1,C_VatAmt,index1
		'******************************************
		  '13차추가]
			if .hdnRetflg.Value = "Y" then
				ggoSpread.spreadUnLock C_RetCd , index1,C_RetCd,index1
				ggoSpread.SpreadUnLock C_Popup8 , index1,C_Popup8,index1
				ggoSpread.spreadlock   C_RetNm , index1,C_RetNm,index1
				ggoSpread.spreadUnLock C_Lot_No , index1,C_Lot_No,index1
				ggoSpread.spreadUnLock C_Lot_Seq , index1,C_Lot_Seq,index1
			else
				ggoSpread.spreadlock C_RetCd , index1,C_RetCd,index1
				ggoSpread.spreadlock C_Popup8 , index1,C_Popup8,index1
				ggoSpread.spreadlock C_RetNm , index1,C_RetNm,index1
		        ggoSpread.spreadlock C_Lot_No , index1,C_Lot_No,index1
		        ggoSpread.spreadlock C_Lot_Seq , index1,C_Lot_Seq,index1
		    end if
		'******************************************
		    ggoSpread.SpreadUnLock C_SLCd , index1,C_SLCd,index1
		    ggoSpread.sssetrequired C_SLCd, index1,index1
		    ggoSpread.SpreadUnLock C_Popup6 , index1,C_Popup6,index1
		    ggoSpread.spreadlock C_SLNm, index1,C_SLNm,index1

            .vspdData.Row = index1
			.vspdData.Col = C_TrackingNo
			if Trim(.vspdData.Text) = "*" then
				ggoSpread.spreadlock C_TrackingNo, index1, C_TrackingNoPop, index1
			else
				ggoSpread.spreadUnlock C_TrackingNo, index1, C_TrackingNoPop, index1
				ggoSpread.sssetrequired C_TrackingNo, index1, index1
			end if

			'************************************************ 13차 

			frm1.vspdData.Row = index1
		    frm1.vspdData.Col = C_PrNo

			if Trim(.vspdData.Text) <> "" then
				ggoSpread.spreadlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.spreadlock C_Popup3 , index1, C_Popup3, index1
		        ggoSpread.spreadlock C_DlvyDT, index1,C_DlvyDT, index1
		        ggoSpread.spreadlock C_TrackingNo, index1, C_TrackingNoPop, index1

				ggoOper.SetReqAttr	frm1.txtGroupCd, "Q"
				ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"
			else
				ggoSpread.spreadUnlock C_OrderUnit, index1, C_OrderUnit, index1
				ggoSpread.sssetrequired C_OrderUnit, index1, index1
				ggoSpread.SpreadUnLock C_DlvyDT, index1,C_DlvyDT, index1
			    ggoSpread.sssetrequired C_DlvyDT, index1, index1
			end if
		    ggoSpread.spreadUnlock C_Under,index1,C_Under,index1
		    ggoSpread.spreadUnlock C_Over,index1,C_Over,index1
		    ggoSpread.spreadlock C_PrNo, index1, C_PrNo, index1
	    next
	End if

	.vspdData.ReDraw = True
	End With

	if frm1.hdnImportflg.value = "Y" then
	    ggoOper.SetReqAttr	frm1.txtDvryDt, "N"
	else
		ggoOper.SetReqAttr	frm1.txtDvryDt, "D"
		ggoOper.SetReqAttr	frm1.txtOffDt, "Q"
		ggoOper.SetReqAttr	frm1.txtApplicantCd, "Q"
		ggoOper.SetReqAttr	frm1.txtApplicantNm, "Q"
		ggoOper.SetReqAttr	frm1.txtIncotermsCd, "Q"
		ggoOper.SetReqAttr	frm1.txtIncotermsNm, "Q"
		ggoOper.SetReqAttr	frm1.txtTransCd, "Q"
		ggoOper.SetReqAttr	frm1.txtTransNm, "Q"
	end if

End Sub
'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SeqNo		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_PlantCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ItemCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SpplSpec	, pvStartRow, pvEndRow '품목규격 추가 
    ggoSpread.SSSetProtected	C_PrNo	, pvStartRow, pvEndRow      '2008-05-23 9:44오전 :: hanc
    ggoSpread.SSSetRequired		C_OrderQty	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_OrderUnit	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_Cost		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_CostCon	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_CostConCd	, pvStartRow, pvEndRow

'    If Trim(.hdnreference.value) = "N" then
'        ggoSpread.SSSetProtected	C_OrderAmt, pvStartRow, pvEndRow
'    else
        ggoSpread.SSSetRequired  C_OrderAmt, pvStartRow, pvEndRow
'    end if

    ggoSpread.SSSetProtected	C_NetAmt, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_DlvyDt, pvStartRow, pvEndRow

    if Trim(frm1.hdnImportflg.value) <> "Y" then
	    ggoSpread.SSSetProtected	C_HSCd	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_Popup5, pvStartRow, pvEndRow
	else
		ggoSpread.spreadUnlock	C_HSCd	, pvStartRow, C_HSCd, pvEndRow
		ggoSpread.sssetrequired	C_HSCd	, pvStartRow, pvEndRow
		ggoSpread.spreadUnlock	C_Popup5, pvStartRow, C_Popup5, pvEndRow
	end if

	ggoSpread.SSSetProtected		C_TrackingNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_TrackingNoPop, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_HSNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_SLCd	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_SLNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatRate, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatAmt , pvStartRow, pvEndRow
    '******************************************
	ggoSpread.SSSetRequired		C_IOFlg	 , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected		C_IOFlgCd, pvStartRow, pvEndRow  '13차추가 
	if .hdnRetflg.Value <> "Y" then
		ggoSpread.SSSetProtected C_RetCd	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Popup8, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RetNm	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Lot_No, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Lot_Seq, pvStartRow, pvEndRow
	end if
	'******************************************
    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColorRef
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColorRef(ByVal pvStartRow, ByVal pvEndRow)
    With frm1

    ggoSpread.SSSetProtected	frm1.vspddata.maxcols, pvStartRow, pvEndRow
    ggoSpread.spreadUnlock C_PlantCd,pvStartRow,frm1.vspddata.maxcols -1,pvStartRow
    ggoSpread.SSSetProtected	C_SeqNo		, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_PlantCd	, pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_PlantCd	, pvStartRow, C_PlantCd, pvEndRow
	ggoSpread.spreadlock		C_Popup1	, pvStartRow, C_Popup1,  pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm	, pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_ItemCd	, pvStartRow, C_ItemCd, pvEndRow
	ggoSpread.spreadlock		C_Popup2	, pvStartRow, C_Popup2,  pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_SpplSpec	, pvStartRow, pvEndRow '품목규격 추가 
    ggoSpread.SSSetRequired		C_OrderQty	, pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_OrderUnit	, pvStartRow, C_OrderUnit, pvEndRow
    ggoSpread.spreadlock		C_Popup3	, pvStartRow, C_OrderUnit, pvEndRow
    ggoSpread.SSSetRequired		C_Cost		, pvStartRow, pvEndRow
    'ggoSpread.SSSetRequired		C_CostCon	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_CostConCd	, pvStartRow, pvEndRow

'    If Trim(.hdnreference.value) = "N" then
'         ggoSpread.SSSetProtected	C_OrderAmt, pvStartRow, pvEndRow
'    else
        ggoSpread.SSSetRequired  C_OrderAmt, pvStartRow, pvEndRow
'    end if

    ggoSpread.SSSetProtected	C_NetAmt, pvStartRow, pvEndRow
    ggoSpread.spreadlock		C_DlvyDt	, pvStartRow, C_DlvyDt, pvEndRow
    if Trim(frm1.hdnImportflg.value) <> "Y" then
	    ggoSpread.SSSetProtected	C_HSCd	, pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected	C_Popup5, pvStartRow, pvEndRow
	else
		ggoSpread.spreadUnlock	C_HSCd	, pvStartRow, C_HSCd, pvEndRow
		ggoSpread.sssetrequired	C_HSCd	, pvStartRow, pvEndRow
		ggoSpread.spreadUnlock	C_Popup5, pvStartRow, C_Popup5, pvEndRow
	end if

	ggoSpread.SSSetProtected		C_TrackingNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_TrackingNoPop, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_HSNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_SLCd	, pvStartRow, pvEndRow
    ggoSpread.spreadUnlock		C_Popup6	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_SLNm	, pvStartRow, pvEndRow
    ggoSpread.spreadUnlock		C_VatType	, pvStartRow, pvEndRow
     ggoSpread.spreadUnlock		C_Popup7	, pvStartRow, pvEndRow
'    ggoSpread.spreadUnlock		C_VatNm	, pvStartRow, pvEndRow
 '   ggoSpread.spreadUnlock		C_VatRate	, pvStartRow, pvEndRow
  '  ggoSpread.spreadUnlock		C_VatAmt	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatNm	, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatRate, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected		C_VatAmt , pvStartRow, pvEndRow
    '******************************************
	ggoSpread.SSSetRequired		C_IOFlg	 , pvStartRow, pvEndRow
	ggoSpread.SSSetProtected		C_IOFlgCd, pvStartRow, pvEndRow  '13차추가 
	if .hdnRetflg.Value <> "Y" then
		ggoSpread.SSSetProtected C_RetCd	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Popup8, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_RetNm	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Lot_No, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Lot_Seq, pvStartRow, pvEndRow
	end if
	'******************************************
    End With
End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   :
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos

	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PlantCd 		= iCurColumnPos(1) 
			C_Popup1		= iCurColumnPos(2) 
			C_PlantNm 		= iCurColumnPos(3) 
			C_itemCd 		= iCurColumnPos(4) 
			C_Popup2 		= iCurColumnPos(5) 
			C_itemNm 		= iCurColumnPos(6) 
			C_SpplSpec      = iCurColumnPos(7) 
			C_OrderQty		= iCurColumnPos(8) 
			C_OrderUnit		= iCurColumnPos(9) 
			C_Popup3		= iCurColumnPos(10)
			C_Cost			= iCurColumnPos(11)
			C_Check			= iCurColumnPos(12)
			C_CostCon		= iCurColumnPos(13)
			C_CostConCd		= iCurColumnPos(14)
			C_OrderAmt		= iCurColumnPos(15)
			C_NetAmt        = iCurColumnPos(16)
			C_OrgNetAmt     = iCurColumnPos(17)
			C_IOFlg		    = iCurColumnPos(18)
			C_IOFlgCd	    = iCurColumnPos(19)
			C_VatType       = iCurColumnPos(20)
			C_Popup7        = iCurColumnPos(21)
			C_VatNm         = iCurColumnPos(22)
			C_VatRate       = iCurColumnPos(23)
			C_VatAmt        = iCurColumnPos(24)
			C_OrgVatAmt     = iCurColumnPos(25)
			C_DlvyDT		= iCurColumnPos(26)
			C_HSCd			= iCurColumnPos(27)
			C_Popup5		= iCurColumnPos(28)
			C_HSNm			= iCurColumnPos(29)
			C_SLCd			= iCurColumnPos(30)
			C_Popup6		= iCurColumnPos(31)
			C_SLNm			= iCurColumnPos(32)
			C_TrackingNo	= iCurColumnPos(33)
			C_TrackingNoPop	= iCurColumnPos(34)
			C_Lot_No        = iCurColumnPos(35)
			C_Lot_Seq       = iCurColumnPos(36)
			C_RetCd         = iCurColumnPos(37)
			C_Popup8        = iCurColumnPos(38)
			C_RetNm         = iCurColumnPos(39)
			C_Over			= iCurColumnPos(40)
			C_Under			= iCurColumnPos(41)
			C_Bal_Qty		= iCurColumnPos(42)
			C_Bal_Doc_Amt	= iCurColumnPos(43)
			C_Bal_Loc_Amt	= iCurColumnPos(44)
			C_ExRate		= iCurColumnPos(45)
			C_SeqNo 		= iCurColumnPos(46)
			C_PrNo			= iCurColumnPos(47)
			C_MvmtNo		= iCurColumnPos(48)
			C_PoNo			= iCurColumnPos(49)
			C_PoSeqNo		= iCurColumnPos(50)
			C_MaintSeq		= iCurColumnPos(51)
			C_SoNo			= iCurColumnPos(52)
			C_SoSeqNo		= iCurColumnPos(53)
			C_OrgNetAmt1    = iCurColumnPos(54)
			C_reference     = iCurColumnPos(55)
			C_Stateflg		= iCurColumnPos(56)
			C_Remrk			= iCurColumnPos(57)

	End Select

End Sub
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgIntFlgMode_Dtl = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    IsOpenPop = False
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
    lgStrPrevKey = ""
    frm1.vspdData.MaxRows = 0


End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.
'*********************************************************************************************************

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.
'*********************************************************************************************************

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

	lgOpenFlag	= False
	lgTabClickFlag	= False
	gSelframeFlg = TAB1
	lgReqRefChk = False

    Call SetToolbar("1110100000001111")
    frm1.rdoRelease(0).checked = true
    frm1.txtOffDt.text = EndDate
    frm1.txtPoDt.text = EndDate
    frm1.hdnCurr.value = Parent.gCurrency
    frm1.btnCfm.disabled = true
    ' === 2005.07.15 단가 일괄불러오기 =============
    frm1.btnCallPrice.disabled = True
    ' === 2005.07.15 단가 일괄불러오기 =============
    'frm1.btnSel.disabled = true

    frm1.btnSend.disabled = true
    frm1.txtGroupCd.Value = Parent.gPurGrp
    frm1.txtXch.Text = ""
	frm1.txtApplicantCd.value = Parent.gCompany
	frm1.txtApplicantNm.value = Parent.gCompanyNm
	frm1.btnCfm.value = "확정"
	frm1.txtPoNo.focus
'	Call InitComboBox
	frm1.cboXchop.value = "*"
	frm1.hdnxchrateop.value ="*"
	frm1.hdnMergPurFlg.value = "N"

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
  	frm1.txtGroupCd.value = lgPGCd
	End If

	Set gActiveElement = document.activeElement

End Sub
'==========================================================================================
'   Event Name : InitComboBox
'   Event Desc : 콤보 박스 초기화 
'==========================================================================================

Sub InitComboBox()
	Call SetCombo(frm1.cboXchop,"*","*")
	Call SetCombo(frm1.cboXchop,"/","/")
End Sub

'******************************************  2.3 Operation 처리함수  *************************************
'	기능: Operation 처리부분 
'	설명: Tab처리, Reference등을 행한다.
'*********************************************************************************************************
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function

	Call changeTabs(TAB1)	<% '~~~ 첫번째 Tab %>
	gSelframeFlg = TAB1

   	'Call setFocus(CLICK_HEADER)
   	frm1.txtPoNo.focus
	'Call SetToolbar("11111000001111")
	Call BtnToolCtrl(TAB1)

	Set gActiveElement = document.activeElement

End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function

	if frm1.txtPotypeCd.value = "" then
		Call DisplayMsgBox("179010", "X", "X", "X")
		frm1.txtPotypeCd.focus
		Exit Function
	End if

   	Call changeTabs(TAB2)
	gSelframeFlg = TAB2

	frm1.txtPoNo.focus
	Call BtnToolCtrl(TAB2)



	Set gActiveElement = document.activeElement


End Function

Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function

   	if frm1.txtPotypeCd.value = "" then
		Call DisplayMsgBox("179010", "X", "X", "X")
		frm1.txtPotypeCd.focus
		Exit Function
	End if

	IF frm1.hdnImportflg.value<>"Y" then
		Call DisplayMsgBox("17a007", "X", "X", "X")
		lgOpenFlag = False
		Exit Function
	End if

	Call changeTabs(TAB3)

	lgOpenFlag	= False
	lgTabClickFlag = False
	gSelframeFlg = TAB3
	frm1.txtExpiryDt.focus
	Call BtnToolCtrl(TAB3)

	Set gActiveElement = document.activeElement

End Function

'------------------------------------------  SetClickflag, ResetClickflag()  -----------------------------
'	Name : SetClickflag, ResetClickflag()
'	Description :
'---------------------------------------------------------------------------------------------------------

Function SetClickflag()
	lgTabClickFlag = True
End Function

Function ResetClickflag()
	lgTabClickFlag = False
End Function

Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD='B9001' And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description
		Err.Clear
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub

'========================================================================================
' Function Name : GetCollectTypeRef
' Function Desc :
'========================================================================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)
		If arrCollectVatType(iCnt, 0) = UCASE(VatType) Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = arrCollectVatType(iCnt, 2)
			Exit Sub
		End If
	Next
	VatTypeNm = ""
	VatRate = ""
End Sub

'========================================================================================
' Function Name : SetVatType
'========================================================================================
Sub SetVatType(byVal iRow)
	Dim VatType, VatTypeNm, VatRate
	Dim txtVatRate ,txtVatAmt, chk_vat_flg

	With frm1.vspdData

       .Row = iRow
	   .Col = C_VatType

		VatType = .text

		Call InitCollectType
		Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

       .Col = C_VatNm
       .text = VatTypeNm
       .Col = C_VatRate
	   .text = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
		txtVatRate =  UNICDbl(.text)


	   ' vat 금액계산 
	   ' 부가세 포함/불포함 부가세 계산 변경 2002.3.9 L.I.P
		.Col		= C_IOFlgCd
		chk_vat_flg	= .text

       .Col          = C_OrderAmt
		if chk_vat_flg = "2"	Then
			txtVatAmt    = UNICDbl(.text) * (txtVatRate/(100 + txtVatRate))
		Else
			txtVatAmt    = UNICDbl(.text) * (txtVatRate/100)
		End If

		.Col = C_VatAmt
		.Text = UNIConvNumPCToCompanyByCurrency(txtVatAmt,frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")

 End With

End Sub


'------------------------------------  Setretflg()  ----------------------------------------------
'	Name : Setreference()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub Setreference()

	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim ireference

    Err.Clear

	Call CommonQueryRs(" reference ", " b_configuration ", " major_cd = 'M9016' and minor_cd = 'CH' and seq_no = '1' ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    ireference = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description
		Err.Clear
		Exit Sub
	End If

    if Trim(lgF0) <> "" then
        frm1.hdnreference.value = UCase(Trim(ireference(0)))

    end if

End Sub


'========================================================================================
' Function Name : setCVatFlg
' Function Desc : 부가세 포함에 따른 의제매입계산 처리 
' Append		: 2002-03-09  L.I.P
'========================================================================================
Sub setCVatFlg(byVal iRow)
	Call setVatType(iRow)
End Sub


'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다.
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

'------------------------------------------  OpenReqRef()  -------------------------------------------------
'	Name : OpenReqRef()
'	Description :구매요청참조 
'---------------------------------------------------------------------------------------------------------

Function OpenReqRef()

	Dim strRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD

	'if lgIntFlgMode = Parent.OPMD_CMODE then
	'	Call DisplayMsgBox("900002", "X", "X", "X")
	'	Exit Function
	'End if

    If CheckRunningBizProcess = True Then
		Exit Function
	End If

	if frm1.txtPotypeCd.value = "" then
		Call DisplayMsgBox("179010", "X", "X", "X")
		frm1.txtPotypeCd.focus
		Exit Function
	End if

	if frm1.hdnRelease.Value = "Y" then

		Call DisplayMsgBox("17a008", "X", "X", "X")
		Exit Function
	End if

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtSupplierCd.value)
	if Trim(frm1.txtSupplierCd.value) = "" then
        arrParam(1) = ""
	Else
        arrParam(1) = Trim(frm1.txtSupplierNm.value)
	End if

	arrParam(2) = Trim(frm1.txtGroupCd.value)
	if Trim(frm1.txtGroupCd.value) = "" then
        arrParam(3) = ""
	else
        arrParam(3) = Trim(frm1.txtGroupNm.value)
	End if

	arrParam(4) = Trim(frm1.hdnSubcontraflg.value)
	arrParam(5) = lgReqRefChk

'	strRet = window.showModalDialog("m2111ra1_1.asp", Array(window.parent,arrParam), _
'			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	iCalledAspName = AskPRAspName("M2111RA1_1_KO44")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M2111RA1_1", "X")
		IsOpenPop = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=560px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False

	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetReqRef(strRet)
	End If

End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : item(item_by_Plant) PopUp
'---------------------------------------------------------------------------------------------------------

Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	frm1.vspdData.Col = C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow

	if  Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Exit Function
	End if

	IsOpenPop = True

	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	arrParam(0) = Trim(frm1.vspdData.Text)

	frm1.vspdData.Col=C_ItemCd
	arrParam(1) = Trim(frm1.vspdData.Text)

	if frm1.hdnSubcontraflg.Value <> "Y" then
		arrParam(2) = "36!PP"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
		arrParam(3) = "30!P"
	else
		arrParam(2) = "12!MO"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
		arrParam(3) = "20!O"
	end if
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)

	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 

    iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_ItemCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_ItemNm
		frm1.vspdData.Text = arrRet(1)
		Call ChangeReturnCost()
		Call vspdData_Change(C_ItemCd, frm1.vspdData.ActiveRow)
	End If
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"
	arrParam(1) = "B_PLANT"
	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	arrParam(2) = Trim(frm1.vspdData.text)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "공장"

    arrField(0) = "PLANT_CD"
    arrField(1) = "PLANT_NM"

   arrHeader(0) = "공장"
    arrHeader(1) = "공장명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col=C_ItemCd
		frm1.vspdData.Row=frm1.vspdData.ActiveRow
		frm1.vspdData.text=""

		frm1.vspdData.Col=C_ItemNM
		frm1.vspdData.Row=frm1.vspdData.ActiveRow
		frm1.vspdData.text=""

		frm1.vspdData.Col = C_PlantCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_PlantNm
		frm1.vspdData.Text = arrRet(1)

		Call ChangeReturnCost()
	End If

End Function

'------------------------------------------  OpenHS()  -------------------------------------------------
'	Name : OpenHS()
'	Description : OpenHS PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenHS()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "HS부호"
	arrParam(1) = "B_HS_code"
	frm1.vspdData.Col=C_HSCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	arrParam(2) = Trim(frm1.vspdData.text)
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "HS부호"

    arrField(0) = "HS_CD"
    arrField(1) = "HS_NM"

    arrHeader(0) = "HS부호"
    arrHeader(1) = "HS명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_HSCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_HSNm
		frm1.vspdData.Text = arrRet(1)
		Call vspdData_Change(C_HsCd, frm1.vspdData.ActiveRow)
	End If

End Function
'------------------------------------------  OpenUnit()  -------------------------------------------------
'	Name : OpenUnit()
'	Description : Unit PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위"
	arrParam(1) = "B_Unit_OF_MEASURE"

	frm1.vspdData.Col=C_OrderUnit
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	arrParam(2) = Trim(frm1.vspdData.text)

	arrParam(4) = ""
	arrParam(5) = "단위"

    arrField(0) = "Unit"
    arrField(1) = "Unit_Nm"

    arrHeader(0) = "단위"
    arrHeader(1) = "단위명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col=C_OrderUnit
		frm1.vspdData.text= arrRet(0)
		Call ChangeReturnCost()
	End If
End Function

'------------------------------------------  OpenSL()  -------------------------------------------------
'	Name : OpenSL()
'	Description : Storage_Location PopUp
'-------------------------------------------------------------------------------------------------------
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow

	if Trim(frm1.vspdData.Text)="" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		Exit function
	End if

	arrParam(4) = "PLANT_CD='" & frm1.vspdData.Text & "'"

	IsOpenPop = True

	frm1.vspdData.Col=C_SLCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow

	arrParam(0) = "창고"
	arrParam(1) = "B_STORAGE_LOCATION"

	arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(5) = "창고"

    arrField(0) = "SL_CD"
    arrField(1) = "SL_NM"

    arrHeader(0) = "창고"
    arrHeader(1) = "창고명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_SLCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_SLNm
		frm1.vspdData.Text = arrRet(1)
	End If
End Function
'------------------------------------------  OpenRet()  -------------------------------------------------
'	Name : OpenRet()
'	Description :
'-------------------------------------------------------------------------------------------------------
Function OpenRet()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

'	If IsOpenPop = True Or UCase(frm1.txtVattype.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
    If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    frm1.vspdData.Col=C_RetCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow

	arrParam(0) = "반품유형"
	arrParam(1) = "B_MINOR"

	arrParam(2) = Trim(frm1.vspdData.Text)

	arrParam(4) = "b_minor.MAJOR_CD='b9017' "
	arrParam(5) = "반품유형"

    arrField(0) = "MINOR_CD"
    arrField(1) = "MINOR_NM"


    arrHeader(0) = "반품유형"
    arrHeader(1) = "반품유형명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_RetCd
		frm1.vspdData.Text = arrRet(0)
		frm1.vspdData.Col = C_RetNm
		frm1.vspdData.Text = arrRet(1)
		Call vspdData_Change(C_RetCd, frm1.vspdData.ActiveRow)
	End If
End Function
'------------------------------------------  OpenTrackingNo()  -------------------------------------------
'	Name : OpenTrackingNo()
'	Description : TrackingNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingNo()

	Dim arrRet
	Dim arrParam(6)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	frm1.vspdData.Col=C_PlantCd
	frm1.vspdData.Row=frm1.vspdData.ActiveRow
	If Trim(frm1.vspdData.Text) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		IsOpenPop = False
		Exit Function
	End if

    arrParam(0) = ""
    arrParam(1) = ""
    arrParam(2) = Trim(frm1.vspdData.Text)
	arrParam(3) = ""

	frm1.vspdData.Col=C_SoNo
	frm1.vspdData.Row=frm1.vspdData.ActiveRow

	arrParam(4) = Trim(frm1.vspdData.Text)
	arrParam(5) = " and A.tracking_no not in ('*') "
	arrParam(6) = "M"

'	arrRet = window.showModalDialog("../s3/s3135pa1.asp", Array(window.parent,arrParam), _
'			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	iCalledAspName = AskPRAspName("S3135PA1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet = "" Then
		Exit Function
	Else
		frm1.vspdData.Col = C_TrackingNo
		frm1.vspdData.Text = arrRet
	End If

End Function


'------------------------------------------  SetReqRef()  -------------------------------------------------
'	Name : SetReqRef()
'	Description :구매요청참조 
'---------------------------------------------------------------------------------------------------------
Function SetReqRef(strRet)
	Dim Index1,index2,Index3,Count1,Count2
	Dim IntIflg
	Dim strMessage
	Dim intstartRow,intEndRow, TempRow
	Dim iInsRow,intInsertRowsCount

    Const C_ReqNo_Ref		= 0
	Const C_PlantCd_Ref		= 1
	Const C_PlantNm_Ref		= 2
	Const C_ItemCd_Ref		= 3
	Const C_ItemNm_Ref		= 4
	Const C_SpplSpec_Ref    = 5                         '품목 규격 추가 
	Const C_Qty_Ref			= 6
	Const C_Unit_Ref		= 7
	Const C_DlvyDt_Ref		= 8
	Const C_Pur_Plan_Dt_Ref	= 9
	Const C_Pr_Type_Ref		= 10
	Const C_Pr_Type_Nm_Ref	= 11
	Const C_SoNo_Ref		= 12
	Const C_SoSeqNo_Ref		= 13
	Const C_TrackingNo_Ref	= 14
	Const C_SLCd_Ref		= 15
	Const C_SLNm_Ref		= 16
	Const C_HSCd_Ref		= 17
	Const C_HSNm_Ref		= 18
	'이성룡 추가 
	Const C_Over_Tol		= 19
	Const C_Under_Tol		= 20


	Count1 = Ubound(strRet,1)

	Count2 = UBound(strRet,2)
	strMessage = ""

	IntIflg=true

	with frm1

	Call changeTabs(TAB2)
	gSelframeFlg = TAB2

	.vspdData.focus
	ggoSpread.Source = .vspdData
	intStartRow = .vspdData.MaxRows + 1
	.vspdData.Redraw = False

	TempRow = .vspdData.MaxRows					'리스트 max값 

	intInsertRowsCount = 0 '중복 안될때만 MAXROW에 1을 추가하기 위한변수 

	'중복된 요청건참조시 MAXROW값계산 로직 수정 200308
	for index1 = 0 to Count1 - 1

		.vspdData.Row=Index1+1

		If TempRow <> 0 Then
			For Index3 = 1 to TempRow
				if GetSpreadText(.vspdData,C_PrNo,index3,"X","X") = strRet(index1,C_ReqNo_Ref) then
					strMessage = strMessage & strRet(Index1,C_ReqNo_Ref) & ";"
					intIflg=False
					intInsertRowsCount = 0		'중복될땐 MAXROW를 증가시키지 않음.
					Exit for
				Else
					intInsertRowsCount =  1
				End if
			Next
		Else
			intInsertRowsCount =  1
		End If

		if IntIflg <> False then
			lgReqRefChk = true

			.vspdData.MaxRows = CLng(TempRow) + CLng(intInsertRowsCount)
			iInsRow = CLng(TempRow) + CLng(intInsertRowsCount)

			TempRow = CLng(TempRow) + CLng(intInsertRowsCount) '다음 MAXROW계산시 베이스가 될 TempRow 를 증가시킴.
			lgBlnFlgChgValue = True

			Call .vspdData.SetText(0		,	iInsRow, ggoSpread.InsertFlag)
'			Call .vspdData.SetText(C_VatType,	iInsRow, .hdnVATType.value)

			If Trim(.hdnVATINCFLG.value) ="2" Then	'포함 
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,0,"X","X")
				Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,0,"X","X")
			Else
				Call SetSpreadValue(.vspdData,C_IOFlg	,iInsRow,1,"X","X")
				Call SetSpreadValue(.vspdData,C_IOFlgCd	,iInsRow,1,"X","X")
			End If

'			If .hdnVATType.value <> "" Then
'				call SetVatType(iInsRow)
'			End If

			Call SetSpreadValue(.vspdData,C_CostCon	,iInsRow,1,"X","X")
			Call SetSpreadValue(.vspdData,C_CostConCd	,iInsRow,1,"X","X")

			'Call `("C",iInsRow)
			Call SetState("R",iInsRow)

			Call .vspdData.SetText(C_PlantCd	,	iInsRow, strRet(index1,C_PlantCd_Ref))
			Call .vspdData.SetText(C_PlantNm	,	iInsRow, strRet(index1,C_PlantNm_Ref))
			Call .vspdData.SetText(C_itemCd		,	iInsRow, strRet(index1,C_ItemCd_Ref))
			Call .vspdData.SetText(C_itemNm		,	iInsRow, strRet(index1,C_ItemNm_Ref))
			Call .vspdData.SetText(C_SpplSpec	,	iInsRow, strRet(index1,C_SpplSpec_Ref))
			Call .vspdData.SetText(C_OrderQty	,	iInsRow, strRet(index1,C_Qty_Ref))
			Call .vspdData.SetText(C_OrderUnit	,	iInsRow, strRet(index1,C_Unit_Ref))
			Call .vspdData.SetText(C_SoNo		,	iInsRow, strRet(index1,C_SoNo_Ref))
			Call .vspdData.SetText(C_SoSeqNo	,	iInsRow, strRet(index1,C_SoSeqNo_Ref))
			Call .vspdData.SetText(C_DlvyDT		,	iInsRow, strRet(index1,C_DlvyDt_Ref))
			Call .vspdData.SetText(C_SLCd		,	iInsRow, strRet(index1,C_SLCd_Ref))
			Call .vspdData.SetText(C_SLNm		,	iInsRow, strRet(index1,C_SLNm_Ref))
			Call .vspdData.SetText(C_HSCd		,	iInsRow, strRet(index1,C_HSCd_Ref))
			Call .vspdData.SetText(C_HSNm		,	iInsRow, strRet(index1,C_HSNm_Ref))
			Call .vspdData.SetText(C_PrNo		,	iInsRow, strRet(index1,C_ReqNo_Ref))
			Call .vspdData.SetText(C_TrackingNo	,	iInsRow, strRet(index1,C_TrackingNo_Ref))
			'이성룡 추가 
			Call .vspdData.SetText(C_Over	,	iInsRow, strRet(index1,C_Over_Tol))
			Call .vspdData.SetText(C_Under	,	iInsRow, strRet(index1,C_Under_Tol))
		Else
			IntIFlg=True
		End if
	next

	intEndRow = iInsRow

'	strBpcd			=	strRet(Count1,0)
'	strPurGrp		=	strRet(Count1,1)
'	strProcuType	=	strRet(Count1,2)

	frm1.txtSupplierCd.value	 = strRet(Count1,0)
	frm1.txtGroupCd.value		 = strRet(Count1,1)
	frm1.hdnSubcontraflg.value	 = strRet(Count1,2)
	frm1.txtGroupNm.value		 = strRet(Count1,3)

	ggoOper.SetReqAttr	frm1.txtGroupCd, "Q"
	ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"
	ggoOper.SetReqAttr	frm1.txtPotypeCd, "Q"

	Call ChangeSupplier()
	if strMessage <> "" then
		Call DisplayMsgBox("17a005", "X",strmessage,"구매요청번호")
		.vspdData.ReDraw = True
		Exit Function
	End if

'	.vspdData.Col 	= C_Stateflg
	'.vspdData.Text = "C"

	Call SetSpreadColorRef(intStartRow,intEndRow)
	Call BtnToolCtrl(TAB2)

	.vspdData.ReDraw = True

     End with
End Function


'==========================================================================================
'   Event Name : InitData()
'   Event Desc : Combo 변경 이벤트 
'==========================================================================================
Sub InitData(lRow)
	Dim intIndex 

		frm1.vspdData.Row = lRow

		frm1.vspdData.Col = C_CostConCd
		intIndex = frm1.vspdData.value
		frm1.vspdData.col = C_CostCon
		frm1.vspdData.value = intindex
End Sub


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
Function SetVatName()
	Dim index1
	with frm1

		For Index1 = 1 to .vspdData.MaxRows step 1
				'Insert Row 시 헤더의 부가세관련 정보 초기값으로 2002.2.19
				.vspdData.Row = index1

				.vspdData.Col = C_VatType
				.vspdData.Text = .hdntxtVatType.value

				.vspdData.Col  = C_VatNm
				.vspdData.Text = .hdntxtVatTypeNm.value

				.vspdData.Col  = C_VatRate
				.vspdData.Text = .hdntxtVatrt.value
		Next
	End With
	'lgReqRefChk = False

End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
Function SetVat(byval arrRet)

    Dim price, chk_vat_flg

    With frm1
		.vspdData.Col = C_VatType
		.vspdData.Text = arrRet(0)
		.vspdData.Col = C_VatNm
		.vspdData.Text = arrRet(1)
		.vspdData.Col = C_VatRate
		.vspdData.Text = arrRet(2)

		.vspdData.Col = C_OrderAmt
		price = UNICDbl(.vspdData.Text)
'	vat 금액계산 
' 부가세 포함/불포함 부가세 계산 변경 2002.3.9 L.I.P
		.vspdData.Col		= C_IOFlgCd
		chk_vat_flg	= .vspdData.text

		.vspdData.Col = C_VatAmt
		if chk_vat_flg = "2"		Then
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(price * UNICDbl(arrRet(2))/(100 + UNICDbl(arrRet(2))),frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
		Else
			.vspdData.Text = UNIConvNumPCToCompanyByCurrency(price * UNICDbl(arrRet(2))/(100 + UNICDbl(arrRet(2))),frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
		End If

	End With
    Call vspdData_Change(C_VatType, frm1.vspdData.ActiveRow)

End Function

'========================================================================================
' Function Name : SetRetCd
' Function Desc : 반납유형 직접 입력시 처리 
'========================================================================================
Sub SetRetCd()
	Dim iRetCd, iRetNm, strQUERY, tmpData
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i

	with frm1.vspdData

		Err.Clear

	   .Col = C_RetCd

		strQUERY = " Minor.MAJOR_CD='B9017' and  Minor.MINOR_CD = " & "'" & FilterVar(Trim( .text), " " , "SNM") & "' "

		Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM ", " B_MINOR Minor ", strQUERY, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If Err.number = 0 Then

			if lgF0 <> "" then
				iRetNm = Split(lgF1, Chr(11))
			   .Col = C_RetNm
			   .text = iRetNm(0)
			  else
			   .Col = C_RetNm
			   .text = ""
			end if
		else
			MsgBox Err.description, VbInformation, parent.gLogoName
			Err.Clear
			Exit Sub
		End If

	End With

End Sub

Function OpenMpOrderRef()

	Dim strRet
	Dim strParam

	if frm1.rdoRelease(1).checked = true then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		Exit Function
	End if

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	strParam = Parent.gColSep'strParam & Trim(frm1.txtSold_to_party.value) & Parent.gColSep
	strParam = strParam & Trim(frm1.txtPoDt.Text)

	strRet = window.showModalDialog("m3011ra1.asp", strParam, _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet(0) = "" Then
		Exit Function
	Else
		Call SetMpOrder(strRet)
	End If

End Function

Function SetMpOrder(strRet)

	frm1.txtMaintNo.value = strRet(0)
	'frm1.RefOnLine.value = "Y"

	Call OnLineQuery()

	lgBlnFlgChgValue = true

End Function
'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------

Function OpenPoNo()

		Dim strRet
		Dim arrParam(2)
		Dim iCalledAspName
		Dim IntRetCD

		If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

		IsOpenPop = True

		arrParam(0) = "N"  'Return Flag
		arrParam(1) = "N"  'Release Flag
		arrParam(2) = ""  'STO Flag

'		strRet = window.showModalDialog("m3111pa1.asp", Array(window.parent,arrParam), _
'				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

		iCalledAspName = AskPRAspName("M3111PA1_KO441")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
			IsOpenPop = False
			Exit Function
		End If

		strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


		IsOpenPop = False

		If strRet(0) = "" Then
			Exit Function
		Else
			Call SetPoNo(strRet(0))
		End If

End Function

Function SetPoNo(strRet)
	frm1.txtPoNo.value = strRet
End Function

'------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : OpenPoType PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPotype()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPotypeCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발주형태"
	arrParam(1) = "M_CONFIG_PROCESS"

	arrParam(2) = Trim(frm1.txtPotypeCd.Value)
	'arrParam(3) = Trim(frm1.txtPotypeNm.Value)

	'if frm1.hdnSubcontraflg.value = "" Then
	'	arrParam(4) = "USAGE_FLG='Y' and Ret_FLG <>'Y' and sto_flg <> 'F'"
	'else
	'	arrParam(4) = "USAGE_FLG='Y' and Ret_FLG <>'Y' and sto_flg <> 'F' and Subcontra_flg ='"&frm1.hdnSubcontraflg.value &"'"
	'end if
	arrParam(4) = "USAGE_FLG='Y' and Ret_FLG <>'Y'"

	'arrParam(4) = "USAGE_FLG='Y'"
	arrParam(5) ="발주형태"

    arrField(0) = "PO_TYPE_CD"
    arrField(1) = "PO_TYPE_NM"

    arrHeader(0) = "발주형태"
    arrHeader(1) = "발주형태명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
	    IsOpenPop = False
		Exit Function
	Else
		Call SetPotype(arrRet)
	    IsOpenPop = False
	End If
End Function

Function SetPotype(byval arrRet)

	frm1.txtPoTypeCd.Value    = arrRet(0)
	frm1.txtPoTypeNm.Value    = arrRet(1)
	lgBlnFlgChgValue = True

	Call PotypeRef()

End Function
'------------------------------------------  OpenCurr()  -------------------------------------------------
'	Name : OpenCurr()
'	Description : OpenCurr PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCurr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtCurr.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "화폐"
	arrParam(1) = "B_Currency"

	arrParam(2) = Trim(frm1.txtCurr.Value)
	'arrParam(3) = Trim(frm1.txtItemNm2.Value)

	arrParam(4) = ""
	arrParam(5) = "화폐"

    arrField(0) = "Currency"
    arrField(1) = "Currency_Desc"

    arrHeader(0) = "화폐"
    arrHeader(1) = "화폐명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCurr(arrRet)
	End If
End Function

Function SetCurr(byval arrRet)
	frm1.txtCurr.Value    = arrRet(0)
	frm1.txtCurrNm.Value  = arrRet(1)
	Call ChangeCurr()
	lgBlnFlgChgValue = True
End Function

Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"
	arrParam(1) = "B_BIZ_PARTNER"

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)

	arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"
	arrParam(5) = "공급처"

    arrField(0) = "BP_Cd"
    arrField(1) = "BP_NM"
    arrField(2) = "BP_RGST_NO"

    arrHeader(0) = "공급처"
    arrHeader(1) = "공급처명"
    arrHeader(2) = "사업자등록번호"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSupplier(arrRet)
	End If

End Function

Function SetSupplier(byval arrRet)

	frm1.txtSupplierCd.Value    = arrRet(0)
	frm1.txtSupplierNm.Value    = arrRet(1)
	lgBlnFlgChgValue = True

	Call SpplRef()

End Function


Function OpenVat(byVal chk)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtVattype.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VAT형태"
	arrParam(1) = "B_MINOR,b_configuration"

	arrParam(2) = Trim(frm1.txtVattype.Value)

	arrParam(4) = "b_minor.MAJOR_CD='b9001' and b_minor.minor_cd=b_configuration.minor_cd "
	arrParam(4) = arrParam(4) & "and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"
	arrParam(5) = "VAT형태"

    arrField(0) = "b_minor.MINOR_CD"
    arrField(1) = "b_minor.MINOR_NM"
    arrField(2) = "b_configuration.REFERENCE"

    arrHeader(0) = "VAT형태"
    arrHeader(1) = "VAT형태명"
    arrHeader(2) = "VAT율"
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		if chk = 1 then
			Call SetVat_H(arrRet)
		Else
			Call SetVat(arrRet)
		End if
	End If
End Function


Function SetVat_H(byval arrRet)
	frm1.txtVattype.Value		 = arrRet(0)
	frm1.txtVattypeNm.Value      = arrRet(1)
	frm1.txtVatRt.Value = UNIFormatNumber(UNICDbl(arrRet(2)), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

	lgBlnFlgChgValue = True

End Function

Function OpenBank()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBankCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "송금은행"
	arrParam(1) = "B_Bank"

	arrParam(2) = Trim(frm1.txtBankCd.Value)
	'arrParam(3) = Trim(frm1.txtGroupNm.Value)

	arrParam(4) = ""
	arrParam(5) = "송금은행"

    arrField(0) = "BANK_CD"
    arrField(1) = "BANK_NM"

    arrHeader(0) = "송금은행"
    arrHeader(1) = "송금은행명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBank(arrRet)
	End If

End Function


Function SetBank(byval arrRet)
	frm1.txtBankCd.Value= arrRet(0)
	frm1.txtBankNm.Value= arrRet(1)
	lgBlnFlgChgValue = True
End Function

Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"
	arrParam(1) = "B_Pur_Grp"

	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)

	arrParam(4) = "USAGE_FLG='Y'"
	arrParam(5) = "구매그룹"

    arrField(0) = "PUR_GRP"
    arrField(1) = "PUR_GRP_NM"

    arrHeader(0) = "구매그룹"
    arrHeader(1) = "구매그룹명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGroup(arrRet)
	End If

End Function


Function SetGroup(byval arrRet)
	frm1.txtGroupCd.Value= arrRet(0)
	frm1.txtGroupNm.Value= arrRet(1)
	lgBlnFlgChgValue = True
End Function

Function OpenPayType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	if Trim(frm1.txtPayTermCd.Value) = "" then
		Call DisplayMsgBox("17a002", Parent.VB_YES_NO,"결제방법", "X")
		Exit Function
	End if

	If IsOpenPop = True Or UCase(frm1.txtPayTypeCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "지급유형"
	arrParam(1) = "B_MINOR,B_CONFIGURATION," _
	& "(Select REFERENCE From B_CONFIGURATION Where MAJOR_CD ='B9004'"_
		& "And MINOR_CD='" & FilterVar(Trim(frm1.txtPayTermCd.value), "", "SNM") & "' And SEQ_NO>=2)C"

	arrParam(2) = Trim(frm1.txtPayTypeCd.Value)

	arrParam(4) = "B_MINOR.MINOR_CD = C.REFERENCE And B_CONFIGURATION.MINOR_CD = B_MINOR.MINOR_CD And B_MINOR.MAJOR_CD = 'A1006' " _
				& "AND B_CONFIGURATION.REFERENCE IN('RP','P')"
	arrParam(5) ="지급유형"

	arrField(0) = "B_MINOR.MINOR_CD"
	arrField(1) = "B_MINOR.MINOR_NM"

    arrHeader(0) = "지급유형"
    arrHeader(1) = "지급유형명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPayTypeCd.Value = arrRet(0)
		frm1.txtPayTypeNm.Value = arrRet(1)
		lgBlnFlgChgValue 		= True
	End If
End Function

Function OpenPaymeth()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPayTermCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "결제방법"
	arrParam(1) = "B_Minor,b_configuration"

	arrParam(2) = Trim(frm1.txtPayTermCd.Value)
	'arrParam(3) = Trim(frm1.txtPayNm.Value)

	arrParam(4) = "b_minor.Major_Cd='B9004' and b_minor.minor_cd=b_configuration.minor_cd and b_minor.major_cd=b_configuration.major_cd and b_configuration.SEQ_NO=1"							<%' Where Condition%>
	arrParam(5) ="결제방법"

    arrField(0) = "b_minor.Minor_Cd"
    arrField(1) = "b_minor.Minor_Nm"
    arrField(2) = "b_configuration.REFERENCE"

    arrHeader(0) = "결제방법"
    arrHeader(1) = "결제방법명"
    arrHeader(2) = "Reference"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPayMeth(arrRet)
	End If
End Function

Function SetPaymeth(byval arrRet)
	frm1.txtPaytermCd.Value    = arrRet(0)
	frm1.txtPaytermNm.Value    = arrRet(1)
	frm1.txtReference.Value	   = arrRet(2)
	lgBlnFlgChgValue = True
	Call changePayterm()
End Function

Function OpenMinorCode(MajorCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strtitle

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	'공통부분 
	arrParam(1) = "B_Minor"

    arrField(0) = "Minor_Cd"
    arrField(1) = "Minor_Nm"
	arrParam(4) = "Major_Cd='" & MajorCode & "'"

    Select Case MajorCode
    Case "B9006"
		if frm1.txtIncotermsCd.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtIncotermsCd.value)
		arrParam(0) = "가격조건"
		arrParam(5) ="가격조건"
		arrHeader(0) = "가격조건"
		arrHeader(1) = "가격조건명"

	Case "B9009"
		if frm1.txtTransCd.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtTransCd.value)
		arrParam(0) = "운송방법"
		arrParam(5) ="운송방법"
		arrHeader(0) = "운송방법"
		arrHeader(1) = "운송방법명"

	Case "B9007"
		if frm1.txtPackingCd.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtPackingCd.value)
		arrParam(0) = "포장조건"
		arrParam(5) ="포장조건"
		arrHeader(0) = "포장조건"
		arrHeader(1) = "포장조건명"

	Case "B9008"
		if frm1.txtInspectCd.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtInspectCd.value)
		arrParam(0) = "검사방법"
		arrParam(5) ="검사방법"
		arrHeader(0) = "검사방법"
		arrHeader(1) = "검사방법명"

	Case "B9095"
		if frm1.txtDvryPlce.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtDvryPlce.value)
		arrParam(0) = "인도장소"
		arrParam(5) ="인도장소"
		arrHeader(0) = "인도장소"
		arrHeader(1) = "인도장소명"

	Case "B9094"
		if frm1.txtOrigin.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtOrigin.value)
		arrParam(0) = "원산지"
		arrParam(5) ="원산지"
		arrHeader(0) = "원산지"
		arrHeader(1) = "원산지명"

	Case "B9096"
		if frm1.txtDisCity.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtDisCity.value)
		arrParam(0) = "도착도시"
		arrParam(5) ="도착도시"
		arrHeader(0) = "도착도시"
		arrHeader(1) = "도착도시명"

	Case "B9092"
		if frm1.txtDisPort.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtDisPort.value)
		arrParam(0) = "도착항"
		arrParam(5) ="도착항"
		arrHeader(0) = "도착항"
		arrHeader(1) = "도착항명"

	Case "B9092-1"
		if frm1.txtLoadPort.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtLoadPort.value)
		arrParam(0) = "선적항"
		arrParam(5) ="선적항"
		arrHeader(0) = "선적항"
		arrHeader(1) = "선적항명"
		arrParam(4) = "Major_Cd='B9092'"

	Case "A1006"
		if frm1.txtPaytypecd.ReadOnly = true then
			IsOpenPop = False
			Exit Function
		End if
		arrParam(2) = Trim(frm1.txtPaytypeCd.value)
		arrParam(0) = "지급유형"
		arrParam(5) ="지급유형"
		arrHeader(0) = "지급유형"
		arrHeader(1) = "지급유형명"

    End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else

		Select Case MajorCode
			Case "B9006"
				frm1.txtIncotermsCd.Value   = arrRet(0)
				frm1.txtIncotermsNm.Value   = arrRet(1)
				lgBlnFlgChgValue 			= True

			Case "B9009"
				frm1.txtTransCd.Value    	= arrRet(0)
				frm1.txtTransNm.Value    	= arrRet(1)
				lgBlnFlgChgValue 			= True

			Case "B9007"
				frm1.txtPackingCd.Value    	= arrRet(0)
				frm1.txtPackingNm.Value    	= arrRet(1)
				lgBlnFlgChgValue 			= True
			Case "B9008"
				frm1.txtInspectCd.Value    	= arrRet(0)
				frm1.txtInspectNm.Value    	= arrRet(1)
				lgBlnFlgChgValue 			= True
			Case "B9095"
				frm1.txtDvryPlce.Value    	= arrRet(0)
				frm1.txtDvryPlceNm.Value    = arrRet(1)
				lgBlnFlgChgValue 			= True
			Case "B9094"
				frm1.txtOrigin.Value    	= arrRet(0)
				frm1.txtOriginNm.Value    	= arrRet(1)
				lgBlnFlgChgValue 			= True
			Case "B9096"
				frm1.txtDisCity.Value    	= arrRet(0)
				frm1.txtDisCityNm.Value    	= arrRet(1)
				lgBlnFlgChgValue 			= True
			Case "B9092"
				frm1.txtDisPort.Value    	= arrRet(0)
				frm1.txtDisPortNm.Value    	= arrRet(1)
				lgBlnFlgChgValue 			= True
			Case "B9092-1"
				frm1.txtLoadPort.Value	   	= arrRet(0)
				frm1.txtLoadPortNm.Value   	= arrRet(1)
				lgBlnFlgChgValue 			= True
			Case "A1006"
				frm1.txtPaytypeCd.Value	   	= arrRet(0)
				frm1.txtPaytypeNm.Value   	= arrRet(1)
				lgBlnFlgChgValue 			= True
		End Select

	End If
End Function

Function OpenBiz(strValue)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function


	'공통부분 
	arrParam(1) = "B_BIZ_PARTNER"

    arrField(0) = "BP_Cd"
    arrField(1) = "BP_Nm"


    Select Case strValue
    Case "Appl"
    	if UCase(frm1.txtApplicantCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
		arrParam(2) = Trim(frm1.txtApplicantCd.value)
		arrParam(0) = "수입자"
		arrParam(5) ="수입자"
		arrHeader(0) = "수입자"
		arrHeader(1) = "수입자명"

	Case "Manu"
		if UCase(frm1.txtManuCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
		arrParam(2) = Trim(frm1.txtManuCd.value)
		arrParam(0) = "제조자"
		arrParam(5) ="제조자"
		arrHeader(0) = "제조자"
		arrHeader(1) = "제조자명"

	Case "Agent"
		if UCase(frm1.txtAgentCd.className) = Ucase(Parent.UCN_PROTECTED) Then Exit Function
		arrParam(2) = Trim(frm1.txtAgentCd.value)
		arrParam(0) = "대행자"
		arrParam(5) ="대행자"
		arrHeader(0) = "대행자"
		arrHeader(1) = "대행자명"

    End Select

	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else

		Select Case strValue
		Case "Appl"
			frm1.txtApplicantCd.Value    = arrRet(0)
			frm1.txtApplicantNm.Value    = arrRet(1)
			lgBlnFlgChgValue = True

		Case "Manu"
			frm1.txtManuCd.Value    = arrRet(0)
			frm1.txtManuNm.Value    = arrRet(1)
			lgBlnFlgChgValue = True

		Case "Agent"
			frm1.txtAgentCd.Value    = arrRet(0)
			frm1.txtAgentNm.Value    = arrRet(1)
			lgBlnFlgChgValue = True

		End Select
	End If

End Function
'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++
'------------------------------------------  SetPlant()  --------------------------------------------------
'	Name : SetPlant()
'	Description : Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------


'------------------------------------------  SetCondPlant()  --------------------------------------------------
'	Name : SetCondPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------

'------------------------------------------  SetOpenCalType()  --------------------------------------------------
'	Name : SetOpenCalType()
'	Description : Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )

   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"Z"
        Case 6                                                              '과부족허용율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","999.9999"
    End Select

End Sub

'============================================  2.5.1 TotalSum()  ========================================
'=	Name : TotalSum()																					=
'=	Description :																						=
'========================================================================================================
Sub TotalSum(ByVal row)

    Dim SumTotal, lRow, tmpGrossAmt, tmpVatAmt,tmpamt,tmpVat, SumVat,SumVatTotal, SumGross
	SumTotal = 0
	ggoSpread.source = frm1.vspdData
	SumTotal = UNICDbl(frm1.txtDetailNetAmt.value)
	frm1.vspdData.Row = row
	frm1.vspdData.Col = C_NetAmt
	tmpGrossAmt = UNICDbl(frm1.vspdData.Text)

	frm1.vspdData.Col = 0

    frm1.vspdData.Col = C_OrgNetAmt
    SumTotal = SumTotal + (tmpGrossAmt - UNICDbl(frm1.vspdData.Text))

    frm1.txtDetailNetAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")

'#### Summing Vat Amount
    SumVatTotal = UNICDbl(frm1.txtDetailVatAmt.Text)
	frm1.vspdData.Row = row
	frm1.vspdData.Col = C_VatAmt
	tmpVat = UNICDbl(frm1.vspdData.Text)
	frm1.vspdData.Col = C_OrgVatAmt
	'SumVatTotal = SumVatTotal + (tmpVat - UNICDbl(frm1.vspdData.Text))
	SumVatTotal = SumVatTotal + tmpVat
    frm1.txtDetailVatAmt.Text = UNIConvNumPCToCompanyByCurrency(SumVatTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")


'#### Summing Gross Amount
    SumGross = uniCdbl(SumTotal) + uniCdbl(SumVatTotal)
    frm1.txtDetailGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumGross, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")

End Sub

'==========================================   ChangeItemPlant()  ======================================
'	Name : ChangeItemPlant()
'=========================================================================================================
Sub ChangeItemPlant(byVal intStartRow ,byVal IntEndRow)
    Err.Clear

	Dim strVal
    Dim intIndex
    Dim lGrpCnt
	Dim igColSep,igRowSep

	igColSep = Parent.gColSep
	igRowSep = Parent.gRowSep

	If Trim(frm1.txtMaintNo.Value) <> "" Then Exit Sub

    frm1.txtMode.Value = "LookUpItemPlant"
	lGrpCnt = 1
	strVal = ""

	For intIndex = intStartRow To intEndRow
		strVal = strVal & CStr(intIndex) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_SLCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OrderUnit,intIndex,"X","X")) & igRowSep

		lGrpCnt = lGrpCnt + 1

		Call frm1.vspdData.SetText(C_Cost	,	intIndex, "")
		Call frm1.vspdData.SetText(C_Over	,	intIndex, "")
		Call frm1.vspdData.SetText(C_Under	,	intIndex, "")
	Next

	frm1.txtMaxRows.value = lGrpCnt-1
	frm1.txtSpread.value = strVal

    If LayerShowHide(1) = False Then Exit Sub

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

End Sub

'==========================================   ChangeItemPlant2()  ======================================
'	Name : ChangeItemPlant2()
'	[2005/09/16 Sim Hae Young Add Sub]
'=========================================================================================================
Sub ChangeItemPlant2(lRow)

	Dim lgF2By2
	Dim arrVal1
	Dim arrVal2

	Dim iStrSelect
	Dim iStrSql

	Dim iOrderUnitArr
	Dim iOrderUnitArr2
	Dim iOrderUnitArr3
	Dim iSLCdArr
	Dim iSLNmArr
	Dim iItemNmArr
	Dim iSpecArr
	Dim iHSCdArr
	Dim iHSNmArr
	Dim iPlantNmArr
	Dim iProcur_type
	Dim iTracking_Flg
	Dim iUnder_Tol
	Dim iOver_Tol

	Err.Clear

	iStrSelect = ""
	iStrSelect = " B.PUR_UNIT, A.ORDER_UNIT_PUR, C.BASIC_UNIT, A.MAJOR_SL_CD, A.SL_NM, C.ITEM_NM, C.SPEC, C.HS_CD,C.HS_NM, D.PLANT_NM, A.PROCUR_TYPE, A.TRACKING_FLG,  "
	iStrSelect = iStrSelect & " B.UNDER_TOL, ISNULL(B.OVER_TOL, A.OVER_TOL) OVER_TOL  "

	iStrSql =""
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT S.ITEM_CD,S.ORDER_UNIT_PUR, S.MAJOR_SL_CD, S.PROCUR_TYPE, T.SL_NM, S.TRACKING_FLG, "
	iStrSql = iStrSql & " 			CASE WHEN S.OVER_RCPT_FLG = 'Y' THEN S.OVER_RCPT_RATE ELSE 0 END OVER_TOL "
	iStrSql = iStrSql & " 	FROM B_ITEM_BY_PLANT S LEFT OUTER JOIN B_STORAGE_LOCATION T ON(S.MAJOR_SL_CD=T.SL_CD AND S.PLANT_CD=T.PLANT_CD) "
	iStrSql = iStrSql & " WHERE S.PLANT_CD=" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND S.ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " )A "
	iStrSql = iStrSql & " LEFT OUTER JOIN  "
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT ITEM_CD,PUR_UNIT,UNDER_TOL,OVER_TOL "
	iStrSql = iStrSql & " 	FROM M_SUPPLIER_ITEM_BY_PLANT "
	iStrSql = iStrSql & " WHERE PLANT_CD=" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " 	AND BP_CD IN(SELECT BP_CD FROM M_PUR_ORD_HDR WHERE PO_NO=" & FilterVar(Trim(frm1.txtPoNo.value), "''" , "S") & ") "
	iStrSql = iStrSql & " )B "
	iStrSql = iStrSql & " ON(A.ITEM_CD=B.ITEM_CD)  "
	iStrSql = iStrSql & " LEFT OUTER JOIN  "
	iStrSql = iStrSql & " ( "
	iStrSql = iStrSql & " SELECT S.ITEM_CD,S.BASIC_UNIT,S.ITEM_NM, S.SPEC, S.HS_CD, T.HS_NM "
	iStrSql = iStrSql & " FROM B_ITEM S LEFT OUTER JOIN B_HS_CODE T ON(S.HS_CD=T.HS_CD) "
	iStrSql = iStrSql & " WHERE S.ITEM_CD =" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_ItemCd,lRow,"X","X")), "''" , "S")
	iStrSql = iStrSql & " )C "
	iStrSql = iStrSql & " ON(A.ITEM_CD=C.ITEM_CD),  "
	iStrSql = iStrSql & " (SELECT PLANT_NM FROM B_PLANT WHERE PLANT_CD=" & FilterVar(Trim(GetSpreadText(frm1.vspdData,C_PlantCd,lRow,"X","X")), "''" , "S") & ")D "



	If CommonQueryRs2by2(iStrSelect, iStrSql, , lgF2By2)= False Then
		Call DisplayMsgBox("122700","X","X","X")
		Err.Clear

		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = C_itemCd
		frm1.vspdData.text = ""
		Exit Sub
	End If

	arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))

	arrVal2 = Split(arrVal1(0), chr(11))

	iOrderUnitArr  	= Trim(arrVal2(1))
	iOrderUnitArr2	= Trim(arrVal2(2))
	iOrderUnitArr3	= Trim(arrVal2(3))
	iSLCdArr		= Trim(arrVal2(4))
	iSLNmArr		= Trim(arrVal2(5))
	iItemNmArr		= Trim(arrVal2(6))
	iSpecArr		= Trim(arrVal2(7))
	iHSCdArr		= Trim(arrVal2(8))
	iHSNmArr		= Trim(arrVal2(9))
	iPlantNmArr		= Trim(arrVal2(10))
	iProcur_type	= Trim(arrVal2(11))
	iTracking_Flg	= Trim(arrVal2(12))
	iUnder_Tol      = Trim(arrVal2(13))
    iOver_Tol       = Trim(arrVal2(14))
   

	With frm1.vspdData
		.Row = lRow

		.Col = C_OrderUnit
		If Trim(iOrderUnitArr)<>"" Then
			.text = Trim(iOrderUnitArr)
		Else
			If Trim(iOrderUnitArr2)<>"" Then
				.text = Trim(iOrderUnitArr2)
			Else
				.text = Trim(iOrderUnitArr3)
			End If
		End If

		'=============================
		'품목의 조달구분 체크 
		'=============================
		If (Trim(iProcur_type)="P") And (Trim(frm1.hdnSubcontraflg.Value) = "Y") then
			Call DisplayMsgBox("179019","X","X","X")
			.Col = C_itemCd
			.text = ""
			Exit Sub
		End If
		If (Trim(iProcur_type)<>"P") And (Trim(frm1.hdnSubcontraflg.Value) = "N") then
			Call DisplayMsgBox("179019","X","X","X")
			.Col = C_itemCd
			.text = ""
			Exit Sub
		End If


		.Col = C_SLCd
		.text = Trim(iSLCdArr)

		.Col = C_SLNm
		.text = Trim(iSLNmArr)

		.Col = C_itemNm
		.text = Trim(iItemNmArr)

		.Col = C_SpplSpec
		.text = Trim(iSpecArr)

		.Col = C_HSCd
		.text = Trim(iHSCdArr)

		.Col = C_HSNm
		.text = Trim(iHSNmArr)

		.Col = C_PlantNm
		.text = Trim(iPlantNmArr)
		
		.Col = C_PrNo

		If .text = "" Then		
			If iTracking_Flg <> "Y" Then
				ggoSpread.spreadlock C_TrackingNo, .Row, C_TrackingNoPop, .Row
				.Col = C_TrackingNo
				.text = "*"
			Else
				ggoSpread.spreadUnlock C_TrackingNo, .Row, C_TrackingNoPop, .Row
				ggoSpread.sssetrequired C_TrackingNo, .Row, .Row
				.Col = C_TrackingNo
				.text = ""
			End If
		End If
		
		'2006.12.8 Modified by KSJ
		.Col = C_Over
		.text = iOver_Tol
		
		.Col = C_Under
		.text = iUnder_Tol

	End With

End Sub

Sub changeItemPlantOK()

	if Trim(frm1.hdnTrackingflg.Value) = "*" then
		ggoSpread.spreadlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
	else
		ggoSpread.spreadUnlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
		ggoSpread.sssetrequired C_TrackingNo, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	end if

End Sub

'==========================================   ChangeItemPlant()  ======================================
'	Name : ChangeItemPlantForUnit()
'	Description : 단위변경시 
'=========================================================================================================

Sub ChangeItemPlantForUnit(byVal intStartRow ,byVal IntEndRow)

    Err.Clear

    Dim strVal
    Dim intIndex
    Dim lGrpCnt
	Dim igColSep,igRowSep

	igColSep = Parent.gColSep
	igRowSep = Parent.gRowSep

	If Trim(frm1.txtMaintNo.Value) <> "" Then Exit Sub

    frm1.txtMode.Value = "LookUpItemPlantForUnit"
	lGrpCnt = 1
	strVal = ""
	For intIndex = intStartRow To intEndRow
		strVal = strVal & CStr(intIndex) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_ItemCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_PlantCd,intIndex,"X","X")) & igColSep
		strVal = strVal & Trim(GetSpreadText(frm1.vspdData,C_OrderUnit,intIndex,"X","X")) & igRowSep

		lGrpCnt = lGrpCnt + 1
	Next

	frm1.txtMaxRows.value = lGrpCnt-1
	frm1.txtSpread.value = strVal

    If LayerShowHide(1) = False Then Exit Sub

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

End Sub

'==========================================   ChangeItemPlantForUnit2()  ======================================
'	Name : ChangeItemPlantForUnit2()
'	Description : 단위변경시 
'=========================================================================================================

Sub ChangeItemPlantForUnit2(byVal lRow)

	Dim strsstemp1,strsstemp2,strsstemp3
	Dim strWhere, strPriceType
    
    ggoSpread.Source = frm1.vspdData

    with frm1.vspdData 		
		.Row = lRow

		.Col 		= C_ItemCd
		strssTemp1 	= Trim(.Text)
		.Col 		= C_PlantCd
		strssTemp2 	= Trim(.Text)
		.Col 		= C_OrderUnit
		strssTemp3 	= Trim(.Text)
		
		If strssTemp1 = "" Or strssTemp2 = "" Or strssTemp3 = "" Then
			Exit Sub
		End if

		' 단가type 의 유무를 조사 
		Call CommonQueryRs(" MINOR_CD ", " B_CONFIGURATION ", " MAJOR_CD = " & FilterVar("M0001", "''", "S") & " AND REFERENCE = " & FilterVar("Y", "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Err.number <> 0 Then
			MsgBox Err.description, VbInformation, parent.gLogoName
			Err.Clear 
			Exit Sub
		End If
		
		If Len(lgF0) > 0 Then
			lgF0 = Split(lgF0, Chr(11))
			strPriceType = lgF0(0)
		Else
			Call DisplayMsgBox("171214", "X", "X", "X")
			Exit Sub
		End If
	
		strWhere = " PLANT_CD = " & FilterVar(strssTemp2, "''", "S")
		strWhere = strWhere & " AND ITEM_CD = " & FilterVar(strssTemp1, "''", "S")
		strWhere = strWhere & " AND BP_CD = " & FilterVar(Trim(frm1.txtSupplierCd.value), "''", "S")
		strWhere = strWhere & " AND PUR_UNIT = " & FilterVar(strssTemp3, "''", "S")
		strWhere = strWhere & " AND PUR_CUR = " & FilterVar(Trim(frm1.txtCurr.value), "''", "S")
		strWhere = strWhere & " AND VALID_FR_DT <= " & FilterVar(Trim(frm1.txtPoDt.text), "''", "S")
		If Trim(strPriceType) = "T" Then
			strWhere = strWhere & " AND PRC_FLG =  'T' "
		End If
		strWhere = strWhere & " ORDER BY VALID_FR_DT DESC "

		Call CommonQueryRs(" PUR_PRC, PRC_FLG ", " M_SUPPLIER_ITEM_PRICE ", strwhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		If Err.number <> 0 Then
			MsgBox Err.description, VbInformation, parent.gLogoName
			Err.Clear 
			Exit Sub
		End If
	
		If Len(lgF0) > 0 Then
			lgF0 = Split(lgF0, Chr(11))
			lgF1 = Split(lgF1, Chr(11))
			.Col = C_Cost
			.Text = lgF0(0)
			.Col = C_CostConCd
			.Text = lgF1(0)
		Else
			strWhere = " PLANT_CD = " & FilterVar(strssTemp2, "''", "S")
			strWhere = strWhere & " AND ITEM_CD = " & FilterVar(strssTemp1, "''", "S")
			strWhere = strWhere & " AND PUR_UNIT = " & FilterVar(strssTemp3, "''", "S")
			strWhere = strWhere & " AND PUR_CUR = " & FilterVar(Trim(frm1.txtCurr.value), "''", "S")
			strWhere = strWhere & " AND VALID_FR_DT <= " & FilterVar(Trim(frm1.txtPoDt.text), "''", "S")
			If Trim(strPriceType) = "T" Then
				strWhere = strWhere & " AND PRC_FLG =  'T' "
			End If
			strWhere = strWhere & " ORDER BY VALID_FR_DT DESC "
	
			Call CommonQueryRs(" PUR_PRC, PRC_FLG ", " M_ITEM_PUR_PRICE ", strwhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			If Err.number <> 0 Then
				MsgBox Err.description, VbInformation, parent.gLogoName
				Err.Clear 
				Exit Sub
			End If
	
			If Len(lgF0) > 0 Then
				lgF0 = Split(lgF0, Chr(11))
				lgF1 = Split(lgF1, Chr(11))
				.Col = C_Cost
				.Text = lgF0(0)
				.Col = C_CostConCd
				.Text = lgF1(0)
			Else
				.Col = C_Cost
				.Text = 0
			End If
		End If
		
		Call vspdData_Change(C_Cost, lRow)
		Call vspdData_Change(C_CostConCd, lRow)
	End With

End Sub

'==========================================   lookupPrice()  ======================================
'	Name : lookupPrice()
'	Description :
'==================================================================================================
Function lookupPrice(ByVal Row)

    Err.Clear

    If CheckRunningBizProcess = True Then
		Exit Function
	End If

    Dim strVal

	lgBlnFlgChgValue = true

	frm1.vspdData.Row = frm1.vspdData.ActiveRow

	' === 2005.07.15 단가관련 수정 =================
	frm1.vspdData.Col = C_ItemCd
	If Trim(frm1.vspdData.text) = "" Then
		Call DisplayMsgBox("169915","X","X","X")
		Call LayerShowHide(0)
		Exit Function
	End If
	' === 2005.07.15 단가관련 수정 =================


    strVal = BIZ_PGM_ID & "?txtMode=" & "lookupPrice"
    strVal = strVal & "&txtStampDt=" & Trim(frm1.hdnPoDt.value)
	''frm1.vspdData.Col = C_SupplierCd
    strVal = strVal & "&txtBpCd=" & Trim(frm1.hdnSupplierCd.Value)
	frm1.vspdData.Col = C_itemCd
    strVal = strVal & "&txtItemCd=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_PlantCd
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.vspdData.text)
	frm1.vspdData.Col = C_OrderUnit
    strVal = strVal & "&txtUnit=" & Trim(frm1.vspdData.text)
	''frm1.vspdData.Col = C_PoCurrency
    strVal = strVal & "&txtCurrency=" & Trim(frm1.hdnCurr.value)
    strVal = strVal & "&txtRow=" & Row
	'frm1.vspdData.Col = C_PoPrice2
	'frm1.vspdData.Text = 0

    If LayerShowHide(1) = False Then Exit Function

	Call RunMyBizASP(MyBizASP, strVal)

End Function
''==========================================   lookupPriceForSelection()  =============================
''	Name : lookupPriceForSelection()
''	Description :
''=====================================================================================================
'Function lookupPriceForSelection()
'
'    Err.Clear
'
'    If CheckRunningBizProcess = True Then
'		Exit Function
'	End If
'
'    Dim strVal
'
'	lgBlnFlgChgValue = true
'
'    If LayerShowHide(1) = False Then Exit Function
'
'	Call RunMyBizASP(MyBizASP, strVal)
'
'    Dim lRow
'    Dim lGrpCnt
'
'	With frm1
'
'    '-----------------------
'    'Data manipulate area
'    '-----------------------
'    lGrpCnt = 1
'    strVal = ""
'    '-----------------------
'    'Data manipulate area
'    '-----------------------
'	.txtMode.value = "lookupPriceForSelection"
'
'	For lRow = 1 To .vspdData.MaxRows
'
'		.vspdData.Row = lRow
'
'		frm1.vspdData.Row = lRow
'		strVal = strVal & Trim(frm1.hdnSupplierCd.Value) & Parent.gColSep
'		frm1.vspdData.Col = C_ItemCd
'		strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
'		frm1.vspdData.Col = C_PlantCd
'		strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
'		frm1.vspdData.Col = C_OrderUnit
'		strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
'		strVal = strVal & Trim(frm1.hdnCurr.value) & Parent.gColSep
'		strVal = strVal & lRow & Parent.gRowSep
'
'		lGrpCnt = lGrpCnt + 1
'	Next
'
'
'	if strVal <> "" then
'		If LayerShowHide(1) = False Then Exit Function
'
'		.hdnMaxRows.value = .vspdData.MaxRows
'		.txtMaxRows.value = lGrpCnt-1
'		.txtSpread.value = strVal
'
'		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
'	End if
'
'	End With
'
'End Function
'==========================================================================================
'   Event Name : ChangeReturnCost
'   Event Desc :
'==========================================================================================
Sub ChangeReturnCost()

Dim IntCol, IntRow
Dim strssTemp1,strssTemp2,strssTemp3

intCol = frm1.vspdData.ActiveCol - 1
intRow = frm1.vspdData.ActiveRow

	if IntCol = C_itemCd or IntCol = C_PlantCd or IntCol = C_OrderUnit then

		frm1.vspdData.Col = C_ItemCd
		strssTemp1 = Trim(frm1.vspdData.Text)
		frm1.vspdData.Col = C_PlantCd
		strssTemp2 = Trim(frm1.vspdData.Text)
		frm1.vspdData.Col = C_OrderUnit
		strssTemp3 = Trim(frm1.vspdData.Text)

		if strssTemp1 = "" or strssTemp2 = ""  then'or strssTemp3 = "" then
			Exit Sub
		End if

		if intCol = C_OrderUnit then
			'//Call ChangeItemPlantForUnit(IntRow,IntRow)
			Call ChangeItemPlantForUnit2(IntRow)
		else
			Call ChangeItemPlant2(IntRow)
		end if

	End if

End Sub

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtPoAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtGrossPoAmt,.txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec

		ggoOper.FormatFieldByObjectOfCur .txtDetailNetAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtDetailVatAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtDetailGrossAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtXch, .txtCurr.value, parent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec		
	End With

End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'단가 
		ggoSpread.SSSetFloatByCellOfCur C_Cost,-1, .txtCurr.value, Parent.ggUnitCostNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'금액 
		ggoSpread.SSSetFloatByCellOfCur C_OrderAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
		'VAT금액 
		ggoSpread.SSSetFloatByCellOfCur C_VatAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
        ggoSpread.SSSetFloatByCellOfCur C_NetAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
	    ggoSpread.SSSetFloatByCellOfCur C_OrgNetAmt,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
        ggoSpread.SSSetFloatByCellOfCur C_OrgNetAmt1,-1, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gComNum1000, Parent.gComNumDec
        
        
	End With

End Sub

'========================================================================================
' Function Name : ChangeCurr()
' Function Desc :
'========================================================================================
Sub ChangeCurr()
	if UCase(Trim(frm1.txtCurr.value)) = UCase(Parent.gCurrency) then
		frm1.txtXch.Text = 1
		Call ggoOper.SetReqAttr(frm1.txtXch,"Q")
		Call ggoOper.SetReqAttr(frm1.cboXchop,"Q")
	else

		Call ggoOper.SetReqAttr(frm1.txtXch,"D")
		Call ggoOper.SetReqAttr(frm1.cboXchop,"N")
		frm1.txtXch.Text = 0	
	end if

	if Trim(frm1.txtCurr.value) <> "" then
	'-- Issue #9739 By Byun Jee Hyun 2005-09-28
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()				'2005-05-16 화폐변경시 포멧정보 가져오도록 함 
     	
		
	  frm1.hdnCurr.value = frm1.txtCurr.value
	End  If  

'    Call CurFormatNumericOCX()


	lgBlnFlgChgValue= true
End Sub
'========================================================================================
' Function Name : changePayterm
' Function Desc :
'========================================================================================
Sub changePayterm()

	frm1.txtPayTypeCd.Value = ""
	frm1.txtPayTypeNm.Value = ""
	frm1.txtPayDur.Text = 0

End Sub

'========================================================================================
' Function Name : Sending
' Function Desc :
'========================================================================================
'Function Sending()
'
'    Err.Clear
'
'    Sending = False
'
'	If LayerShowHide(1) = False Then Exit Function
'
'    Dim strVal
'
'    strVal = BIZ_PGM_ID & "?txtMode=" & "SendingB2B"
'    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)
'
'	Call RunMyBizASP(MyBizASP, strVal)
'
'    Sending = True
'
'
'End Function

'Function SendingOK()
'	'msgbox "전송이 완료 되었습니다."
'End Function
'========================================================================================
' Function Name : OnLineQuery
' Function Desc : 주문서관리번호로 OnLine관련 조회 
'========================================================================================
 'Function OnLineQuery()
 '
 '   Err.Clear
 '
 '   OnLineQuery = False
 '
'	If LayerShowHide(1) = False Then Exit Function
'
'    Dim strVal
'
'    strVal = BIZ_OnLine_ID & "?txtMode=" & "OnLineLookUp"
'    strVal = strVal & "&txtMaintNo=" & Trim(frm1.txtMaintNo.value)
'
'	Call RunMyBizASP(MyBizASP, strVal)
'
'    OnLineQuery = True
'
'End Function
<%
'================================== =====================================================
' Function Name : InitCollectType
' Function Desc : 소비세유형코드/명/율 저장하기 
' 여기부터 키보드에서 소비세유형코드를 변경시 소비세유형명,소비세율,매입금액,NetAmount를 변경시키는 함수 
'========================================================================================
%>
Sub InitCollectType()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iCodeArr, iNameArr, iRateArr

    Err.Clear

	Call CommonQueryRs(" Minor.MINOR_CD,  Minor.MINOR_NM, Config.REFERENCE ", " B_MINOR Minor,B_CONFIGURATION Config ", " Minor.MAJOR_CD='B9001' And Config.MAJOR_CD = Minor.MAJOR_CD And Config.MINOR_CD = Minor.MINOR_CD And Config.SEQ_NO = 1 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iCodeArr = Split(lgF0, Chr(11))
    iNameArr = Split(lgF1, Chr(11))
    iRateArr = Split(lgF2, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description
		Err.Clear
		Exit Sub
	End If

	Redim arrCollectVatType(UBound(iCodeArr) - 1, 2)

	For i = 0 to UBound(iCodeArr) - 1
		arrCollectVatType(i, 0) = iCodeArr(i)
		arrCollectVatType(i, 1) = iNameArr(i)
		arrCollectVatType(i, 2) = iRateArr(i)
	Next
End Sub

'========================================================================================
' Function Name : GetCollectTypeRef
' Function Desc :
'========================================================================================
Sub GetCollectTypeRef(ByVal VatType, ByRef VatTypeNm, ByRef VatRate)

	Dim iCnt

	For iCnt = 0 To Ubound(arrCollectVatType)
		If arrCollectVatType(iCnt, 0) = UCASE(VatType) Then
			VatTypeNm = arrCollectVatType(iCnt, 1)
			VatRate   = arrCollectVatType(iCnt, 2)
			Exit Sub
		End If
	Next
	VatTypeNm = ""
	VatRate = ""
End Sub

'========================================================================================
' Function Name : SetVatType
' Function Desc :
'========================================================================================
Sub SetVatTypeHdr()
	Dim VatType, VatTypeNm, VatRate

	VatType = frm1.txtVattype.value

	Call InitCollectType
	Call GetCollectTypeRef(VatType, VatTypeNm, VatRate)

	frm1.txtVatTypeNm.value = VatTypeNm

	frm1.txtVatrt.text = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

	frm1.hdntxtVatType.value = VatType
	frm1.hdntxtVatTypeNm.value = VatTypeNm
	frm1.hdntxtVatrt.value = UNIFormatNumber(UNICDbl(VatRate), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)

	If lgReqRefChk then
		Call SetVatName()
	End If

End Sub
'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################


'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리 
'*********************************************************************************************************
'========================================================================================
' Function Name : vspdData_Click
' Function Desc :
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

	If lgIntFlgMode = Parent.OPMD_UMODE Then
		if Trim(frm1.hdnRelease.Value) = "N" then
			Call SetPopupMenuItemInf("1101111111")
		else
			Call SetPopupMenuItemInf("0000111111")
		end if

	Else
		'Call SetPopupMenuItemInf("0000111111")
		Call SetPopupMenuItemInf("1101111111")
	End If

	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If

		Exit Sub
	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc :
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc :
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc :
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	'------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc :
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   :
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   :
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
    Call CurFormatNumSprSheet()
    Call ggoSpread.ReOrderingSpreadData()
    Call SetSpreadLockAfterQuery()
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim strsstemp1,strsstemp2,strsstemp3
	Dim tmpVatAmt, tmpDocAmt
	Dim Qty, Price, DocAmt, VatAmt, VatRate, chk_vat_flg, orgNetAmt, chkState
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6,PhntmFlg,strWhere
	Dim iNameArr, strPlantCd

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    with frm1.vspdData
		.Row = Row
		.Col = 0

		if Trim(.Text) = ggoSpread.DeleteFlag  then
		    Exit Sub
		end if

		.Col = C_Stateflg:	.Row = Row
		chkState = .Text

		if Trim(.Text) = "" then
			.Text = "U"
		End if

		if chkState <> "R" then
			Select Case Col
				'공장 
				Case C_PlantCd			
					.Col	= C_ItemCd
					.text 	= ""
					
					.Col 	= C_ItemNM
					.text 	= ""
				'품목 
				Case C_ItemCd			
					.Col 		= C_ItemCd
					strssTemp1 	= Trim(.Text)
					.Col 		= C_PlantCd
					strssTemp2 	= Trim(.Text)
					
					If strssTemp1 = "" Or strssTemp2 = "" then
						ggoSpread.spreadlock C_TrackingNo, Row, C_TrackingNoPop, Row
						.Col 	= C_TrackingNo
						.Text   = ""
						Exit Sub
					End if

					strWhere = "ITEM_CD = "& "'" & FilterVar(strssTemp1, " " , "SNM") & "' "
	
					Call CommonQueryRs("phantom_flg", "B_ITEM", strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
					If Err.number = 0 Then
						If lgF0 <> "" then
							PhntmFlg = Split(lgF0, Chr(11))
	
							if PhntmFlg(0) = "Y" then
								Call DisplayMsgBox("179024","X","X","X")
								Call LayerShowHide(0)
								Exit Sub
							End if
						End If
					End If
	
					'Call ChangeItemPlant(Row,Row)
					Call ChangeItemPlant2(Row)
			    ' 단위    
			    Case C_OrderUnit					
					Call  ChangeItemPlantForUnit2(Row)
					
					
				' 납기일 
				Case C_DlvyDt
					.Col = C_DlvyDt
					strsstemp1 = .Text
					if strsstemp1 = "" then Exit Sub
					strsstemp2 = frm1.txtPoDt.text
					if UniConvDateToYYYYMMDD(strsstemp2,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(strsstemp1,Parent.gDateFormat,"") then
						Call DisplayMsgBox("970023", "X", "납기일", frm1.txtPoDt.Alt)
					end if
				' HS부호 
				Case C_HSCd
	    			Err.Clear
					
					.Col = C_HSCd
					Call CommonQueryRs(" HS_NM ", " B_HS_CODE ", " HS_CD = " & FilterVar(.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
					If Err.number <> 0 Then
						MsgBox Err.description, VbInformation, parent.gLogoName
						Err.Clear 
						Exit Sub
					End If
	
					.Col = C_HSNm
					If Len(lgF0) > 0 Then
						iNameArr = Split(lgF0, Chr(11))
						.Text = iNameArr(0)
					Else
						.Text = ""
						.Col = C_HSCd
						Call DisplayMsgBox("203227", "X", .Text, "X")
						.Text = ""
					End If
				' 반품유형 
				Case C_RetCd
					Call SetRetCd()
			End Select
		End If
		
		Select Case Col
			'발주수량, 단가 
			Case C_OrderQty, C_Cost
				.Col = C_OrderQty
				If Trim(.Text) = "" Or IsNull(.Text) then
					Qty = 0
				Else
					Qty = UNICDbl(.Text)
				End If
				
				.Col = C_Cost
				If Trim(.Text) = "" Or IsNull(.Text) then
					Price = 0
				Else
					Price = UNICDbl(.Text)
				End If
				
				DocAmt 	= Qty * Price

				.Col 	= C_OrderAmt		
				.Text 	= UNIConvNumPCToCompanyByCurrency(DocAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo,"X","X")
				
				Call InitData(Row)
		
		        Call vspdData_Change(C_OrderAmt, Row)
		
			' 단가구분 
			Case C_CostCon
				Call vspdData_ComboSelChange(C_CostCon, Row)	' Line 복사시 SelChange를 강제로 일어나게 한다.

			' 금액 
			Case C_OrderAmt
				.Col = C_OrderAmt
			    DocAmt = UNICDbl(.Text)
			     
			    'VAT 금액 추가  -->
				.Col = C_VatRate ' VAT 율 
				If Trim(.Text) = "" OR IsNull(.Text) then
					VatRate = 0
				Else
					VatRate = UNICDbl(.Text)
				End If
		
				' 부가세 포함/불포함 부가세 계산 변경 2002.3.9 L.I.P
				.Col = C_IOFlgCD
				chk_vat_flg	= .text

				if chk_vat_flg = "2"	Then	'포함 
					VatAmt    = DocAmt * (VatRate/(100+VatRate))
				Else                            '별도 
					VatAmt    = DocAmt * (VatRate/100)
				End If
		
				.Col = C_VatAmt
				.Text = UNIConvNumPCToCompanyByCurrency(VatAmt,frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")		
				VatAmt = UNICDbl(.Text)

				' 통합 
				.Col = C_OrgVatAmt
				tmpVatAmt = UNICDbl(frm1.txtDetailVatAmt.Text) - UNICDbl(.Text) + UNICDbl(VatAmt)
				frm1.txtDetailVatAmt.Text = UNIConvNumPCToCompanyByCurrency(tmpVatAmt,frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")

				'VAT금액부터 계산후 발주순금액을 계한한다.(금액-VAT금액(함수 적용한 금액))
				if chk_vat_flg = "2"	Then	'포함 
					.Col = C_NetAmt		
				    .Text = UNIConvNumPCToCompanyByCurrency(DocAmt - VatAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo,"X","X")
				    orgNetAmt = .Text

					' 통합 
				    .Col = C_OrgNetAmt1
				    tmpDocAmt = UNICDbl(frm1.txtDetailNetAmt.Text) - UNICDbl(.Text) + UNICDbl(orgNetAmt)
				    frm1.txtDetailNetAmt.Text = UNIConvNumPCToCompanyByCurrency(tmpDocAmt,frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
				Else                            '별도 
					.Col = C_NetAmt		
				    .Text = UNIConvNumPCToCompanyByCurrency(DocAmt, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo,"X","X")
				    orgNetAmt = .Text

					' 통합 
				    .Col = C_OrgNetAmt1
				    tmpDocAmt = UNICDbl(frm1.txtDetailNetAmt.Text) - UNICDbl(.Text) + UNICDbl(orgNetAmt)
				    frm1.txtDetailNetAmt.Text = UNIConvNumPCToCompanyByCurrency(tmpDocAmt,frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
				End If
						
				' 통합 
				frm1.txtDetailGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(tmpDocAmt+tmpVatAmt,frm1.txtCurr.value,Parent.ggAmtOfMoneyNo,Parent.gTaxRndPolicyNo,"X")
				'<-- VAT 금액 추가 2002.2.18 L.I.P
				'Call TotalSum(Row)					'총품목금액합계 
		
				.Col 	= C_OrgNetAmt1
				.Text 	= orgNetAmt
		
				.Col 	= C_OrgVatAmt
				.Text 	= VatAmt
			' VAT포함여부 
			Case C_IOFlg
				.Col = C_IOFlg
				Call vspdData_ComboSelChange(C_IOFlg, Row)	' Line 복사시 SelChange를 강제로 일어나게 한다.
				Call vspdData_Change(C_OrderAmt, Row)
				Call setCVatFlg(Row)	
			' VAT
			Case C_VatType 'or Col = C_VatAmt then
				Call SetVatType(Row)     ' C_VatNm,C_VatRate 세팅 
				call vspdData_Change(C_OrderAmt, Row)
			' 창고 
			Case C_SLCd
    			Err.Clear
				.Col = C_PlantCd
				strPlantCd = .Text
				.Col = C_SLCd
				Call CommonQueryRs(" SL_NM ", " B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(.Text, "''", "S") & " AND PLANT_CD = " & FilterVar(strPlantCd, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

				If Err.number <> 0 Then
					MsgBox Err.description, VbInformation, parent.gLogoName
					Err.Clear 
					Exit Sub
				End If

				.Col = C_SLNm
				If Len(lgF0) > 0 Then
					iNameArr = Split(lgF0, Chr(11))
					.Text = iNameArr(0)
				Else
					.Text = ""
					.Col = C_SLCd
					.Text = ""
					Call DisplayMsgBox("169922", "X", "X", "X")
				End If
		End Select

    End With


	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

    Call CheckMinNumSpread(frm1.vspdData, Col, Row)

End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
Dim intIndex

	With frm1.vspdData

		.Row = Row
		.Col = Col

		if Col = C_CostCon then
				intIndex = .Value
				.Col = C_CostCon+1
				.Value = intIndex
		else
		        intIndex = .Value
				.Col = C_IOFlg+1
				.Value = intIndex
        end if

  End With

End Sub
'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
Dim strTemp, strRow
Dim intPos1

	With frm1.vspdData

    ggoSpread.Source = frm1.vspdData

    If Row > 0 Then
        .Col = Col
        .Row = Row
         strRow = Row

		Select Case Col

		Case C_Popup1
			Call OpenPlant()
		Case C_Popup2
			Call OpenItem()
		Case C_Popup3
			Call OpenUnit()
		Case C_Check
			Call lookupPrice(strRow)
		Case C_Popup5
			Call OpenHS()
		Case C_Popup6
			Call OpenSL()
		Case C_TrackingNoPop
			Call OpenTrackingNo()
		case C_Popup7
			Call OpenVat(2)
		case C_Popup8
		    Call OpenRet()
		End Select

    End If

    End With
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    With frm1.vspdData

    If Row >= NewRow Then
        Exit Sub
    End If

    If NewRow = .MaxRows Then
        'DbQuery
    End if

    End With

End Sub


'================ vspdData_TopLeftChange() ==========================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)

	If OldLeft <> NewLeft Then
	    Exit Sub
	End If


	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
 Sub Form_Load()

    Call LoadInfTB19029
    Call AppendNumberRange("0","0","999")

'    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'    Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'    Call ggoOper.LockField(Document, "N")

    call initFormatField()
    Call InitComboBox
    Call GetValue_ko441()
    Call SetDefaultVal
    Call InitVariables
    Call InitSpreadSheet
    '----------  Coding part  -------------------------------------------------------------
    'Call ggoOper.FormatNumber(frm1.txtPayDur,99,0)
    'Call ggoOper.FormatNumber(frm1.txtXch,"99999999","0",true,ggExchRate.DecPoint,Parent.gComNumDec,Parent.gComNum1000)
    Call Changeflg
    'Call SetToolbar("1110100000001111")
    Call CookiePage(0)

	Call changeTabs(TAB1)

	' === 2005.07.15 단가 일괄불러오기 관련 수정 =======
	Call SetPriceType
	' === 2005.07.15 단가 일괄불러오기 관련 수정 =======


    'sgbox parent.gpurgrp
	gIsTab     = "Y"
	gTabMaxCnt = 2

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'==========================================================================================
'   Event Name : OCX Event
'   Event Desc :
'==========================================================================================
 Sub txtPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoDt.Action = 7
	End if
End Sub

 Sub txtPoDt_Change()
	lgBlnFlgChgValue = true
	frm1.hdnPoDt.value = frm1.txtPoDt.text
 End Sub

 Sub txtOffDt_DblClick(Button)
	if Button = 1 then
		frm1.txtOffDt.Action = 7
	End if
End Sub

 Sub txtOffDt_Change()
	lgBlnFlgChgValue = true
End Sub
 Sub txtXch_Change()
	lgBlnFlgChgValue = true
End Sub
 Sub txtPayDur_Change()
	lgBlnFlgChgValue = true
End Sub
 Sub txtDvryDt_DblClick(Button)
	if Button = 1 then
		frm1.txtDvryDt.Action = 7
	End if
End Sub
 Sub txtDvryDt_Change()
	lgBlnFlgChgValue = true
End Sub
 Sub txtExpiryDt_DblClick(Button)
	if Button = 1 then
		frm1.txtExpiryDt.Action = 7
	End if
End Sub
 Sub txtExpiryDt_Change()
	lgBlnFlgChgValue = true
End Sub
Sub rdoVatFlg1_OnClick()
	lgBlnFlgChgValue = true
End Sub

Sub rdoVatFlg2_OnClick()
	lgBlnFlgChgValue = true
End Sub
'==========================================================================================
'   Event Name : txtVat_Type_OnChange
'   Event Desc :
'==========================================================================================
Sub txtVattype_OnChange()
	Call SetVatTypeHdr()
End Sub

'==========================================================================================
'   Event Name : txtVat_Type_OnChange
'   Event Desc : 수주형태별로 무역정보 필수입력 처리 
'==========================================================================================
Sub cboXchop_OnChange()
	lgBlnFlgChgValue = True

	if frm1.cboXchop.value ="*" then
		frm1.hdnxchrateop.value = "*"
	Else
		frm1.hdnxchrateop.value = "/"
	End if

End Sub

'--------------------------------------------------------------------
'		Name        : SetState()
'		Description : Spread의 Row상태를 "R","C"로 Setting
'					  R-reference 참조      C-InsertRow
'--------------------------------------------------------------------

Sub SetState(byval strState,byval IRow)
	frm1.vspdData.Row=IRow
	frm1.vspdData.Col=C_Stateflg
	frm1.vspdData.Text=strState
End Sub

Sub setVatAmt()
 dim sum

 with frm1
     sum = UNICDbl(.txtVatrt.text) * UNICDbl(.txtPoAmt.text)/100
     '.txtVatAmt.text = UNIFormatNumber(UNICDbl(sum), ggExchRate.DecPoint, -2, 0, ggExchRate.RndPolicy, ggExchRate.RndUnit)
     '.txtVatAmt.text = uniFormatNumberByTax(UNICDbl(sum),.txtCurr.value,Parent.ggAmtOfMoneyNo)'vatloc 라운딩 

 end with
end sub
'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다.
'	      Toolbar의 위치순서대로 기술하는 것으로 한다.
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'*********************************************************************************************************
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery()
    Dim IntRetCD

    FncQuery = False
    Err.Clear

    If lgBlnFlgChgValue = True Then

		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call InitVariables

	If Not chkFieldByCell(frm1.txtPoNo, "A",1)	then
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
        End If

        Exit Function
    End If

    If DbQuery = False Then Exit Function
    Call Changeflg

 '   lgBlnFlgChgValue = False
    FncQuery = True


End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew()
    Dim IntRetCD

    FncNew = False

    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ClickTab1()
    Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")
    Call ggoOper.ClearField(Document, "3")
    Call ggoOper.LockField(Document, "N")
    Call ChangeTag(False)
    Call SetDefaultVal
    Call InitVariables


    frm1.txtPoNo.focus
	Set gActiveElement = document.activeElement

    FncNew = True

End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete()

	Dim IntRetCD,lRow

    FncDelete = False

    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")

    If IntRetCD = vbNo Then Exit Function

    'If lgIntFlgMode <> Parent.OPMD_UMODE Then
     '   Call DisplayMsgBox("900002", "X", "X", "X")
     '   Exit Function
    'End If

    'if frm1.vspdData.Maxrows < 1	then exit function

    With frm1.vspdData

	'	.focus
		 ggoSpread.Source = frm1.vspdData

		 For lRow = 1 To .MaxRows step 1
		    .Row  = lRow
	       	.Col  = 0
			.Text = ggoSpread.DeleteFlag
		 Next
		'lDelRows = ggoSpread.DeleteRow
    End With
    If DbDelete = False Then Exit Function

    FncDelete = True

End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Save Button of Main ToolBar
'========================================================================================
Function FncSave()

    Dim IntRetCD
    Dim ReleaseChk
    Dim chkSts

    chkSts  = "DB"

    FncSave = False

    Err.Clear

    If CheckRunningBizProcess = True Then
		Exit Function
	End If

    If (lgBlnFlgChgValue = False) And (ggospread.SSCheckChange = False) Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If

    '2007.2 패치 환율 입력필수 삭제- KSJ
    'If frm1.txtxch.text = 0 Then
 	'    IntRetCD =  DisplayMsgBox("200095", "X", "X", "X")
	'    Call ClickTab1
	'    frm1.txtxch.focus
	'    Exit Function
    'End If
    '2007.2 패치End 환율 입력필수 삭제- KSJ

	IF frm1.hdnImportflg.value="Y" then

	    If Not chkEachFieldDomestic() Then
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        	Call BtnToolCtrl(chkSts)
	        Exit Function
	    End If

	    If Not chkEachFieldImport() Then
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        	Call BtnToolCtrl(chkSts)
	        Exit Function
	    End If

	else
		If Not chkEachFieldDomestic() Then
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        	Call BtnToolCtrl(chkSts)
	        Exit Function
	    End If

	End if


	'# Tab 1
	If Not chkFieldLengthByCell (frm1.txtSuppSalePrsn,"A",1) then
		Exit Function
	End If

	If Not chkFieldLengthByCell (frm1.txtTel,"A",1) then
		Exit Function
	End If

	If Not chkFieldLengthByCell (frm1.txtPayTermstxt,"A",1) then
		Exit Function
	End If

	If Not chkFieldLengthByCell (frm1.txtRemark,"A",1) then
		Exit Function
	End If


	'# Tab 2
    ggoSpread.Source = frm1.vspdData

	If Not chkField(Document, "2") OR Not ggoSpread.SSDefaultCheck Then
		Exit Function
	End If


	'# Tab 3
	If Not chkFieldLengthByCell (frm1.txtInvNo,"A",3) then
		Exit Function
	End If

	If Not chkFieldLengthByCell (frm1.txtShipment,"A",3) then
		Exit Function
	End If

	if frm1.rdoVatFlg1.checked = true then
    	frm1.hdvatFlg.Value = "1"	'별도 
    else
    	frm1.hdvatFlg.Value = "2"	'포함 
    End if

    Call Changeflg


	If DbSave("ToolBar") = False Then Exit Function

    FncSave = True

End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy()
	Dim IntRetCD,SumTotal,tmpGrossAmt,SumVatTotal,tmpVatAmt

	If gSelframeFlg = TAB1 Then
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If

	    lgIntFlgMode = Parent.OPMD_CMODE	'헤더 

	    <% ' 조건부 필드를 삭제한다. %>
	    Call ggoOper.ClearField(Document, "1")
	    Call ggoOper.LockField(Document, "N")
	    Call Changeflg
	    Call ChangeTag(False)

	    frm1.rdoRelease(0).checked = true
	 '   Call SetToolbar("11101000000111")

	'	frm1.txtPoAmt.Text		= UniNumClientFormat(0,ggAmtOfMoney.DecPoint,0)
	'	frm1.txtPoLocAmt.Text	= UniNumClientFormat(0,ggAmtOfMoney.DecPoint,0)
	'	frm1.txtVatAmt.Text		= UniNumClientFormat(0,ggAmtOfMoney.DecPoint,0)
		frm1.txtPoAmt.Text		= 0
		frm1.txtPoLocAmt.Text	= 0
		frm1.txtVatAmt.Text		= 0
		frm1.txtPoNo2.value = ""
		frm1.btnCfm.disabled = True
	    frm1.btnSend.disabled = True

	    lgBlnFlgChgValue = True
	Else

		if frm1.vspdData.Maxrows < 1	then exit function
		ggoSpread.Source = frm1.vspdData

		ggoSpread.CopyRow

		frm1.vspdData.ReDraw = False

		Call SetSpreadColor(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)

		frm1.vspdData.ReDraw = True

		Call SetState("C",frm1.vspdData.ActiveRow)

		'복사한 것은 긴급발주로 간주.
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_SeqNo
		frm1.vspdData.Text = ""

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_PrNo
		frm1.vspdData.Text = ""

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_MvmtNo
		frm1.vspdData.Text = ""

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_SoNo
		frm1.vspdData.Text = ""

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_SoSeqNo
		frm1.vspdData.Text = ""

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_TrackingNo

 		if Trim(frm1.vspdData.Text) = "*" then
			ggoSpread.spreadlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
		else
		    ggoSpread.spreadUnlock C_TrackingNo, frm1.vspdData.ActiveRow, C_TrackingNoPop, frm1.vspdData.ActiveRow
			ggoSpread.sssetrequired C_TrackingNo, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
		end if

		frm1.vspdData.ReDraw = True
		 '총품목금액합계 
		SumTotal = UNICDbl(frm1.txtDetailNetAmt.value)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_NetAmt
		tmpGrossAmt = UNICDbl(frm1.vspdData.Text)
		SumTotal = SumTotal + tmpGrossAmt
		frm1.txtDetailNetAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")

		'총Vat금액 합계 
		SumVatTotal = UNICDbl(frm1.txtDetailVatAmt.value)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = C_VatAmt
		tmpVatAmt = UNICDbl(frm1.vspdData.Text)
		SumVatTotal = SumVatTotal + tmpVatAmt
		frm1.txtDetailVatAmt.Text = UNIConvNumPCToCompanyByCurrency(SumVatTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
		frm1.txtDetailGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal+SumVatTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
    End if

		Dim iRow
		If lgPLCd <> "" then 
	    ggoSpread.SSSetProtected	C_PlantCd	, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	    ggoSpread.SSSetProtected	C_Popup1	, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
	    For iRow=frm1.vspdData.ActiveRow To frm1.vspdData.ActiveRow
	    	Call frm1.vspddata.SetText(C_PlantCd,iRow,lgPLCd)
	  	Next
		End If

End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()

    Dim maxrow,maxrow1,SumTotal,tmpGrossAmt,index,index1,orgtmpGrossAmt
    Dim SumVatTotal, tmpVatAmt, orgtmpVatAmt
	Dim starindex ,endindex,delflag

	if frm1.vspdData.Maxrows < 1	then exit function

	maxrow = frm1.vspdData.Maxrows
	index1 = 0

	starindex = frm1.vspdData.SelBlockRow
	endindex  = frm1.vspdData.SelBlockRow2

    Redim orgtmpGrossAmt(endindex - starindex)
    Redim tmpGrossAmt(endindex - starindex)
    Redim tmpVatAmt(endindex - starindex)
    Redim orgtmpVatAmt(endindex - starindex)
    Redim delflag(endindex - starindex)

    SumTotal	= UNICDbl(frm1.txtDetailNetAmt.value)
	SumVatTotal = UNICDbl(frm1.txtDetailVatAmt.value)

	for index = starindex to endindex
		frm1.vspdData.Row = index

	    frm1.vspdData.Col = C_NetAmt							'화면의 발주순금액 
	    tmpGrossAmt(index1) = UNICDbl(frm1.vspdData.Text)

	    frm1.vspdData.Col = C_OrgNetAmt1						'원래 발주순금액 
	    orgtmpGrossAmt(index1) = UNICDbl(frm1.vspdData.value)

	    frm1.vspdData.Col = C_VatAmt							 '화면의 vat금액 
	    tmpVatAmt(index1) = UNICDbl(frm1.vspdData.Text)

	    frm1.vspdData.Col = C_OrgVatAmt							'원래 vat금액 
	    orgtmpVatAmt(index1) = UNICDbl(frm1.vspdData.value)

	    frm1.vspdData.Col = 0
	    delflag(index1) = frm1.vspdData.Text
	    index1 = index1 + 1
	next

	ggoSpread.Source = frm1.vspdData
	index = frm1.vspdData.ActiveRow - starindex

    '//for index = 0 to index1 - 1
     if delflag(index) = ggoSpread.UpdateFlag then
            SumTotal	= SumTotal		+ (orgtmpGrossAmt(index) - tmpGrossAmt(index) )
            SumVatTotal = SumVatTotal	+ (orgtmpVatAmt(index) - tmpVatAmt(index) )
     elseif  delflag(index) = ggoSpread.DeleteFlag then
            SumTotal	= SumTotal		+ orgtmpGrossAmt(index)
            SumVatTotal = SumVatTotal	+ orgtmpVatAmt(index)
     elseif delflag(index) = ggoSpread.InsertFlag  then
            SumTotal = SumTotal - tmpGrossAmt(index)
            SumVatTotal = SumVatTotal - tmpVatAmt(index)
     end if
    '//Next

     ggoSpread.EditUndo
     maxrow1 = frm1.vspdData.Maxrows

     frm1.txtDetailNetAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
     frm1.txtDetailVatAmt.Text = UNIConvNumPCToCompanyByCurrency(SumVatTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
     frm1.txtDetailGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal+SumVatTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim inti
    inti=1

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then Exit Function
	End If

	With frm1

        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow

        For inti= .vspdData.ActiveRow  to .vspdData.ActiveRow +imRow-1
			.Row=inti
			ggoSpread.SetCombo "가단가" & vbtab & "진단가",C_CostCon
			ggoSpread.SetCombo "F" & vbtab & "T",C_CostConCd
			ggoSpread.SetCombo "포함" & vbtab & "별도",C_IOFlg
			ggoSpread.SetCombo "2" & vbtab & "1",C_IOFlgCd
			Call SetState("C",inti)

			'공장 기본값 추가 
			Call .vspdData.SetText(C_PlantCd,	inti, Parent.gPlant)
			Call .vspdData.SetText(C_PlantNm,	inti, Parent.gPlantNm)
			Call .vspdData.SetText(C_OrderAmt,	inti, "0")
			Call .vspdData.SetText(C_Cost,		inti, "0")
			Call .vspdData.SetText(C_DlvyDT,	inti, .txtDvryDt.Text)

			'Insert Row 시 헤더의 부가세관련 정보 초기값으로 2002.2.19
			if Trim(.txtVattype.value) = "" then
			    Call .vspdData.SetText(C_VatType,	inti, .hdntxtVatType.value)
			Else
			    Call .vspdData.SetText(C_VatType,	inti, .txtVattype.value)
			End if

			If .rdoVatFlg1.checked = False Then	'포함 
				Call .vspdData.SetText(C_IOFlg,		inti, 0)
				Call .vspdData.SetText(C_IOFlgCd,	inti, 0)
			Else
				Call .vspdData.SetText(C_IOFlg,		inti, 1)
				Call .vspdData.SetText(C_IOFlgCd,	inti, 1)
			End If

			if Trim(.hdntxtVatTypeNm.value) = "" then
				call SetVatType(inti)
			else
			    Call .vspdData.SetText(C_VatNm,	inti, .hdntxtVatTypeNm.value)
			end if

			if frm1.txtVatrt.Text = "" then
			    Call .vspdData.SetText(C_VatRate, inti, .hdntxtVatrt.values)
			else
			    Call .vspdData.SetText(C_VatRate, inti, .txtVatrt.Text)
			end if

			Call .vspdData.SetText(C_TrackingNo,	inti, "*")

			'---------------------------------------------------------
			'ggoSpread.spreadUnlock	C_PlantCd,.vspdData.Row,C_PlantCd,.vspdData.Row
			'ggoSpread.sssetrequired	C_PlantCd,.vspdData.Row,.vspdData.Row
			'ggoSpread.spreadUnlock	C_Popup1,.vspdData.Row,C_Popup1,.vspdData.Row
			'ggoSpread.spreadUnlock	C_ItemCd,.vspdData.Row,C_ItemCd,.vspdData.Row
			'ggoSpread.sssetrequired	C_ItemCd,.vspdData.Row,.vspdData.Row
			'ggoSpread.spreadUnlock	C_Popup2,.vspdData.Row,C_Popup2,.vspdData.Row
			'ggoSpread.spreadUnlock	C_IOFlg,.vspdData.Row,C_IOFlg,.vspdData.Row
			'ggoSpread.sssetrequired	C_IOFlg,.vspdData.Row,.vspdData.Row

			'If .hdnImportflg.value = "Y" Then
			'	ggoSpread.spreadUnlock	C_HsCd	,.vspdData.Row,C_Popup5	,.vspdData.Row
			'	ggoSpread.spreadUnlock	C_Popup5,.vspdData.Row,C_Popup5	,.vspdData.Row
			'	ggoSpread.sssetrequired	C_HsCd	,.vspdData.Row,.vspdData.Row
			'End If

			Call .vspdData.SetText(C_CostCon,	inti, 1)
			Call .vspdData.SetText(C_CostConCd,	inti, 1)
		Next

		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		'---------------------------------------------------------
		ggoSpread.spreadUnlock	C_PlantCd,.vspdData.ActiveRow,C_PlantCd,.vspdData.ActiveRow + imRow - 1
		ggoSpread.sssetrequired	C_PlantCd,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
		ggoSpread.spreadUnlock	C_Popup1,.vspdData.ActiveRow,C_Popup1,.vspdData.ActiveRow + imRow - 1
		ggoSpread.spreadUnlock	C_ItemCd,.vspdData.ActiveRow,C_ItemCd,.vspdData.ActiveRow + imRow - 1
		ggoSpread.sssetrequired	C_ItemCd,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
		ggoSpread.spreadUnlock	C_Popup2,.vspdData.ActiveRow,C_Popup2,.vspdData.ActiveRow + imRow - 1
		ggoSpread.spreadUnlock	C_IOFlg,.vspdData.ActiveRow,C_IOFlg,.vspdData.ActiveRow + imRow - 1
		ggoSpread.sssetrequired	C_IOFlg,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1

		If .hdnImportflg.value = "Y" Then
			ggoSpread.spreadUnlock	C_HsCd	,.vspdData.ActiveRow,C_Popup5	,.vspdData.ActiveRow + imRow - 1
			ggoSpread.spreadUnlock	C_Popup5,.vspdData.ActiveRow,C_Popup5	,.vspdData.ActiveRow + imRow - 1
			ggoSpread.sssetrequired	C_HsCd	,.vspdData.ActiveRow,.vspdData.ActiveRow + imRow - 1
		End If

		Dim iRow
		If lgPLCd <> "" then 
	    ggoSpread.SSSetProtected	C_PlantCd	, .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	    ggoSpread.SSSetProtected	C_Popup1	, .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	    For iRow=.vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
	    	Call frm1.vspddata.SetText(C_PlantCd,iRow,lgPLCd)
	  	Next
		End If

        .vspdData.ReDraw = True
    End With

	If Err.number = 0 Then FncInsertRow = True                                                          '☜: Processing is OK

    'frm1.btnSel.disabled = False
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow()
    Dim lDelRows
    Dim iDelRowCnt, i
    Dim index,SumTotal,SumVatTotal,idel
    if frm1.vspdData.Maxrows < 1	then exit function

    With frm1.vspdData

		.focus
		ggoSpread.Source = frm1.vspdData

		lDelRows = ggoSpread.DeleteRow

		SumTotal = UNICDbl(frm1.txtDetailNetAmt.value)
		SumVatTotal = UNICDbl(frm1.txtDetailVatAmt.value)

		for index = .SelBlockRow to .SelBlockRow2
			.Row = index
			.Col = C_Stateflg
			idel = .text
			.Col = 0

			if Trim(.text) <> ggoSpread.InsertFlag and Trim(idel) <> "D" then
			    .Col = C_NetAmt
		         SumTotal = SumTotal - UNICDbl(.Text)

		         .Col = C_VatAmt
		         SumVatTotal = SumVatTotal - UNICDbl(.Text)

		         .Col = C_Stateflg
			     frm1.vspdData.text = "D"
		         frm1.txtDetailNetAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
		         frm1.txtDetailVatAmt.Text = UNIConvNumPCToCompanyByCurrency(SumVatTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
		         frm1.txtDetailGrossAmt.Text = UNIConvNumPCToCompanyByCurrency(SumTotal + SumVatTotal, frm1.txtCurr.value, Parent.ggAmtOfMoneyNo, "X" , "X")
		    end if
		Next
   End With
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
	Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev()
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================
Function FncExcel()
    Call parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
    Call parent.FncFind(Parent.C_SINGLE , False)
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
 Function DbDelete()
    Err.Clear

    DbDelete = False

    Dim strVal
    Dim lRow  ,strDel ,lGrpCnt, ColSep, RowSep

	ColSep = Parent.gColSep
	RowSep = Parent.gRowSep
	strDel  = ""
	lGrpCnt = 0
	With frm1
		ggoSpread.Source = frm1.vspdData
	'	msgbox  frm1.vspdData.MaxRows
		For lRow = 1 To .vspdData.MaxRows step 1
			.vspdData.Row = lRow
			frm1.vspdData.Col = 0
			Select Case .vspdData.Text
				Case ggoSpread.DeleteFlag
				strDel = strDel & "D" & ColSep

				.vspdData.Col = C_SeqNo
				strDel = strDel & Trim(.vspdData.Text) & ColSep

				.vspdData.Col = C_PrNo
				strDel = strDel & Trim(.vspdData.Text) & ColSep

				strDel = strDel & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep
				strDel = strDel & lRow & RowSep

			lGrpCnt = lGrpCnt + 1
			End Select
		Next
		frm1.txtSpread.value  = strDel
		frm1.txtMaxRows.value = lGrpCnt
	End With

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
	strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)
	strVal = strVal & "&txtSpread=" & frm1.txtSpread.value
	strVal = strVal & "&txtMaxRows=" & frm1.txtMaxRows.value

    If LayerShowHide(1) = False Then Exit Function

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()
	lgBlnFlgChgValue = False
	Call MainNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
 Function DbQuery()

    Err.Clear

    DbQuery = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)
    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows

    frm1.hdnMaxRows.value = frm1.vspdData.MaxRows
    If LayerShowHide(1) = False Then Exit Function

	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
	Dim lRow, lgTab, chkSts
    '-----------------------
    'Reset variables area
    '-----------------------
    'set vat
    '*************************
    call setVatAmt
    '**************************

    Call ggoOper.LockField(Document, "Q")
	Call SetSpreadLockAfterQuery
	Call CurFormatNumericOCX()
	Call CurFormatNumSprSheet()			

	'Totalsum 계산 
'	For lRow = 1 To frm1.vspdData.MaxRows step 1
'		Call TotalSum(lRow)
'	Next
	lgIntFlgMode = Parent.OPMD_UMODE
	lgIntFlgMode_Dtl = Parent.OPMD_UMODE

	 chkSts = "DB"
	Call BtnToolCtrl(chkSts)
    lgBlnFlgChgValue = False

	ggoOper.SetReqAttr	frm1.txtPoNo2, "Q"
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
	End If
'	Set gActiveElement = document.activeElement
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

Sub BtnToolCtrl(byval chkSts)

	Dim lgTab
	lgTab = gSelframeFlg

	If lgTab = TAB1 or lgTab = TAB3 Then
		if frm1.hdnRelease.value  = "Y"  then
			if chkSts = "DB" then
				Call ChangeTag(true)
			End if
			Call SetToolbar("11100000001111")
			If frm1.hdclsflg.value = "Y" then
				frm1.btnCfm.disabled = true
				'frm1.btnSel.disabled = True
			else
			    frm1.btnCfm.disabled = False
			    'frm1.btnSel.disabled = True
			End If
			frm1.btnCfm.value = "확정취소"
			frm1.btnSend.disabled = False
			' === 2005.07.15 단가 일괄 불러오기 ==============
			frm1.btnCallPrice.disabled = True
			' === 2005.07.15 단가 일괄 불러오기 ==============
		Else
			if chkSts = "DB" then
				Call ChangeTag(False)
			end if
			frm1.txtPoNo.focus
		'	Call SetToolbar("11111000001111")
			if frm1.hdclsflg.value = "Y" then ' 미확정이나 close 됐으면 disable
				frm1.btnCfm.disabled = true		'close되지 않았으면 able
				'frm1.btnSel.disabled = True
			else
			    frm1.btnCfm.disabled = False
			    'frm1.btnSel.disabled = True
			end if
			frm1.btnCfm.value = "확정"
			frm1.btnSend.disabled = True
			' === 2005.07.15 단가 일괄 불러오기 ==============
			frm1.btnCallPrice.disabled = True
			' === 2005.07.15 단가 일괄 불러오기 ==============
		End if

		If lgIntFlgMode = Parent.OPMD_CMODE  Then
			Call SetToolbar("11101000000111")
		Else
			Call SetToolbar("11111000000111")
		End If

	Elseif lgTab = TAB2	then
		if frm1.hdnRelease.value  = "Y"  then
			if chkSts = "DB" then
				Call ChangeTag(true)
			end if
			Call SetToolbar("11100000001111")
			if frm1.hdclsflg.value = "Y" then
				frm1.btnCfm.disabled = true
				'frm1.btnSel.disabled = True
			else
			    frm1.btnCfm.disabled = False
			    'frm1.btnSel.disabled = True
			end if
			frm1.btnCfm.value = "확정취소"
			frm1.btnSend.disabled = False
			' === 2005.07.15 단가 일괄 불러오기 ==============
			frm1.btnCallPrice.disabled = True
			' === 2005.07.15 단가 일괄 불러오기 ==============
		Else
			if chkSts = "DB" then
				Call ChangeTag(False)
			end if
			frm1.txtPoNo.focus
		'	Call SetToolbar("11111000001111")
			if frm1.hdclsflg.value = "Y" then ' 미확정이나 close 됐으면 disable
				frm1.btnCfm.disabled = true		'close되지 않았으면 able
				'frm1.btnSel.disabled = True
			else
			    frm1.btnCfm.disabled = False
			    'frm1.btnSel.disabled = False
			end if
			frm1.btnCfm.value = "확정"
			frm1.btnSend.disabled = True
			' === 2005.07.15 단가 일괄 불러오기 ==============
			frm1.btnCallPrice.disabled = False
			' === 2005.07.15 단가 일괄 불러오기 ==============
			If lgIntFlgMode_Dtl = Parent.OPMD_CMODE  Then
				Call SetToolbar("11101101001111")
				'Call SetToolbar("11101111001111")
			Else
				if frm1.vspdData.MaxRows = 0 then
					Call SetToolbar("11101101001111")
				Else
					Call SetToolbar("11111111001111")
				End if
			End If

		End if
	End if

End Sub


'========================================================================================
' Function Name :
' Function Desc : 실제 저장 로직을 수행 , 성공적이면 DBSaveOk 호출됨 
'========================================================================================
 Function DbSave(byval btnflg)

    Dim lRow
    Dim lGrpCnt
    Dim retVal
    Dim boolCheck
    Dim lStartRow
    Dim lEndRow
    Dim lRestGrpCnt
	Dim strVal,strDel
	Dim ColSep, RowSep
	Dim intIndex
	Dim iOrderQty
	Dim iCost
	Dim iOrderAmt
'<!--20040513-->
	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규]
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규]
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제]
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size

	Dim ii


'<!--20040513-->
    Err.Clear
    DbSave = False

   	ColSep = Parent.gColSep
	RowSep = Parent.gRowSep

	With frm1
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtMode.value = Parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

		.hdnchgValue.value = lgBlnFlgChgValue
		.hdnSSCheckValue.value = ggospread.SSCheckChange

		if btnflg = "Cfm" then
			.txtMode.value = "Release"
		elseif btnflg = "UnCfm" then
			.txtMode.value = "UnRelease"
		end if

		if frm1.rdoMergPurFlg(0).Checked = True then
			frm1.hdnMergPurFlg.Value = "Y"
		else
			frm1.hdnMergPurFlg.Value = "N"
		end if

		'-----------------------
		'Data manipulate area - For PO Detail
		'-----------------------
		lGrpCnt = 0

		strVal = ""
		strDel = ""

		iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
		iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

		ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
		ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

		iTmpCUBufferCount = -1
		iTmpDBufferCount = -1

		strCUTotalvalLen = 0
		strDTotalvalLen  = 0

		'-----------------------
		'Data manipulate area
		'-----------------------

'		If frm1.vspdData.MaxRows = 0 then
'			Call DisplayMsgBox("173200", "X", "X", "X")
'			Call ClickTab2
'			Exit Function
'		End if

		If ggospread.SSCheckChange = true then
			'.hdnSSCheckValue.value = true
			For lRow = 1 To .vspdData.MaxRows step 1
		    frm1.vspdData.Row = lRow
	       		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
				Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
				if Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))=ggoSpread.InsertFlag then
					strVal = "C" & ColSep
				Else
					strVal = "U" & ColSep
				End if

				If Trim(UNICDbl(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))) = "" Or Trim(UNICDbl(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))) = "0" then
					Call DisplayMsgBox("970021", "X","발주수량", "X")
					Call LayerShowHide(0)
					Exit Function
				End if

				strVal = strVal & Trim(GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")) & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_PlantCd,lRow,"X","X")) & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_itemCd,lRow,"X","X")) & ColSep

                If Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X")),0)  & ColSep
				End If

                strVal = strVal & Trim(GetSpreadText(.vspdData,C_OrderUnit,lRow,"X","X"))  & ColSep

                If Trim(GetSpreadText(.vspdData,C_Cost,lRow,"X","X"))="" Then
					strVal = strVal & "0"  & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Cost,lRow,"X","X")),0)  & ColSep
				End If

                strVal = strVal & Trim(GetSpreadText(.vspdData,C_CostConCd,lRow,"X","X"))  & ColSep

                If Trim(GetSpreadText(.vspdData,C_OrderAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0"  & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderAmt,lRow,"X","X")),0)  & ColSep
				End If

                strVal = strVal & Trim(GetSpreadText(.vspdData,C_IOFlgCd,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_VatType,lRow,"X","X"))  & ColSep

                If Trim(GetSpreadText(.vspdData,C_VatRate,lRow,"X","X"))="" Then
					strVal = strVal & "0"  & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_VatRate,lRow,"X","X")),0)  & ColSep
				End If

                If Trim(GetSpreadText(.vspdData,C_VatAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0" & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_VatAmt,lRow,"X","X")),0)   & ColSep
				End If

                strVal = strVal & UNIConvDate(Trim(GetSpreadText(.vspdData,C_DlvyDT,lRow,"X","X")))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_HSCd,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_SLCd,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_TrackingNo,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Lot_No,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Lot_Seq,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_RetCd,lRow,"X","X"))  & ColSep

                If Trim(GetSpreadText(.vspdData,C_Over,lRow,"X","X"))="" Then
					strVal = strVal & "0"  & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Over,lRow,"X","X")),0)   & ColSep
				End If

                If Trim(GetSpreadText(.vspdData,C_Under,lRow,"X","X"))="" Then
					strVal = strVal & "0"  & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_Under,lRow,"X","X")),0)  & ColSep
				End If

                strVal = strVal & Trim(GetSpreadText(.vspdData,C_PrNo,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_MvmtNo,lRow,"X","X"))  & ColSep
                '비고 추가 
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Remrk,lRow,"X","X"))  & ColSep

                strVal = strVal & Trim(GetSpreadText(.vspdData,C_MaintSeq,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_SoNo,lRow,"X","X"))  & ColSep
                strVal = strVal & Trim(GetSpreadText(.vspdData,C_Stateflg,lRow,"X","X"))  & ColSep

	                '반품등록 추가 C_IVNO,C_IVSEQ  27,28
                strVal = strVal & ""  & ColSep 'IV No.
                strVal = strVal & ""  & ColSep 'IV Seq.

                If Trim(GetSpreadText(.vspdData,C_NetAmt,lRow,"X","X"))="" Then
					strVal = strVal & "0"  & ColSep
				Else
					strVal = strVal & UNIConvNum(Trim(GetSpreadText(.vspdData,C_NetAmt,lRow,"X","X")),0)   & ColSep
				End If

                iOrderQty=UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderQty,lRow,"X","X")),0)
                iCost=UNIConvNum(Trim(GetSpreadText(.vspdData,C_Cost,lRow,"X","X")),0)
                iOrderAmt=UNIConvNum(Trim(GetSpreadText(.vspdData,C_OrderAmt,lRow,"X","X")),0)

				If UNIConvNum(UNIConvNumPCToCompanyByCurrency(iOrderQty*iCost,frm1.txtCurr.value, Parent.ggAmtOfMoneyNo,"X","X"),0) = iOrderAmt Then
					strVal = strVal & "N"  & ColSep
				Else
					strVal = strVal & "Y"  & ColSep
				End If

                strVal = strVal & lRow & RowSep

            Case ggoSpread.DeleteFlag

		        strDel = "D" & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_SeqNo,lRow,"X","X")) & ColSep
				strDel = strDel & Trim(GetSpreadText(.vspdData,C_PrNo,lRow,"X","X")) & ColSep
				strDel = strDel & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep & ColSep
                strDel = strDel & lRow & RowSep

				lGrpCnt = lGrpCnt + 1
        End Select

		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

		         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		            objTEXTAREA.name = "txtCUSpread"
		            objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)

		            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
				 ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If

		         iTmpCUBufferCount = iTmpCUBufferCount + 1

		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		   Case ggoSpread.DeleteFlag
		         If strDTotalvalLen + Len(strDel) >  parent.C_FORM_LIMIT_BYTE Then   '한개의 form element에 넣을 한개치가 넘으면 
		            Set objTEXTAREA   = document.createElement("TEXTAREA")
		            objTEXTAREA.name  = "txtDSpread"
		            objTEXTAREA.value = Join(iTmpDBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)

		            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
		            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
		            iTmpDBufferCount = -1
		            strDTotalvalLen = 0
		         End If

		         iTmpDBufferCount = iTmpDBufferCount + 1

		         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '버퍼의 조정 증가치를 넘으면 
		            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
		            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
		         End If

		         iTmpDBuffer(iTmpDBufferCount) =  strDel
		         strDTotalvalLen = strDTotalvalLen + Len(strDel)
		End Select
	Next


	frm1.txtMaxRows.value = lGrpCnt
	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If

	If iTmpDBufferCount > -1 Then    ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)
	End If

 		End if
End with
	If LayerShowHide(1) = False Then Exit Function
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	DbSave = True

End Function
'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()
	lgBlnFlgChgValue = False
	Call MainQuery()
	'Call fncQuery()
End Function

'========================================================================================
' Function Name : chkEachFieldDomestic, chkEachFieldImport
' Function Desc : Manual check whether a value is entered at required field
'========================================================================================
Function chkEachFieldDomestic()
	chkEachFieldDomestic = True

	If Not chkFieldByCell (frm1.txtPotypeCd, "A",1) then
		chkEachFieldDomestic = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtSupplierCd, "A",1) then
		chkEachFieldDomestic = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtPoDt, "A",1) then
		chkEachFieldDomestic = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtGroupCd, "A",1) then
		chkEachFieldDomestic = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtCurr, "A",1) then
		chkEachFieldDomestic = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtPayTermCd, "A",1) then
		chkEachFieldDomestic = False
		Exit Function
	End If

End Function

Function chkEachFieldImport()
	chkEachFieldImport	= True

	If Not chkFieldByCell (frm1.txtDvryDt, "A",1) then
		chkEachFieldImport = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtOffDt, "A",1) then
		chkEachFieldImport = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtIncotermsCd, "A",1) then
		chkEachFieldImport = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtTransCd, "A",1) then
		chkEachFieldImport = False
		Exit Function
	End If

	If Not chkFieldByCell (frm1.txtApplicantCd, "A",1) then
		chkEachFieldImport = False
		Exit Function
	End If

End Function


'========================================================================================
' Function Name : initFormatField()
' Function Desc : Manual Formatting fields as amount or date
'========================================================================================
Function  initFormatField()

	'Header
	call FormatDateField(frm1.txtPoDt)
	call FormatDateField(frm1.txtDvryDt)
	call FormatDateField(frm1.txtOffDt)
	call FormatDateField(frm1.txtExpiryDt)

	call FormatDoubleSingleField(frm1.txtXch)
	call FormatDoubleSingleField(frm1.txtPoAmt)
	call FormatDoubleSingleField(frm1.txtPoLocAmt)
	call FormatDoubleSingleField(frm1.txtGrossPoAmt)
	call FormatDoubleSingleField(frm1.txtGrossPoLocAmt)
	call FormatDoubleSingleField(frm1.txtVatAmt)
	call FormatDoubleSingleField(frm1.txtVatrt)
	call FormatDoubleSingleField(frm1.txtPayDur)


	call LockobjectField(frm1.txtPoDt,"R")
	call LockobjectField(frm1.txtDvryDt,"O")
	call LockobjectField(frm1.txtOffDt,"R")
	call LockobjectField(frm1.txtExpiryDt,"O")

	call LockobjectField(frm1.txtXch,"O")
	call LockobjectField(frm1.txtPoAmt,"P")
	call LockobjectField(frm1.txtPoLocAmt,"P")
	call LockobjectField(frm1.txtGrossPoAmt,"P")
	call LockobjectField(frm1.txtGrossPoLocAmt,"P")
	call LockobjectField(frm1.txtVatAmt,"P")
	call LockobjectField(frm1.txtVatrt,"P")
	call LockobjectField(frm1.txtPayDur,"O")


	call ggoOper.SetReqAttr(frm1.txtDvryDt, "D")
	call ggoOper.SetReqAttr(frm1.txtOffDt, "Q")
	call ggoOper.SetReqAttr(frm1.txtApplicantCd, "Q")
	call ggoOper.SetReqAttr(frm1.txtIncotermsCd, "Q")
	call ggoOper.SetReqAttr(frm1.txtTransCd, "Q")

    'Deail
	call FormatDoubleSingleField(frm1.txtDetailNetAmt)
    call FormatDoubleSingleField(frm1.txtDetailVatAmt)
	call FormatDoubleSingleField(frm1.txtDetailGrossAmt)

	call LockobjectField(frm1.txtDetailNetAmt,"P")
	call LockobjectField(frm1.txtDetailVatAmt,"P")
	call LockobjectField(frm1.txtDetailGrossAmt,"P")

End Function


' === 2005.07.15 단가 일괄 불러오기 관련 수정 ===========================================
Sub btnCallPrice_OnClick()
	Dim index

	If frm1.vspdData.Maxrows <= 0 then
		Exit Sub
	End if

'	If Trim(frm1.txtSupplierCd.value) = "" then
'		Call DisplayMsgBox("SCM003","X","X","X")
'		Call LayerShowHide(0)
'		frm1.txtSupplierCd.focus
'		Exit Sub
'	End If
'
'	If Trim(frm1.txtCurr.value) = "" then
'		Call DisplayMsgBox("SCM003","X","X","X")
'		Call LayerShowHide(0)
'		frm1.txtCurr.focus
'		Exit Sub
'	End If
'
'	Call SetPriceType2
	Call lookupPriceForSelection()

	For index = 1 to  frm1.vspdData.Maxrows
'	    frm1.vspdData.row = index
'	    frm1.vspdData.Col = C_SelCheck
'
'	    If frm1.vspdData.Text = "1" then
'			frm1.vspdData.Col = 0
			ggoSpread.UpdateRow index
'	    Else
'			'frm1.vspdData.Col = 0
'			'ggoSpread.EditUndo
'	    End If
	Next

End Sub

Sub btnCallPrice_Ok()
Dim lRow
	With frm1
	For lRow = 1 To .vspdData.MaxRows
'		.vspdData.Row = lRow
'		.vspdData.Col = C_Check

'		If .vspdData.Text <> "0" Then
			Call vspdData_Change(C_Cost, lRow)
'		End If
	Next
	End With
End Sub

Sub SetPriceType()
	Dim IntRetCd, lsPriceType

	IntRetCD = CommonQueryRs("MINOR_CD", "B_CONFIGURATION", "(MAJOR_CD = 'M0001' AND REFERENCE = 'Y' )", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	lsPriceType = TRIM(REPLACE(lgF0,CHR(11),""))

	frm1.hdnPriceType.value = lsPriceType			'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)

End Sub

Sub SetPriceType2()
	If frm1.rdoPrcTypeflg1.checked = true then
		lsPriceType = "T"
		frm1.hdnPriceType.value = "T"				'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)
	Else
		lsPriceType = "N"
		frm1.hdnPriceType.value = "N"				'2005-05-27 수정(M0001에 설정되어 있는 규칙에 의한 단가 불러오기)
	End if

End Sub


Function lookupPriceForSelection()
    Err.Clear
    Dim strVal
    Dim lColSep,lRowSep
    Dim lRow
    Dim lGrpCnt

    If CheckRunningBizProcess = True Then
		Exit Function
	End If

	If Not chkField(Document, "2") Then
		Exit Function
	End If

'	If Not chkField(Document, "2") OR Not ggoSpread.SSDefaultCheck Then
'		Exit Function
'	End If

	lgBlnFlgChgValue = true

    If LayerShowHide(1) = False Then Exit Function

	With frm1

    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    strVal = ""

    '-----------------------
    'Data manipulate area
    '-----------------------
	.txtMode.value = "lookupPriceForSelection"

	For lRow = 1 To .vspdData.MaxRows

		.vspdData.Row = lRow
		.vspdData.Col = C_Check

		If .vspdData.Text <> "0" Then

			frm1.vspdData.Row = lRow

			strVal = strVal & Trim(frm1.txtSupplierCd.Value) & parent.gColSep
			frm1.vspdData.Col = C_ItemCd
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			frm1.vspdData.Col = C_PlantCd
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			frm1.vspdData.Col = C_OrderUnit
			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			strVal = strVal & Trim(frm1.txtCurr.Value) & parent.gColSep & parent.gColSep
'			frm1.vspdData.Col = C_PoPrice1
'			strVal = strVal & Trim(frm1.vspdData.text) & Parent.gColSep
			strVal = strVal & lRow & Parent.gRowSep

			lGrpCnt = lGrpCnt + 1

			frm1.vspdData.Col = C_Cost
			frm1.vspdData.Text = 0
		End If
	Next

	If strVal <> "" Then
		If LayerShowHide(1) = False Then Exit Function

		.hdnMaxRows.value = .vspdData.MaxRows
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End If
	End With
End Function

' === 2005.07.15 단가 일괄 불러오기 관련 수정 ===========================================



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>발주일반정보(KO441)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" onMouseOver="vbscript:SetClickflag" onMouseOut="vbscript:ResetClickflag" onFocus="vbscript:SetClickflag" onBlur="vbscript:ResetClickflag">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>발주내역정보</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab3()" onMouseOver="vbscript:SetClickflag" onMouseOut="vbscript:ResetClickflag" onFocus="vbscript:SetClickflag" onBlur="vbscript:ResetClickflag">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>발주무역정보</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenReqRef">구매요청참조</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT Class = required  TYPE=TEXT NAME="txtPoNo" SIZE=32  MAXLENGTH=18 ALT="발주번호" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS=TD6></TD>
									<TD CLASS=TD6></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR height="*">
					<TD WIDTH=100% valign=top>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="발주번호" NAME="txtPoNo2"  MAXLENGTH=18 SIZE=34 tag="21XXXU" ></TD>
									<TD CLASS="TD5" NOWRAP>확정여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="발주확정" NAME="rdoRelease" CLASS="RADIO" checked tag="24" ONCLICK="vbscript:SetChangeflg()"><label for="rdoRelease">&nbsp;미확정&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="발주확정" NAME="rdoRelease" CLASS="RADIO" ONCLICK="vbscript:setChangeflg()" tag="24"><label for="rdoRelease">&nbsp;확정&nbsp;</label></TD>
								</TR>
									<TD CLASS="TD5" NOWRAP>발주형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="발주형태" NAME="txtPotypeCd"  MAXLENGTH=5 SIZE=10 tag="23NXXU" ONChange="vbscript:ChangePotype()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPotype()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT AlT="발주형태" NAME="txtPotypeNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="공급처" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT AlT="공급처" ID="txtSupplierNm" NAME="arrCond" tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주일</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
										   <TR>
									          <TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASS = required ALT=발주일 NAME="txtPoDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="22X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											  </TD>
											  <TD NOWRAP>
												&nbsp;예상납기일
											  </TD>
											  <TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=예상납기일 NAME="txtDvryDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 style="HEIGHT: 20px; WIDTH: 100px" tag="21X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											  </TD>
											</TR>
										</Table>
									</TD>
									<TD CLASS="TD5" NOWRAP>구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="구매그룹" NAME="txtGroupCd" MAXLENGTH=4 SIZE=10 tag="22NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
														   <INPUT TYPE=TEXT AlT="구매그룹" NAME="txtGroupNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>화폐</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="화폐" NAME="txtCurr" MAXLENGTH=3 SIZE=10 tag="22NXXU" onChange="ChangeCurr()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCurr()">
														   <INPUT TYPE=HIDDEN AlT="화폐" NAME="txtCurrNm" tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>환율</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
												<TR>
													<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=환율 NAME="txtXch" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 style="HEIGHT: 20px; WIDTH: 120px" tag="21X5Z" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
													<TD NOWRAP>
														&nbsp;<SELECT NAME="cboXchop" tag="22" STYLE="WIDTH:82px:" Alt="환율"></SELECT>
													</TD>
												</TR>
										</Table>
									</TD>

								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주순금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주금액     NAME="txtPoAmt"    CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>발주순자국금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주자국금액 NAME="txtPoLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></td>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>발주총금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주총금액     NAME="txtGrossPoAmt"    CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>발주총자국금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주총자국금액 NAME="txtGrossPoLocAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>VAT</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVattype" ALT="VAT"  MAXLENGTH=5 SIZE=10 tag="21NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVat(1)">
														   <INPUT TYPE=TEXT ALT="VAT" NAME="txtVatTypeNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>VAT금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT금액 NAME="txtVatAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 style="HEIGHT: 20px; WIDTH: 250px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>VAT율</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=VAT율 NAME="txtVatrt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 style="HEIGHT: 20px; WIDTH: 160px" tag="24X5" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT>&nbsp;&nbsp;%</TD>
																<TD CLASS="TD5" nowrap>VAT포함구분</TD>
								    <TD CLASS="TD6" nowrap>
								    <INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT포함구분" CLASS="RADIO" checked id="rdoVatFlg1" tag="21X"><label for="rdoVatFlg">별도 </label>
									<INPUT TYPE=radio NAME="rdoVatFlg" ALT="VAT포함구분" CLASS="RADIO" id="rdoVatFlg2"  tag="21X"><label for="rdoVatFlg">포함&nbsp;</label></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>결제방법</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT AlT="결제방법" NAME="txtPayTermCd"  MAXLENGTH=5 SIZE=10 tag="22NXXU" OnChange="VBScript:changePayterm()" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPayMeth()">
														   <INPUT TYPE=TEXT AlT="결제방법" NAME="txtPayTermNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 >
														   <INPUT TYPE=HIDDEN AlT="결제방법" NAME="txtReference" tag="24X"></TD>
									<TD CLASS="TD5" NOWRAP>결제기간</TD>
									<TD CLASS="TD6" NOWRAP>
										<Table cellpadding=0 cellspacing=0>
											<TR>
												<TD NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=결제기간 NAME="txtPayDur" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 style="HEIGHT: 20px; WIDTH: 80px" tag="21X70" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT>
												</TD>
												<TD NOWRAP>
													&nbsp;일
												</TD>
											</TR>
										</Table>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>지급유형</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="지급유형" NAME="txtPayTypeCd"  MAXLENGTH=5 SIZE=10 tag="21NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPayType()">
														   <INPUT TYPE=TEXT AlT="지급유형" NAME="txtPayTypeNm" SIZE=20 tag="24X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>통합구매여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="통합구매여부" NAME="rdoMergPurFlg" CLASS="RADIO" tag="21" id="rdoMergPurFlg1" ONCLICK="vbscript:SetChangeflg()"><label for="rdoMergPurFlg1">&nbsp;YES&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="통합구매여부" NAME="rdoMergPurFlg" CLASS="RADIO" checked id="rdoMergPurFlg2" ONCLICK="vbscript:setChangeflg()" tag="21"><label for="rdoMergPurFlg2">&nbsp;NO&nbsp;</label></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>OFFER작성일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASS = required ALT=OFFER작성일 NAME="txtOffDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="31X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>수입자</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required  TYPE=TEXT NAME="txtApplicantCd" MAXLENGTH=10 SIZE=10 ALT ="수입자" tag="31NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBiz('Appl')">
														   <INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 ALT ="수입자" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>가격조건</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT NAME="txtIncotermsCd" ALT ="가격조건"  MAXLENGTH=5 SIZE=10 tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9006')">
														   <INPUT TYPE=TEXT NAME="txtIncotermsNm" ALT ="가격조건" SIZE=20 tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>운송방법</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS = required TYPE=TEXT NAME="txtTransCd"  MAXLENGTH=5 SIZE=10 ALT ="운송방법" tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9009')">
														   <INPUT TYPE=TEXT NAME="txtTransNm" SIZE=20 ALT ="운송방법" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처영업담당</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="공금처영업담당" NAME="txtSuppSalePrsn" MAXLENGTH=50 SIZE=34 tag="21"></TD>
									<TD CLASS="TD5" NOWRAP>긴급연락처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="긴급연락처" NAME="txtTel" MAXLENGTH=30 SIZE=34 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">대금결제참조</TD>
									<TD CLASS="TD6" colspan=3 width=100% NOWRAP><INPUT TYPE=TEXT AlT="대금결제참조" Size="90" NAME="txtPayTermstxt" MAXLENGTH=120 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">비고</TD>
									<TD CLASS=TD6 Colspan=3 WIDTH=100% NOWRAP><INPUT TYPE=TEXT  NAME="txtRemark" ALT="비고" tag = "21" SIZE=90 MAXLENGTH=70></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(2)%>
							</TABLE>
						</div>
						<!--두번째 탭 -->
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주순금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="발주금액" NAME="txtDetailNetAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 234px" tag="24X2" Title="FPDOUBLESINGLE" CLASS = protected readonly = True TabIndex = -1 ></OBJECT>');</SCRIPT></td>
									<TD CLASS="TD5" NOWRAP>VAT금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="VAT금액" NAME="txtDetailVatAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 234px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>발주총금액</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="발주총금액" NAME="txtDetailGrossAmt" CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 style="HEIGHT: 20px; WIDTH: 234px" tag="24X2" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></td>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>

								</TR>
								<TR>
									<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</DIV>
						<!--세번째 탭 -->
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>유효일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=유효일 NAME="txtExpiryDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 style="HEIGHT: 20px; WIDTH: 100px" tag="31X1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>INVOICE NO.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInvNo" MAXLENGTH=50 SIZE=34 ALT="INVOICE NO." MAXLENGTH=20 tag="31"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>송금은행</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBankCd"  MAXLENGTH=10 SIZE=10 ALT ="송금은행" tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBank()">
														   <INPUT TYPE=TEXT NAME="txtBankNm" SIZE=20 ALT ="송금은행" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>인도장소</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDvryPlce" MAXLENGTH=5 SIZE=10 ALT="인도장소" tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9095')">
														   <INPUT TYPE=TEXT NAME="txtDvryPlceNm" SIZE=20 ALT ="인도장소" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>대행자</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAgentCd"  MAXLENGTH=10 SIZE=10 ALT ="대행자" tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBiz('Agent')">
														   <INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 ALT ="대행자" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>제조자</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtManuCd" MAXLENGTH=10 SIZE=10 ALT ="제조자" tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBiz('Manu')">
														   <INPUT TYPE=TEXT NAME="txtManuNm" SIZE=20 ALT ="제조자" tag="34X" CLASS = protected readonly = True TabIndex = -1  ></TD>
								</TR>
								<TR>

									<TD CLASS="TD5" NOWRAP>원산지</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOrigin"  MAXLENGTH=5 SIZE=10 ALT="원산지" tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9094')">
														   <INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 ALT ="원산지" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>포장조건</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPackingCd" MAXLENGTH=5 SIZE=10 ALT ="포장조건" tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9007')">
														   <INPUT TYPE=TEXT NAME="txtPackingNm" SIZE=20 ALT ="포장조건" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사방법</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtInspectCd" MAXLENGTH=5 SIZE=10 ALT ="검사방법" tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9008')">
														   <INPUT TYPE=TEXT NAME="txtInspectNm" SIZE=20 ALT ="검사방법" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>도착도시</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDisCity" MAXLENGTH=5 ALT="도착도시" SIZE=10 tag="31NXXU"  ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9096')">
														   <INPUT TYPE=TEXT NAME="txtDisCityNm" SIZE=20 ALT ="도착도시" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>선적항</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLoadPort" MAXLENGTH=5 ALT="선적항" SIZE=10 tag="31NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9092-1')">
														   <INPUT TYPE=TEXT NAME="txtLoadPortNm" SIZE=20 ALT ="선적항" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
									<TD CLASS="TD5" NOWRAP>도착항</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDisPort" MAXLENGTH=5 ALT="도착항" SIZE=10 tag="31XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenMinorCode('B9092')">
														   <INPUT TYPE=TEXT NAME="txtDisPortNm" SIZE=20 ALT ="도착항" tag="34X" CLASS = protected readonly = True TabIndex = -1 ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>선적기한</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtShipment" MAXLENGTH=70 ALT="선적기한" SIZE=34 tag="31"></TD>
									<TD CLASS=TD5></TD>
									<TD CLASS=TD6></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(2)%>
							</TABLE>
						</DIV>
					</TD>
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td ><button name="btnCfmSel" id="btnCfm" class="clsmbtn" ONCLICK="vbscript:Cfm()">확정</button>&nbsp;
						<BUTTON NAME="btnCallPrice" CLASS="CLSMBTN">단가불러오기</BUTTON>&nbsp;

					 <Div  STYLE="DISPLAY: none"><button name="btnSend" id="btnSend" class="clsmbtn" ONCLICK="Sending()">주문서발송</button></Div>
					</td>
                                        <TD WIDTH=10><a><button name="btnAutoTest" class="clsmbtn">테스트</button></a>&nbsp;</TD>
					<td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(2)">경비등록</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<!--	추가부분 시작	-->
<P ID="divTextArea"></P>
<!--	추가부분 끝	    -->
<TEXTAREA CLASS="hidden"  NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRelease" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCurr" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

<INPUT TYPE=HIDDEN NAME="hdnreference"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBLflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCCflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdvatFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIssueType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMergPurFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaintNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdclsflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdntotPoAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnVATINCFLG" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnxchrateop" tag="2">
<!-- 20031117-->
<INPUT TYPE=HIDDEN NAME="hdnMaxRows" tag="14">
<INPUT TYPE=HIDDEN NAME="hdntxtVatType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdntxtVatTypeNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdntxtVatrt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnchgValue"  tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSSCheckValue" tag="24">

<!-- 2005.07.15 -->
<INPUT TYPE=HIDDEN NAEM="hdnPriceType" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>