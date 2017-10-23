<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111MA3
'*  4. Program Name         : 일괄출고등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/14
'*  8. Modified date(Last)  : 2002/06/17
'*  9. Modifier (First)     : 조인국 
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                               

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s4111mb3.asp"											'☆: Head Query 비지니스 로직 ASP명 

Const TAB1 = 1                 '☜: Tab의 위치 
Const TAB2 = 2

' Popup 관련 상수 
Const C_PopPlant		= 1			' 공장 
Const C_PopDnType		= 2			' 출하형태 
Const C_PopShipToParty	= 3			' 납품처 
Const C_PopSalesGrp		= 4			' 영업그룹 

Const C_PopTransMeth	= 1			' 운송방법 
Const C_PopInvMgr		= 2			' 재고담당자 

'☆: Spread Sheet의 Column별 상수 
Dim C_SELECT
Dim C_PROMISE_DT
Dim C_SHIP_TO_PARTY
Dim C_BP_NM
Dim C_ITEM_CD
Dim C_ITEM_NM
Dim C_REMAIN_QTY		'SC.CFM_QTY - SC.REQ_QTY AS '잔량' 
Dim C_BONUS_REMAIN_QTY	'SC.CFM_BONUS_QTY - SC.REQ_BONUS_QTY AS '덤잔량'
Dim C_SO_UNIT
Dim C_GI_QTY 
Dim C_GI_BONUS_QTY
Dim C_PLANT_CD
Dim C_PLANT_NM
Dim C_SL_CD
Dim C_SL_CD_POP
Dim C_SL_NM 
Dim C_SU_ONHAND_QTY		'수주단위 재고량	
Dim C_ONHAND_QTY		'OS.GOOD_ON_HAND_QTY - OS.PICKING_QTY AS '현재고량'
Dim C_BASIC_UNIT 
Dim C_SO_NO 
Dim C_SO_SEQ 
Dim C_SO_SCHD_NO	
Dim C_TRACKING_NO 
Dim C_SPEC	
Dim C_DN_TYPE
Dim	C_DN_TYPE_NM
Dim C_REMARK
Dim C_SO_TYPE
Dim C_SALES_GRP
' 전체선택 취소시 이전 데이터 복원위해 사용 
Dim C_OLD_SL_CD				' 변경전 창고 
Dim C_OLD_SL_NM				' 변경전 창고명 
Dim C_OLD_GI_QTY			' 변경전 출고가능 수량 
Dim C_OLD_GI_BONUS_QTY		' 변경전 출고가능 덤수량 

'=========================================
Dim lgBlnFlgChgValue			' Variable is for Dirty flag
Dim lgBlnFlgChgValue3			' Tag가 '3'으로 시작하는 필드이 변경여부(납품 및 운송정보)
Dim lgStrAllocInvFlag			' 재고할당 사용여부 

Dim lgIntFlgMode				' Variable is for Operation Status

Dim lgSortKey

Dim lgStrPrevKey
Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i   '@@@CommonQueryRs 를 위한 변수 

Dim lgBlnOpenPop						
Dim gSelframeFlg

'=========================================
Sub initSpreadPosVariables()

	C_SELECT			=	1
	C_PROMISE_DT		=	2
	C_SHIP_TO_PARTY		=	3
	C_BP_NM				=	4			
	C_ITEM_CD			=	5	
	C_ITEM_NM			=	6
	C_REMAIN_QTY		=	7	'SC.CFM_QTY - SC.REQ_QTY AS '잔량', 
	C_BONUS_REMAIN_QTY	=	8	'SC.CFM_BONUS_QTY - SC.REQ_BONUS_QTY AS '덤잔량',
	C_SO_UNIT			=	9
	C_GI_QTY			=	10 
	C_GI_BONUS_QTY		=	11
	C_PLANT_CD			=	12
	C_PLANT_NM			=	13
	C_SL_CD				=	14
	C_SL_CD_POP			=	15
	C_SL_NM				=	16 
	C_SU_ONHAND_QTY		=	17	
	C_ONHAND_QTY		=	18	'OS.GOOD_ON_HAND_QTY - OS.PICKING_QTY AS '현재고량',
	C_BASIC_UNIT		=	19 
	C_SO_NO				=	20 
	C_SO_SEQ			=	21 
	C_SO_SCHD_NO		=	22	
	C_TRACKING_NO		=	23 
	C_SPEC				=	24
	C_DN_TYPE			=	25
	C_DN_TYPE_NM		=	26
	C_REMARK			=	27
	C_SO_TYPE			=	28
	C_SALES_GRP			=	29
	C_OLD_SL_CD			=	30
	C_OLD_SL_NM			=	31
	C_OLD_GI_QTY		=	32 
	C_OLD_GI_BONUS_QTY	=	33
End Sub

'=========================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE            
    lgBlnFlgChgValue = False
    lgBlnFlgChgValue3 = False
    lgStrPrevKey = ""   
    	
End Sub

'=========================================
Sub SetDefaultVal()
	frm1.txtConPlant.focus
	frm1.txtConReqDateFrom.Text = EndDate
	frm1.txtConReqDateTo.Text = EndDate
	frm1.txtActualGIDt.Text = EndDate
	' 공장 
	frm1.txtConPlant.value = parent.gPlant
	frm1.txtConPlantNm.value = parent.gPlantNm

	Set gActiveElement = document.ActiveElement
	
	gSelframeFlg = TAB1
End Sub

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'=========================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    

	With ggoSpread
		.Source = frm1.vspdData
		.Spreadinit "V20030618",,parent.gAllowDragDropSpread    
	    
		frm1.vspdData.ReDraw = false
		frm1.vspdData.MaxCols = C_OLD_GI_BONUS_QTY + 1											'☜: 최대 Columns의 항상 1개 증가시킴	    
		frm1.vspdData.MaxRows = 0
	
		Call GetSpreadColumnPos("A")

					   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
		.SSSetCheck		C_SELECT,			"선택",			10,,,true
		.SSSetDate		C_PROMISE_DT,		"출고예정일",	12,2,Parent.gDateFormat    
	    .SSSetEdit		C_SHIP_TO_PARTY,	"납품처",		12,,,,2
	    .SSSetEdit		C_BP_NM,			"납품처명",		20,,,,2
	    .SSSetEdit		C_ITEM_CD,			"품목",			12,,,,2
	    .SSSetEdit		C_ITEM_NM,			"품목명",		20,,,,2
	    .SSSetFloat		C_REMAIN_QTY,		"잔량",			10,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    .SSSetFloat		C_BONUS_REMAIN_QTY,"덤잔량",		10,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    .SSSetEdit		C_SO_UNIT,			"단위",			6,		2
		.SSSetFloat		C_GI_QTY,			"출고량",		15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		.SSSetFloat		C_GI_BONUS_QTY,		"덤출고량",		15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		.SSSetEdit		C_PLANT_CD,			"공장",			8,,,7,2
		.SSSetEdit		C_PLANT_NM,			"공장명",		15,,,,2	    		
		.SSSetEdit		C_SL_CD,			"창고",			8,,,7,2
		.SSSetButton	C_SL_CD_POP
		.SSSetEdit		C_SL_NM,			"창고명",		15,,,,2	 		
		.SSSetFloat		C_SU_ONHAND_QTY,	"수주단위재고량",15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    .SSSetFloat		C_ONHAND_QTY,		"현재고량",		15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
	    .SSSetEdit		C_BASIC_UNIT,		"재고단위",		10,		2
	    .SSSetEdit		C_SO_NO,			"수주번호",		15,,,,2
		.SSSetEdit		C_SO_SEQ,			"수주순번",		10,		1,					,	  ,		  2
		.SSSetEdit		C_SO_SCHD_NO,		"납품순서",		10,		1,					,	  ,		  2		
		.SSSetEdit		C_TRACKING_NO,		"Tracking No",	15,,,,2
		.SSSetEdit		C_SPEC,				"품목규격",		30
		.SSSetEdit		C_DN_TYPE,			"출하형태",		10,		2,					,		,		2
		.SSSetEdit		C_DN_TYPE_NM,		"출하형태명",	20
		.SSSetEdit		C_REMARK,			"비고",			30,		,					,	  120
		.SSSetEdit		C_SO_TYPE,			"수주형태",		0
		.SSSetEdit		C_SALES_GRP,		"영업그룹",		0
		.SSSetEdit		C_OLD_SL_CD,		"창고",			0
		.SSSetEdit		C_OLD_SL_NM,		"창고명",		0
		.SSSetFloat		C_OLD_GI_QTY,		"출고량",		0,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		.SSSetFloat		C_OLD_GI_BONUS_QTY,	"덤출고량",		0,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
		
		Call .MakePairsColumn(C_SL_CD,C_SL_CD_POP)
	    Call .SSSetColHidden(C_SO_TYPE, frm1.vspdData.MaxCols, True)
	    
	    Call SetSpreadLock
    End With
    
	frm1.vspdData.ReDraw = True

End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock 1 , -1
	ggoSpread.SpreadUnLock	C_SELECT, -1, C_SELECT
End Sub

'=========================================
Sub SetSpreadColor(ByVal iIntRow, ByVal iIntRow2, ByVal pvIntButtonDown) 
	frm1.vspdData.Redraw = False
	With ggoSpread
		' 선택 
		If pvIntButtonDown = 1 Then
			.SpreadUnLock C_GI_QTY,			iIntRow, C_GI_QTY,		 iIntRow2
			.SpreadUnLock C_GI_BONUS_QTY,	iIntRow, C_GI_BONUS_QTY, iIntRow2
			.SpreadUnLock C_SL_CD,			iIntRow, C_SL_CD_POP,	 iIntRow2
			.SSSetRequired C_GI_QTY,	    iIntRow, iIntRow2  
			.SSSetRequired C_GI_BONUS_QTY,  iIntRow, iIntRow2  			
			.SSSetRequired C_SL_CD,			iIntRow, iIntRow2
		Else
			.SpreadLock C_GI_QTY,		iIntRow, C_GI_QTY,		 iIntRow2
			.SpreadLock C_GI_BONUS_QTY, iIntRow, C_GI_BONUS_QTY, iIntRow2
			.SpreadLock C_SL_CD,		iIntRow, C_SL_CD_POP,	 iIntRow2
		End If
	End With
	frm1.vspdData.Redraw = True
End Sub

' 에러 발생시 해당 위치로 Focus이동 
'=========================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow

           If Not Frm1.vspdData.ColHidden Then
			  Call SetActiveCell(frm1.vspdData, iDx, iRow,"M","X","X")
              Exit For
           End If
       Next
    End If   
End Sub

' 조회조건 Popup
'=========================================
Function OpenConPopUp(Byval pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True

	With frm1
		Select Case pvIntWhere
			Case C_PopPlant		'공장 
				iArrParam(1) = "dbo.B_PLANT"									
				iArrParam(2) = Trim(.txtConPlant.value)				
				iArrParam(3) = ""										
				iArrParam(4) = ""										
				
				iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"
							
				iArrHeader(0) = .txtConPlant.alt						
				iArrHeader(1) = .txtConPlantNm.alt					
	
				.txtConPlant.focus

			Case C_PopDnType	'출하형태 
				iArrParam(1) = "dbo.B_MINOR MN "		
				iArrParam(2) = Trim(.txtConDnType.value)					
				iArrParam(3) = ""											
				iArrParam(4) = "MN.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND EXISTS (SELECT * FROM dbo.S_SO_TYPE_CONFIG ST WHERE	ST.MOV_TYPE = MN.MINOR_CD) "			
				
				iArrField(0) = "ED15" & Parent.gColSep & "MN.MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MN.MINOR_NM"
				
				iArrHeader(0) = .txtConDnType.alt							
				iArrHeader(1) = .txtConDnTypeNm.alt	
				
				frm1.txtConDnType.focus

			Case C_PopShipToParty	'납품처 
				iArrParam(1) = "dbo.B_BIZ_PARTNER BP INNER JOIN dbo.B_COUNTRY CT ON (CT.COUNTRY_CD = BP.CONTRY_CD)"								
				iArrParam(2) = Trim(.txtConShipToParty.value)			
				iArrParam(3) = ""											
				iArrParam(4) = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BPF WHERE BPF.PARTNER_BP_CD = BP.BP_CD AND BPF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ")"						
	
				iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
				iArrField(2) = "ED10" & Parent.gColSep & "BP.CONTRY_CD"
				iArrField(3) = "ED20" & Parent.gColSep & "CT.COUNTRY_NM"
    
				iArrHeader(0) = .txtConShipToParty.alt					
				iArrHeader(1) = .txtConShipToPartyNm.alt					
				iArrHeader(2) = "국가"
				iArrHeader(3) = "국가명"

				.txtConShipToParty.focus
			
			' 영업그룹 
			Case C_PopSalesGrp												
				iArrParam(1) = "dbo.B_SALES_GRP"
				iArrParam(2) = Trim(.txtConSalesGrp.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
					
				iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
				iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
			    iArrHeader(0) = .txtConSalesGrp.Alt
			    iArrHeader(1) = .txtConSalesGrpNm.Alt
				    
			    .txtConSalesGrp.focus
		End Select
	End With
	
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConPopUp = SetConPopUp(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================
Function OpenConSoNo(ByRef prObjSoNo)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True

	iArrParam(1) = "S_SO_HDR SH, B_BIZ_PARTNER SP, B_SALES_GRP SG"
	iArrParam(2) = Trim(prObjSoNo.value)
	iArrParam(3) = ""
				
	' 재고할당을 사용여부 
	If lgStrAllocInvFlag = "N" Then
		iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & "  AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_DTL SD WHERE SD.SO_NO = SH.SO_NO AND SD.SO_QTY + SD.BONUS_QTY > SD.REQ_QTY + SD.REQ_BONUS_QTY) "
	Else
		iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & "  AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_SCHD SC WHERE SC.SO_NO = SH.SO_NO AND SC.ALLC_QTY + SC.ALLC_BONUS_QTY > SC.REQ_QTY + SC.REQ_BONUS_QTY) "
	End If
	iArrParam(5) = "수주번호"

	iArrField(0) = "ED12" & Parent.gColSep & "SH.SO_NO"
	iArrField(1) = "ED10" & Parent.gColSep & "SH.SOLD_TO_PARTY"
	iArrField(2) = "ED20" & Parent.gColSep & "SP.BP_NM"
	iArrField(3) = "DD10" & Parent.gColSep & "SH.SO_DT"
	iArrField(4) = "ED15" & Parent.gColSep & "SG.SALES_GRP_NM"
	iArrField(5) = "ED10" & Parent.gColSep & "SH.PAY_METH"
				
	iArrHeader(0) = "수주번호"
	iArrHeader(1) = "주문처"
	iArrHeader(2) = "주문처명"
	iArrHeader(3) = "수주일"
	iArrHeader(4) = "영업그룹명"
	iArrHeader(5) = "결제방법"
	
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=500px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		prObjSoNo.value = iArrRet(0)
	End If
	
	prObjSoNo.Focus
End Function

' 입력 관련 Popup
'=========================================
Function OpenPopUp(Byval pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True

	With frm1
		Select Case pvIntWhere
			Case C_PopTransMeth	'운송방법 
				iArrParam(1) = "dbo.B_MINOR"
				iArrParam(2) = Trim(.txtTransMeth.value)
				iArrParam(3) = ""											
				iArrParam(4) = "MAJOR_CD = " & FilterVar("B9009", "''", "S") & ""
				
				iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
							
				iArrHeader(0) = .txtTransMeth.alt						
				iArrHeader(1) = .txtTransMethNm.alt						

				.txtTransMeth.focus

			'재고담당 
			Case C_PopInvMgr
				iArrParam(1) = "dbo.B_MINOR"
				iArrParam(2) = Trim(.txtInvMgr.value)
				iArrParam(3) = ""											
				iArrParam(4) = "MAJOR_CD = " & FilterVar("I0004", "''", "S") & ""
				
				iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
							
				iArrHeader(0) = .txtInvMgr.alt						
				iArrHeader(1) = .txtInvMgrNm.alt						

				.txtInvMgr.focus
				
		End Select
	End With
	
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		OpenPopUp = SetPopUp(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================
Function OpenZip()
	Dim iArrRet
	Dim iArrParam(2)

	If lgBlnOpenPop = True Then Exit Function
	
	lgBlnOpenPop = True
	
	iArrParam(0) = Trim(frm1.txtZIPcd.value)
	iArrParam(1) = ""
	iArrParam(2) = Trim(frm1.txtHCntryCd.value)
	
	iArrRet = window.showModalDialog("../../comasp/ZipPopup.asp", Array(window.parent, iArrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	lgBlnOpenPop = False

	frm1.txtZIPcd.focus
		
	If iArrRet(0) <> "" Then
		frm1.txtZIPcd.value = iArrRet(0)
		frm1.txtADDR1.value = iArrRet(1)
		frm1.txtSTPInfoNo.value = ""
		lgBlnFlgChgValue3 = True		
	End If	
			
End Function

'========================================
Function OpenTransCo()
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgBlnOpenPop = True Then Exit Function

	lgBlnOpenPop = True

	iArrParam(0) = "운송회사"							
	iArrParam(1) = "B_MAJOR A , B_MINOR B"						
	iArrParam(2) = ""										
	iArrParam(3) = ""									
	iArrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9031", "''", "S") & " "				
	iArrParam(5) = "운송회사"							

	iArrField(0) = "B.MINOR_CD"								
	iArrField(1) = "B.MINOR_NM"								

	iArrHeader(0) = "순번"							
	iArrHeader(1) = "운송회사명"						

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	frm1.txtTransCo.focus
	
	If iArrRet(0) <> "" Then
		frm1.txtTransCo.value = iArrRet(1)
		frm1.txtTransInfoNo.value = ""
	End If
End Function

'========================================
Function OpenVehicleNo()
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgBlnOpenPop = True Then Exit Function
		
	lgBlnOpenPop = True

	iArrParam(0) = "차량번호"							
	iArrParam(1) = "B_MAJOR A , B_MINOR B"						
	iArrParam(2) = ""			
	iArrParam(3) = ""									
	iArrParam(4) = " A.MAJOR_CD = B.MAJOR_CD AND B.MAJOR_CD = " & FilterVar("B9032", "''", "S") & " "				
	iArrParam(5) = "차량번호"							

	iArrField(0) = "B.MINOR_CD"								
	iArrField(1) = "B.MINOR_NM"								

	iArrHeader(0) = "순번"							
	iArrHeader(1) = "차량번호"						

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	frm1.txtVehicleNo.focus
	
	If iArrRet(0) <> "" Then
		frm1.txtVehicleNo.value = iArrRet(1)
		frm1.txtTransInfoNo.value = ""
	End If
End Function

' Spread button popup
'===========================================
Function OpenSpreadPopup(ByVal pvIntCol, ByVal pvIntRow)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenSpreadPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	With frm1.vspdData
		.Row = pvIntRow		:	.Col = pvIntCol - 1
		
		Select Case pvIntCol
			' 창고 
			Case C_SL_CD_POP
				iArrParam(1) = "dbo.B_STORAGE_LOCATION"		' FROM Clause
				iArrParam(2) = .Text													' Code Condition
				iArrParam(3) = ""														' Name Cindition
				.Col = C_PLANT_CD			' 공장 
				iArrParam(4) = "PLANT_CD = "	& FilterVar(.Text, "''", "S") & ""		' Where Condition
				
				iArrField(0) = "ED15" & Parent.gColSep & "SL_CD"		' 창고 
				iArrField(1) = "ED30" & Parent.gColSep & "SL_NM"		' 창고명 

				.Row = 0
				.Col = C_SL_CD	: iArrHeader(0) = .Text 			' Header명(0)
				.Col = C_SL_NM	: iArrHeader(1) = .Text			' Header명(1)
		End Select
	End With
 
	iArrParam(0) = iArrHeader(0)							' 팝업 명칭 
	iArrParam(5) = iArrHeader(0)							' 조회조건 TextBox 명칭 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	If iArrRet(0) <> "" Then
		OpenSpreadPopup = SetSpreadPopup(iArrRet,pvIntCol, pvIntRow)
	End If	

End Function

'========================================
Function SetConPopUp(ByVal pvArrRet,ByVal pvIntWhere)

	With frm1
		Select Case pvIntWhere
			Case C_PopPlant
				.txtConPlant.value = pvArrRet(0)
				.txtConPlantNm.value = pvArrRet(1) 

			Case C_PopDnType
				.txtConDnType.value = pvArrRet(0)
				.txtConDnTypeNm.value = pvArrRet(1) 

			Case C_PopShipToParty
				.txtConShipToParty.value = pvArrRet(0)
				.txtConShipToPartyNm.value = pvArrRet(1) 

			Case C_PopSalesGrp
				.txtConSalesGrp.value = pvArrRet(0)
				.txtConSalesGrpNm.value = pvArrRet(1) 

			Case C_PopTransMeth
				.txtTransMeth.value = pvArrRet(0)
				.txtTransMethNm.value = pvArrRet(1) 
		End Select
	End With

End Function

'========================================
Function SetPopUp(ByVal pvArrRet,ByVal pvIntWhere)

	With frm1
		Select Case pvIntWhere
			Case C_PopTransMeth
				.txtTransMeth.value = pvArrRet(0)
				.txtTransMethNm.value = pvArrRet(1) 

			Case C_PopInvMgr
				.txtInvMgr.value = pvArrRet(0)
				.txtInvMgrNm.value = pvArrRet(1) 
		End Select
	End With
	lgBlnFlgChgValue = True
End Function

'========================================
Function SetSpreadPopup(Byval pvArrRet,ByVal pvIntCol, ByVal pvIntRow)
	SetSpreadPopup = False

	With frm1.vspdData
		.Row = pvIntRow		:	.Col = pvIntCol - 1
		.Text = pvArrRet(0)
		
		Select Case pvIntCol
			Case C_SL_CD_POP
				.Col = C_SL_NM	: .Text = pvArrRet(1)
		End Select
	End With
	
	SetSpreadPopup = True
End Function

' 재고할당 여부를 Fetch한다.
'=========================================
Sub GetAllocInvFlag()
	Dim iStrSelectList, iStrFromList, iStrWhereList, iStrRs
	Dim iArrRs 

	iStrSelectList = "REFERENCE"
	iStrFromList = "dbo.B_CONFIGURATION"
	iStrWhereList = "MAJOR_CD = " & FilterVar("S0017", "''", "S") & " AND MINOR_CD = " & FilterVar("A", "''", "S") & "  AND SEQ_NO = 1 "

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList, iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		lgStrAllocInvFlag = iArrRs(1)
	Else
		err.Clear
		lgStrAllocInvFlag = "N"
	End If
End Sub

' Description : 창고명을 Fetch한다.
'===========================================
Function GetSlNm(ByVal pvIntRow)

	Dim iStrPlantCd, iStrSlCd
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs, iArrRs

	GetSlNm = False

	With frm1.vspdData
		.Row = pvIntRow
		.Col = C_PLANT_CD	:	iStrPlantCd = .text		' 공장 
		.Col = C_SL_CD		:	iStrSlCd = .text		' 창고 
	End With
	
	iStrSelectList = " SL_CD, SL_NM "
	iStrFromList = " dbo.B_STORAGE_LOCATION "
	iStrWhereList = " PLANT_CD =  " & FilterVar(iStrPlantCd , "''", "S") & " AND SL_CD =  " & FilterVar(iStrSlCd , "''", "S") & ""
			    
	'품목정보 Fetch
	If CommonQueryRs2by2(iStrSelectList, iStrFromList, iStrWhereList, iStrRs) Then
		iStrRs = Replace(iStrRs, parent.gColSep & parent.gRowSep, "")
		iArrRs = Split(Mid(iStrRs, 2), parent.gColSep)
		GetSlNm = SetSpreadPopup(iArrRs, C_SL_CD_POP, pvIntRow)
	Else
		If Err.number = 0 Then
			'Editing한경우 해당 창공정보가 존재하지 않으면 창고 Popup을 Display한다.
			GetSlNm = OpenSpreadPopup(C_SL_CD_POP, pvIntRow)
		Else
			MsgBox Err.description, vbObjectError, Parent.gLogoName 
			Err.Clear
		End If
			
		Call SetActiveCell(frm1.vspdData, C_SL_CD, pvIntRow,"M","X","X")
	End if
End Function

'=====================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_SELECT			= iCurColumnPos(1)
			C_PROMISE_DT		= iCurColumnPos(2)
			C_SHIP_TO_PARTY		= iCurColumnPos(3)
			C_BP_NM				= iCurColumnPos(4)			
			C_ITEM_CD			= iCurColumnPos(5)	
			C_ITEM_NM			= iCurColumnPos(6)
			C_REMAIN_QTY		= iCurColumnPos(7)	'SC.CFM_QTY - SC.REQ_QTY AS '잔량', 
			C_BONUS_REMAIN_QTY	= iCurColumnPos(8)	'SC.CFM_BONUS_QTY - SC.REQ_BONUS_QTY AS '덤잔량',
			C_SO_UNIT			= iCurColumnPos(9)
			C_GI_QTY			= iCurColumnPos(10)
			C_GI_BONUS_QTY		= iCurColumnPos(11)
			C_PLANT_CD			= iCurColumnPos(12)
			C_PLANT_NM			= iCurColumnPos(13)
			C_SL_CD				= iCurColumnPos(14)
			C_SL_CD_POP			= iCurColumnPos(15)
			C_SL_NM				= iCurColumnPos(16)	
			C_SU_ONHAND_QTY		= iCurColumnPos(17)	 	
			C_ONHAND_QTY		= iCurColumnPos(18)	'OS.GOOD_ON_HAND_QTY - OS.PICKING_QTY AS '현재고량',
			C_BASIC_UNIT		= iCurColumnPos(19) 
			C_SO_NO				= iCurColumnPos(20) 
			C_SO_SEQ			= iCurColumnPos(21) 
			C_SO_SCHD_NO		= iCurColumnPos(22)	
			C_TRACKING_NO		= iCurColumnPos(23) 
			C_SPEC				= iCurColumnPos(24)	
			C_DN_TYPE			= iCurColumnPos(25)	
			C_DN_TYPE_NM		= iCurColumnPos(26)	
			C_REMARK			= iCurColumnPos(27)	
			C_SO_TYPE			= iCurColumnPos(28)	
			C_SALES_GRP			= iCurColumnPos(29)	
			C_OLD_SL_CD			= iCurColumnPos(30)	
			C_OLD_SL_NM			= iCurColumnPos(31)	
			C_OLD_GI_QTY		= iCurColumnPos(32)	
			C_OLD_GI_BONUS_QTY	= iCurColumnPos(33)	
    End Select    
End Sub

'========================================
Sub GetCntryCd
	Dim iStrSelectList, iStrFromList, iStrWhereList, iStrRs
	Dim iArrRs 

	iStrSelectList = "CONTRY_CD"
	iStrFromList = "dbo.B_BIZ_PARTNER BP (NOLOCK) INNER JOIN dbo.B_COUNTRY CT (NOLOCK) ON (CT.COUNTRY_CD = BP.CONTRY_CD)"
	iStrWhereList = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") " & _
					"AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BPF (NOLOCK) WHERE BPF.PARTNER_BP_CD = BP.BP_CD AND BPF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ") " & _
					"AND BP_CD =  " & FilterVar(frm1.txtHConShipToParty.value, "''", "S") & ""

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList, iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		frm1.txtHCntryCd.value = iArrRs(1)
	Else
		frm1.txtHCntryCd.value = ""
		err.Clear
	End If
End Sub

'========================================
Sub GetTransMethInfo
	Dim iStrSelectList, iStrFromList, iStrWhereList, iStrRs
	Dim iArrRs 

	iStrSelectList = "MINOR_CD, MINOR_NM"
	iStrFromList = "dbo.B_MINOR (NOLOCK)"
	iStrWhereList = "MAJOR_CD = " & FilterVar("B9009", "''", "S") & " " & _
					"AND MINOR_CD =  " & FilterVar(frm1.txtTransMeth.value, "''", "S") & ""

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList, iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		frm1.txtTransMeth.value = iArrRs(1)
		frm1.txtTransMethNm.value = iArrRs(2)
	Else
		frm1.txtTransMethNm.value = ""
		Err.Clear
	End If
End Sub

'========================================
Sub GetInvMgrInfo
	Dim iStrSelectList, iStrFromList, iStrWhereList, iStrRs
	Dim iArrRs 

	iStrSelectList = "MINOR_CD, MINOR_NM"
	iStrFromList = "dbo.B_MINOR (NOLOCK)"
	iStrWhereList = "MAJOR_CD = " & FilterVar("I0004", "''", "S") & " " & _
					"AND MINOR_CD =  " & FilterVar(frm1.txtInvMgr.value, "''", "S") & ""

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList, iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		frm1.txtInvMgr.value = iArrRs(1)
		frm1.txtInvMgrNm.value = iArrRs(2)
	Else
		frm1.txtInvMgrNm.value = ""
		Err.Clear
	End If
End Sub

'========================================
Sub Form_Load()
	Call LoadInfTB19029              '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitSpreadSheet

	Call SetDefaultVal
    Call SetToolbar("11100000000011")										'⊙: 버튼 툴바 제어 
	Call InitVariables
	Call GetAllocInvFlag
	
	frm1.vspdDataH.MaxCols = frm1.vspdData.MaxCols + 1
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================
Function ClickTab1()
 
	If gSelframeFlg = TAB1 Then Exit Function
	
	Call ChangeTabs(TAB1)
	 
	gSelframeFlg = TAB1
End Function

'=========================================
Function ClickTab2()

	' 조회된 경우에만 변경 가능 
	If lgIntFlgMode <> Parent.OPMD_UMODE Then Exit Function
	 
	If gSelframeFlg = TAB2 Then Exit Function

	Call ChangeTabs(TAB2)
 
	gSelframeFlg = TAB2
End Function

'========================================
Sub txtConReqDateFrom_DblClick(Button)
	If Button = 1 Then
		frm1.txtConReqDateFrom.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtConReqDateFrom.focus
	End If
End Sub

'========================================
Sub txtConReqDateTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtConReqDateTo.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtConReqDateTo.focus
	End If
End Sub

'========================================
Sub txtActualGIDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtActualGiDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtActualGiDt.focus
	End If
End Sub

'========================================
Sub txtConReqDateFrom_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()	
End Sub

'========================================
Sub txtConReqDateTo_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'========================================
Sub chkArFlag_OnClick()
	If Not frm1.chkArFlag.checked Then
		frm1.chkVatFlag.checked = False
	End If
End Sub

'========================================
Sub chkVatFlag_OnClick()
	If frm1.chkVatFlag.checked Then
		frm1.chkArFlag.checked = True
	End If
End Sub

' 전체선택 
'========================================
Sub chkSelectAll_onClick()
	Dim iStrOldValue
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	ggoSpread.Source = frm1.vspdData	
	With frm1.vspdData
		.Row = 1			:	.Row2 = .MaxRows
		
		' 전체선택 
		If frm1.chkSelectAll.checked Then
			' Row Header 설정(수정)
			.Col = 0			:	.Col2 = 0
			.Clip = Replace(.Clip, vbCrLf, ggoSpread.UpdateFlag & vbCrLf)
			
			' 선택버튼의 선택여부 설정 
			.Col = C_SELECT		:	.Col2 = C_SELECT
			.Clip = Replace(.Clip, "0", "1")
			
			Call SetSpreadColor(1, .MaxRows, 1)
			
		' 전체선택 취소 
		Else
			' Row Header 설정(수정)
			.Col = 0			:	.Col2 = 0
			.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")

			.Col = C_SELECT		:	.Col2 = C_SELECT
			.Clip = Replace(.Clip, "1", "0")
			
			Call RestoreDataByClip(C_OLD_GI_QTY,C_GI_QTY)				'출고수량 
			Call RestoreDataByClip(C_OLD_GI_BONUS_QTY,C_GI_BONUS_QTY)	'덤출고수량 
			Call RestoreDataByClip(C_OLD_SL_CD,C_SL_CD)					'창고 
			Call RestoreDataByClip(C_OLD_SL_NM,C_SL_NM)					'창고명 
			Call SetSpreadColor(1, .MaxRows, 0)
		End if
	End With

	' Active Cell 설정	
	Call SetActiveCell(frm1.vspdData,C_SELECT, 1,"M","X","X")
End Sub

' 전체전택 취소시 이전 값으로 복원 
'========================================
Sub RestoreDataByClip(ByVal pvIntOldCol, ByVal pvIntCol)
	Dim iStrClip
	
	With frm1.vspdData
		' 변경전 값 
		.Col = pvIntOldCol	:	.Col2 = pvIntOldCol
		iStrClip = .Clip
			
		.Col = pvIntCol		:	.Col2 = pvIntCol
		.Clip = iStrClip
	End With	
End Sub

'========================================
Function txtTransMeth_OnChange()
	If Trim(frm1.txtTransMeth.value) = "" Then
		frm1.txtTransMethNm.value = ""
	Else
		Call GetTransMethInfo
	End If
End Function

'========================================
Function txtInvMgr_OnChange()
	If Trim(frm1.txtInvMgr.value) = "" Then
		frm1.txtInvMgrNm.value = ""
	Else
		Call GetInvMgrInfo
	End If
End Function

'========================================
Sub txtArrivalDt_Change()
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtArrivalTime_OnChange()
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtRemark_OnChange()
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtSTPInfoNo_OnChange()
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtZIPcd_OnChange()
	frm1.txtSTPInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtReceiver_OnChange()
	frm1.txtSTPInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtAddr1_OnChange()
	frm1.txtSTPInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtAddr2_OnChange()
	frm1.txtSTPInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtAddr3_OnChange()
	frm1.txtSTPInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtTelNo1_OnChange()
	frm1.txtSTPInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtTelNo2_OnChange()
	frm1.txtSTPInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtShipToPlace_OnChange()
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtTransInfoNo_OnChange()
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtTransCo_OnChange()
	frm1.txtTransInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtSender_OnChange()
	frm1.txtTransInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtVehicleNo_OnChange()
	frm1.txtTransInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Sub txtDriver_OnChange()
	frm1.txtTransInfoNo.value = ""
	lgBlnFlgChgValue3 = True
End Sub

'========================================
Function btnShipToPlceRef_OnClick()
	Dim iCalledAspName
	Dim iArrRet
	Dim iStrShipToParty
	
	On Error Resume Next

	If lgBlnOpenPop = True Then Exit Function

	lgBlnOpenPop = True

	iCalledAspName = AskPRAspName("S4111RA1")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4111RA1", "x")
		lgBlnOpenPop = False
		exit Function
	end if
	
	iStrShipToParty = Trim(frm1.txtHConShipToParty.value)
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent , iStrShipToParty),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) = "" Then
		Err.Clear 
	Else
		With frm1
			.txtSTPInfoNo.value = iArrRet(0)			'납품처상세정보번호	
			.txtZIPcd.value = iArrRet(1)				'우편번호 
			.txtADDR1.value = iArrRet(2)				'납품주소1
			.txtAddr2.value = iArrRet(3)				'납품주소2	
			.txtADDR3.value = iArrRet(4)				'납품주소3
			.txtReceiver.value = iArrRet(5)				'인수자명 
			.txtTelNo1.value = iArrRet(6)				'전화번호1
			.txtTelNo2.value = iArrRet(7)				'전화번호2
			lgBlnFlgChgValue3 = True
		End With
	End If	

End Function

'========================================
Function btnTrnsMethRef_OnClick()
	Dim iCalledAspName
	Dim iArrRet
	
	On Error Resume Next

	If lgBlnOpenPop = True Then Exit Function

	lgBlnOpenPop = True
	
	iCalledAspName = AskPRAspName("S4111RA2")
		
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "S4111RA2", "x")
		lgBlnOpenPop = False
		exit Function
	end if

	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent , ""),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) = "" Then
		Err.Clear 
	Else
		frm1.txtTransInfoNo.value = iArrRet(0)			'운송정보번호 
		frm1.txtTransCo.value = iArrRet(1)				'운송회사 
		frm1.txtDriver.value = iArrRet(2)				'운전자명 
		frm1.txtVehicleNo.value = iArrRet(3)			'차량번호	
		frm1.txtSender.value = iArrRet(4)				'인계자명 
		lgBlnFlgChgValue3 = True
	End If	

End Function

'========================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If lgIntFlgMode = Parent.OPMD_CMODE Then Exit Sub

	ggoSpread.Source = frm1.vspdData
	
	If Row > 0 Then
		Select Case Col
		Case C_SELECT
			If ButtonDown = 0 then					'---선택이 취소된 경우				
				frm1.vspddata.row = Row
				Call FncCancel()
				Call SetSpreadColor(Row, Row, ButtonDown)
			Else									'--- 선택된 경우 
				ggoSpread.UpdateRow Row	
				Call SetSpreadColor(Row, Row, ButtonDown)
			End If			
			
		Case C_SL_CD_POP
			Call OpenSpreadPopup(Col, Row)

			Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
		End Select
	End If

End Sub

'=======================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
	
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    	

End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_SELECT Or NewCol <= C_SELECT Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================
Sub vspdData_Change(ByVal Col , ByVal Row)

	With frm1.vspdData
		If Trim(.Text) = "" Then Exit Sub
		
		Select Case Col
			Case C_SL_CD
				If Not GetSlNm(Row) Then
					.Row = Row
					.Col = Col : .Text = ""
				End If
		End Select
	End With
End Sub

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
			If CheckRunningBizProcess Then Exit Sub
			Call DisableToolBar(Parent.TBC_QUERY)
            Call DbQuery()
        End If
    End if
End Sub

'=====================================================
Function FncQuery() 

	On Error Resume Next
	    
    Dim IntRetCD 
        
    FncQuery = False                                                        
    
    Err.Clear                                                               
	
	' 조회조건의 입력필수 항목 check
    If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(frm1.txtConReqDateFrom, frm1.txtConReqDateTo) = False Then Exit Function

    If lgBlnFlgChgValue3 = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoOper.ClearField(Document, "3")          

    Call ggoSpread.ClearSpreadData()
	frm1.chkSelectAll.checked = False
    Call InitVariables															

    Call DbQuery																<%'☜: Query db data%>

    FncQuery = True																
        
End Function

'=====================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          

    If lgBlnFlgChgValue Or lgBlnFlgChgValue3 Or ggoSpread.SSCheckChange Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
	Call ClickTab1
    Call SetDefaultVal
    Call InitVariables															

    Call SetToolbar("11100000000011")										'⊙: 버튼 툴바 제어 

    FncNew = True																

End Function

'=====================================================
Function FncSave() 
    FncSave = False                                                         
    
    Err.Clear                                                               
    
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")		
	    Exit Function
    End If

	ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") OR ggoSpread.SSDefaultCheck = False Then
		Call ClickTab1
		frm1.vspdData.focus
		Set gActiveElement = document.ActiveElement
       Exit Function
    End If

	If UniConvDateToYYYYMMDD(frm1.txtActualGIDt.text , parent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate , parent.gDateFormat , "") Then  
		Call DisplayMsgBox("970024", "X", frm1.txtActualGIDt.ALT, "현재일") 
		Call ClickTab1
		Call SetFocusToDocument("M")	
		frm1.txtActualGIDt.focus
		Set gActiveElement = document.ActiveElement
		Exit Function
	End If

    CAll DbSave
    
    FncSave = True                                                          
    
End Function

'=====================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	Dim iLngLoop, iLngFirstRow, iLngLastRow
	Dim iLngActiveRow, iLngActiveCol
	
	With frm1.vspdData 
		iLngActiveRow = .ActiveRow	:	iLngActiveCol = .ActiveCol
		
		iLngFirstRow = .SelBlockRow
		If iLngFirstRow = -1 Then
			iLngFirstRow = 1
			iLngLastRow = .MaxRows
		Else
			iLngLastRow = .SelBlockRow2
		End If
	End With

	ggoSpread.Source = frm1.vspdData 

	For iLngLoop = iLngFirstRow To iLngLastRow
		' Active Cell 설정	
		Call SetActiveCell(frm1.vspdData,iLngActiveCol, iLngLoop,"M","X","X")
		
	    ggoSpread.EditUndo
	Next
	
	Call SetActiveCell(frm1.vspdData,iLngActiveCol, iLngFirstRow,"M","X","X")
End Function

'=====================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'=====================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function

'=====================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
End Function

'=====================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'=====================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=====================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub

'=====================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'=====================================================
Function DbQuery() 

    On Error Resume Next                                                          
    Err.Clear
    
	If LayerShowHide(1) = False Then
		Exit Function 
    End If
	  
	Dim iStrVal
	
    DbQuery = False
    
    With frm1
		
		' 재쿼리시(Scrollbar)
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			iStrVal = BIZ_PGM_ID & "?txtMode="			& Parent.UID_M0001							
			iStrVal = iStrVal & "&txtConPlant="			& Trim(.txtHConPlant.value)			
			iStrVal = iStrVal & "&txtConReqDateFrom="	& Trim(.txtHConReqDateFrom.value)
			iStrVal = iStrVal & "&txtConReqDateTo="		& Trim(.txtHConReqDateTo.value)		
			iStrVal = iStrVal & "&txtConDnType="		& Trim(.txtHConDnType.value)			
			iStrVal = iStrVal & "&txtConShipToParty="	& Trim(.txtHConShipToParty.value)		
			iStrVal = iStrVal & "&txtConSalesGrp="		& Trim(.txtHConSalesGrp.value)
			iStrVal = iStrVal & "&txtConFrSoNo="		& Trim(.txtHConFrSoNo.value)
			iStrVal = iStrVal & "&txtConToSoNo="		& Trim(.txtHConToSoNo.value)
			iStrVal = iStrVal & "&lgStrPrevKey="		& lgStrPrevKey
		Else
			iStrVal = BIZ_PGM_ID & "?txtMode="			& Parent.UID_M0001						
			iStrVal = iStrVal & "&txtConPlant="			& Trim(.txtConPlant.value)			
			iStrVal = iStrVal & "&txtConReqDateFrom="	& Trim(.txtConReqDateFrom.text)
			iStrVal = iStrVal & "&txtConReqDateTo="		& Trim(.txtConReqDateTo.text)		
			iStrVal = iStrVal & "&txtConDnType="		& Trim(.txtConDnType.value)			
			iStrVal = iStrVal & "&txtConShipToParty="	& Trim(.txtConShipToParty.value)		
			iStrVal = iStrVal & "&txtConSalesGrp="		& Trim(.txtConSalesGrp.value)
			iStrVal = iStrVal & "&txtConFrSoNo="		& Trim(.txtConFrSoNo.value)
			iStrVal = iStrVal & "&txtConToSoNo="		& Trim(.txtConToSoNo.value)
			iStrVal = iStrVal & "&lgStrPrevKey="
		End if
		
		If .chkBatchQuery.checked Then
			iStrVal = iStrVal & "&txtBatchQuery=Y"
		Else
			iStrVal = iStrVal & "&txtBatchQuery=N"
		End If
		
		iStrVal = iStrVal & "&txtLastRow=" & .vspdData.MaxRows
		
    End With

	Call RunMyBizASP(MyBizASP, iStrVal)											
               
    If Err.number = 0 Then	 
       DbQuery = True                                                           
    End If

    Set gActiveElement = document.ActiveElement    
    
End Function

'=====================================================
Function DbQueryOk()
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		' 납품처에 대한 국가코드를 Fetch한다.
		Call GetCntryCd
		Call ClickTab1
		
		lgBlnFlgChgValue = False
		lgBlnFlgChgValue3 = False
	
		Call SetToolbar("11101001000111")	

	    lgIntFlgMode = Parent.OPMD_UMODE
	End If
	
	frm1.vspdData.focus
End Function

'=====================================================
Function DbSave() 
	On Error Resume Next

    Err.Clear																

    DbSave = False      
    
	If LayerShowHide(1) = False Then Exit Function 

	Dim iIntRow
	Dim iArrColData
	Dim iStrIns
	
	Dim iColSep, iRowSep, iFormLimitByte, iChunkArrayCount
	Dim iLngCTotalvalLen		'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장 

	Dim iTmpCBuffer				'현재의 버퍼 [수정,신규] 
	Dim iTmpCBufferCount		'현재의 버퍼 Position
	Dim iTmpCBufferMaxCount		'현재의 버퍼 Chunk Size

	' 속도 향상을 위해 Local 변수로 재정의 
	iColSep = parent.gColSep
	iRowSep = parent.gRowSep
	iFormLimitByte = parent.C_FORM_LIMIT_BYTE
	iChunkArrayCount = parent.C_CHUNK_ARRAY_COUNT
	
	iTmpCBufferMaxCount = iChunkArrayCount '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpCBufferCount = -1
	iLngCTotalvalLen = 0
	
	ReDim iTmpCBuffer(iTmpCBufferMaxCount)	'최기 버퍼의 설정[신규]
	Redim iArrColData(14)

	' vspdData의 Data를 Hidden Spread로 복사한다.
	Call CopyVspdDataToVspdDataH
	
	With frm1.vspdDataH
		'-----------------------
		'Data manipulate area
		'-----------------------
		For iIntRow = 1 To .MaxRows    
		    .Row = iIntRow
			.Col = .MaxCols			: iArrColData(0) = .Text					' Row 번호 
			.Col = C_SHIP_TO_PARTY	: iArrColData(1) = Trim(.Text)				' 1 : 납품처 
			.Col = C_SO_NO			: iArrColData(2) = Trim(.Text)				' 2 : 수주번호 
			.Col = C_SO_SEQ			: iArrColData(3) = Trim(.Text)				' 3 : 수주순번 
			.Col = C_SO_SCHD_NO		: iArrColData(4) = Trim(.Text)				' 4 : 납품순번 
			.Col = C_GI_QTY			: iArrColData(5) = UNIConvNum(.Text,0)		' 5 : 출고량 
			.Col = C_GI_BONUS_QTY	: iArrColData(6) = UNIConvNum(.Text,0)		' 6 : 덤출고량 
			.Col = C_SL_CD			: iArrColData(7) = Trim(.Text)				' 7 : 창고 
			.Col = C_ITEM_CD		: iArrColData(8) = Trim(.Text)				' 8 : 품목 
			.Col = C_PLANT_CD		: iArrColData(9) = Trim(.Text)				' 9 : 공장 
			.Col = C_SO_UNIT		: iArrColData(10) = Trim(.Text)				' 10 : 수주단위 
			.Col = C_REMARK			: iArrColData(11) = Trim(.Text)				' 11 : 비고 
			.Col = C_DN_TYPE		: iArrColData(12) = Trim(.Text)				' 12 : 출하형태 
			.Col = C_SO_TYPE		: iArrColData(13) = Trim(.Text)				' 13 : 수주형태 
			.Col = C_SALES_GRP		: iArrColData(14) = Trim(.Text)				' 14 : 영업그룹 
				
			iStrIns = Join(iArrColData, iColSep) & iRowSep

			If iLngCTotalvalLen + Len(iStrIns) >  iFormLimitByte Then			'한개의 form element에 넣을 Data 한개치가 넘으면 
				Call MakeTextArea("txtCSpread", iTmpCBuffer)
							
			   iTmpCBufferMaxCount = iChunkArrayCount			                ' 임시 영역 새로 초기화 
			   ReDim iTmpCBuffer(iTmpCBufferMaxCount)
			   iTmpCBufferCount = -1
			   iLngCTotalvalLen  = 0
			End If
						   
			iTmpCBufferCount = iTmpCBufferCount + 1
						  
			If iTmpCBufferCount > iTmpCBufferMaxCount Then                      ' 버퍼의 조정 증가치를 넘으면 
			   iTmpCBufferMaxCount = iTmpCBufferMaxCount + iChunkArrayCount		' 버퍼 크기 증성 
			   ReDim Preserve iTmpCBuffer(iTmpCBufferMaxCount)
			End If   
			iTmpCBuffer(iTmpCBufferCount) =  iStrIns         
			iLngCTotalvalLen = iLngCTotalvalLen + Len(iStrIns)
		Next

		' Hidden Object Clear
		.MaxRows = 0
	End With

   ' 나머지 데이터 처리 
	If iTmpCBufferCount > -1 Then Call MakeTextArea("txtCSpread", iTmpCBuffer)

	With frm1
		.txtMode.value = Parent.UID_M0002
		
		' 후속 작업여부 설정(매출채권)
		If .chkArFlag.checked Then
			.txtHArflag.value = "Y"
		Else
			.txtHArflag.value = "N"
		End If
		
		' 후속 작업여부 설정(세금계산서)
		If .chkVatFlag.checked Then
			.txtHVatFlag.value = "Y"
		Else
			.txtHVatFlag.value = "N"
		End If
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
    DbSave = True                                                           
    
End Function

'=====================================================
Function DbSaveOk()
	
	Call ggoSpread.ClearSpreadData()
    Call InitVariables
    Call RemovedivTextArea
    Call MainQuery()

End Function

'========================================
Sub MakeTextArea(ByVal pvStrName, ByRef prArrData)
	Dim iObjTEXTAREA		'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Set iObjTEXTAREA = document.createElement("TEXTAREA")            '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
	iObjTEXTAREA.name = pvStrName
	iObjTEXTAREA.value = Join(prArrData,"")
	divTextArea.appendChild(iObjTEXTAREA)
End Sub

'========================================
Function RemovedivTextArea()
	Dim iIntIndex
	
	For iIntIndex = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Function

'CopyVspdDataToVspdDataH
'========================================
Sub CopyVspdDataToVspdDataH()
	Dim iIntRow
	
	' Hidden Object로 저장할 데이터 복사 
	frm1.vspdData.Col = 1
	frm1.vspdData.Col2 = frm1.vspdData.MaxCols

	frm1.vspdDataH.Col = 1
	frm1.vspdDataH.Col2 = frm1.vspdDataH.MaxCols

	For iIntRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row = iIntRow
		frm1.vspdData.Row2 = iIntRow
							
		If frm1.vspdData.Text = "1" Then
			frm1.vspdDataH.MaxRows = frm1.vspdDataH.MaxRows + 1
			frm1.vspdDataH.Row = frm1.vspdDataH.MaxRows
			frm1.vspdDataH.Row2 = frm1.vspdDataH.MaxRows
			frm1.vspdDataH.Clip = Replace(frm1.vspdData.Clip, vbCrLf, vbTab & iIntRow & vbCrLf)
		End if		
	Next

	' Hidden Object내에서 Data 정렬 
	Call SortvspdDataH

End Sub

'========================================
Sub SortvspdDataH()
	Dim iArrSortKeys, iArrSortKeyOrder
	' 출하형태, 수주유형, 영업그룹, 납품처로 정렬한다 
	iArrSortKeys = Array(C_DN_TYPE, C_SO_TYPE, C_SALES_GRP, C_SHIP_TO_PARTY, C_PROMISE_DT, C_ITEM_CD)
	iArrSortKeyOrder = Array(1, 1, 1, 1, 1, 1)

	With frm1.vspdDataH
		.Sort 1, 1,.MaxCols, .MaxRows, 0, iArrSortKeys, iArrSortKeyOrder
	End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>일괄출고</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납품 및 운송정보</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6><INPUT NAME="txtConPlant" TYPE="Text" Alt="공장" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopPlant">&nbsp;<INPUT NAME="txtConPlantNm" TYPE="Text" Alt="공장명" SIZE=25 tag="14"></TD>									
									<TD CLASS="TD5" NOWRAP>출고예정일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s4111ma3_fpDateTime1_txtConReqDateFrom.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s4111ma3_fpDateTime2_txtConReqDateTo.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출하형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConDnType" TYPE="Text" MAXLENGTH="3" SIZE=10 Alt="출하형태" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopDnType">&nbsp;<INPUT NAME="txtConDnTypeNm" TYPE="Text" Alt="출하형태명" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6><INPUT NAME="txtConShipToParty" TYPE="Text" Alt="납품처" MAXLENGTH=10 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopShipToParty">&nbsp;<INPUT NAME="txtConShipToPartyNm" TYPE="Text" Alt="납품처명" SIZE=25 tag="14"></TD>									
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6><INPUT NAME="txtConSalesGrp" TYPE="Text" Alt="영업그룹" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSalesGrp">&nbsp;<INPUT NAME="txtConSalesGrpNm" TYPE="Text" Alt="영업그룹명" SIZE=25 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP>수주번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConFrSoNo" TYPE="Text" MAXLENGTH="18" SIZE=18 Alt="수주번호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSoNo frm1.txtConFrSoNO">&nbsp;~&nbsp;<INPUT NAME="txtConToSoNo" TYPE="Text" MAXLENGTH="18" SIZE=18 Alt="수주번호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSoNo frm1.txtConToSoNo"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>일괄조회</TD>
									<TD CLASS=TD6>
										<INPUT TYPE=CHECKBOX NAME="chkBatchQuery" ID="chkBatchQuery" tag="11" Class="Check">
									</TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>실제출고일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4111ma3_fpDateTime1_txtActualGIDt.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>운송방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransMeth" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="운송방법" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopUp C_PopTransMeth">&nbsp;<INPUT NAME="txtTransMethNm" TYPE="Text" Alt="운송방법명" SIZE=25 tag="24"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>재고담당자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvMgr" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="21XXXU" ALT="재고담당자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInvMgr" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopup C_PopInvMgr">&nbsp;<INPUT NAME="txtInvMgrNm" TYPE="Text" SIZE=25 tag="24" ALT="재고담당자명"></TD>
									<TD CLASS=TD5 NOWRAP>후속작업여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=CHECKBOX NAME="chkArFlag" tag="21" Class="Check"><LABEL ID="lblArFlag" FOR="chkArFlag">매출채권</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkVatFlag" tag="21" Class="Check"><LABEL ID="lblVatFlag" FOR="chkVatFlag">세금계산서</LABEL>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>전체선택</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=CHECKBOX NAME="chkSelectAll" ID="chkSelectAll" tag="21" Class="Check">
									</TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
										<script language =javascript src='./js/s4111ma3_vaSpread1_vspdData.js'></script>
									</TD>
								</TR>
							</TABLE>
						</DIV>
						<DIV ID="TabDiv" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>실제납품일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s4111ma3_fpDateTime2_txtArrivalDt.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>납품시간</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtArrivalTime" TYPE="Text" ALT="납품시간" MAXLENGTH="10" SIZE=36 tag="31"></TD>
								</TR>							
								<TR>	
									<TD CLASS=TD5 NOWRAP>비고</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtRemark" TYPE="Text" MAXLENGTH="120" SIZE=91 ALT="비고" tag="31"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품처상세정보번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSTPInfoNo" ALT="납품처상세정보번호" TYPE="Text" MAXLENGTH="18" SIZE=18 tag="34XXXU">&nbsp;<BUTTON NAME = "btnShipToPlceRef" CLASS="CLSMBTN">납품처상세정보참조</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>우편번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZIPCd" TYPE="Text" ALT="우편번호" MAXLENGTH="12" SIZE=20 tag="31XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnZIP_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenZip" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()"></TD>
									<TD CLASS=TD5 NOWRAP>인수자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtReceiver" TYPE="Text" ALT="인수자명" MAXLENGTH="50" SIZE=36 tag="31"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품주소</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtAddr1" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="31"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtAddr2" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="31"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtAddr3" TYPE="Text" ALT="납품주소" MAXLENGTH="100" SIZE=91 tag="31"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>납품장소</TD>
									<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtShipToPlace" ALT="납품장소" TYPE="Text" MAXLENGTH=30 SiZE=91 tag="31XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>전화번호1</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTelNo1" TYPE="Text" ALT="전화번호1" MAXLENGTH="20" SIZE=34 tag="31XXXU"></TD>
									<TD CLASS=TD5 NOWRAP>전화번호2</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTelNo2" TYPE="Text" ALT="전화번호2" MAXLENGTH="20" SIZE=34 tag="31XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>운송정보번호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransInfoNo" ALT="운송정보번호" TYPE="Text" MAXLENGTH="18" SIZE=18 tag="34XXXU">&nbsp;<BUTTON NAME = "btnTrnsMethRef" CLASS="CLSMBTN">운송정보참조</BUTTON></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
								<TR>							
									<TD CLASS=TD5>운송회사</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtTransCo" SIZE=20 MAXLENGTH=50 TAG="31XXXX" ALT="운송회사"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransCo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenTransCo()"></TD>
									<TD CLASS=TD5 NOWRAP>인계자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSender" TYPE="Text" ALT="인계자명" MAXLENGTH="50" SIZE=37 tag="31"></TD>
								</TR>
								<TR>							
									<TD CLASS=TD5>차량번호</TD>
									<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtVehicleNo" SIZE=20 MAXLENGTH=20 TAG="31XXXX" ALT="차량번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVehicleNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenVehicleNo()"></TD>							
									<TD CLASS=TD5 NOWRAP>운전자명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDriver" TYPE="Text" ALT="운전자명" MAXLENGTH="50" SIZE=37 tag="31"></TD>
								</TR>
								   <%Call SubFillRemBodyTD5656(1)%>
							</TABLE>
						</DIV>					
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
			<script language =javascript src='./js/s4111ma3_OBJECT1_vspdDataH.js'></script>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConPlant" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConReqDateFrom" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConReqDateTo" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConDnType" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConShipToParty" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSalesGrp" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConFrSoNo" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToSoNo" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHArFlag" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVatFlag" tag="34" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHCntryCd" tag="34" TABINDEX="-1">

<P ID="divTextArea"></P>
</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
