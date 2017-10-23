<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        :
'*  3. Program ID           : MM112QA1
'*  4. Program Name         : 멀티컴퍼니수발주조회(발주별)-멀티 
'*  5. Program Desc         : 멀티컴퍼니수발주조회(발주별)-멀티 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/24
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Han Kwang Soo
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
Dim interface_Production

Const BIZ_PGM_ID = "KMM112QB1.asp"
'Const BIZ_PGM_ID2 = "MM211QB101.asp"
'Const BIZ_PGM_SAVE_ID = "m2111mb5.asp"
'Const BIZ_PGM_JUMP_ID_PO_DTL = "MM211MA1.asp"
											'☆: 비지니스 로직 ASP명 
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Dim		C_SoCompany							'수주법인 
Dim		C_SoCompanyNm
Dim		C_PoNo								'발주번호 
Dim		C_PoSeqNo							'발주순번 
Dim		C_SoNo								'수주번호 
Dim		C_SoSeqNo							'수주순번 
Dim		C_ItemCd							'품목 
Dim		C_ItemNm							'품목명 
Dim		C_Spec								'품목규격 
Dim		C_PoSts								'발주법인상태 
Dim		C_PoStsNm							'발주법인상태 
Dim		C_SoSts								'수주법인상태 
Dim		C_SoStsNm							'수주법인상태 
Dim		C_PoUnit							'단위 
Dim		C_PoQty								'발주수량 
Dim		C_SoQty								'수주수량 
Dim		C_PoLcQty							'수입L/C수량 
Dim		C_SoLcQty							'수출L/C수량 
Dim		C_SoReqQty							'출하요청수량 
Dim		C_SoIssueQty						'출고수량 
Dim		C_SoCcQty							'수출통관수량 
Dim		C_PoCcQty							'수입통관수량 
Dim		C_PoRcptQty							'입고수량 
Dim		C_SoBillQty							'매출수량 
Dim		C_PoIvQty							'매입수량 


Dim lgSpdHdrClicked	'2003-03-01 Release 추가 
'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim lgIntFlgModeM                 'Variable is for Operation Status

Dim lgStrPrevKeyM			'Multi에서 재쿼리를 위한 변수 
Dim lglngHiddenRows		'Multi에서 재쿼리를 위한 변수	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.

Dim lgStrPrevKey1
Dim lgStrPrevKey2

Dim lgSortKey1
Dim lgSortKey2

Dim IsOpenPop
Dim lgCurrRow
Dim strInspClass

Dim lgPageNo1

Dim EndDate, StartDate,CurrDate, iDBSYSDate

' === 2005.07.22 수정 ===========================================================
StartDate   = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
StartDate   = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
EndDate     = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
' === 2005.07.22 수정 ===========================================================



'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
	lgIntFlgMode = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgIntFlgModeM = Parent.OPMD_CMODE		'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0						'initializes Group View Size

    lgStrPrevKey1 = ""						'initializes Previous Key
    lgStrPrevKey2 = ""						'initializes Previous Key

    lgLngCurRows = 0						'initializes Deleted Rows Count
    lgSortKey1 = 2
    lgSortKey2 = 2
    lgPageNo = 0
    lgPageNo1 = 0


    frm1.txtPoCompanyCd.value = parent.gCompany
	call  CommonQueryRs(" CO_FULL_NM "," B_COMPANY "," CO_CD = " & FilterVar(frm1.txtPoCompanyCd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    frm1.txtPoCompanyNm.value = Trim(Replace(lgF0,Chr(11),""))

    '###검사분류별 변경부분 Start###
    strInspClass = "R"
	'###검사분류별 변경부분 End###
    'ggoSpread.ClearSpreadData
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()

	frm1.txtPoFrDt.Text=StartDate
	frm1.txtPoToDt.Text=EndDate

	Call SetToolbar("1100000000001111")

    frm1.txtSoCompanyCd.focus

    Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    Call InitSpreadPosVariables()

	With frm1.vspdData

	ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030901",,Parent.gAllowDragDropSpread

	.ReDraw = false

    .MaxCols = C_PoIvQty + 1
    .Col = .MaxCols:	.ColHidden = True
    .MaxRows = 0

    Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit		C_SoCompany									  	,	"수주법인"    		,15
	ggoSpread.SSSetEdit		C_SoCompanyNm									,	"수주법인명"		,20
	ggoSpread.SSSetEdit		C_PoNo										  	,	"발주번호"    		,15
	ggoSpread.SSSetEdit		C_PoSeqNo									  	,	"발주순번"    		,12
	ggoSpread.SSSetEdit		C_SoNo										  	,	"수주번호"    		,15
	ggoSpread.SSSetEdit		C_SoSeqNo									  	,	"수주순번"    		,12
	ggoSpread.SSSetEdit		C_ItemCd									  	,	"품목"        		,15
	ggoSpread.SSSetEdit		C_ItemNm									  	,	"품목명"      		,20
	ggoSpread.SSSetEdit		C_Spec										  	,	"품목규격"    		,15
	ggoSpread.SSSetEdit		C_PoSts										  	,	"발주법인상태"		,8
	ggoSpread.SSSetEdit		C_PoStsNm									  	,	"발주법인상태"		,15
	ggoSpread.SSSetEdit		C_SoSts										  	,	"수주법인상태"		,8
	ggoSpread.SSSetEdit		C_SoStsNm									  	,	"수주법인상태"		,15
	ggoSpread.SSSetEdit		C_PoUnit									  	,	"단위"        		,10
	SetSpreadFloatLocal		C_PoQty										  	,	"발주수량"    		,15,1,5
	SetSpreadFloatLocal		C_SoQty										  	,	"수주수량"    		,15,1,5
	SetSpreadFloatLocal		C_PoLcQty									  	,	"수입L/C수량" 		,15,1,5
	SetSpreadFloatLocal		C_SoLcQty									  	,	"수출L/C수량" 		,15,1,5
	SetSpreadFloatLocal		C_SoReqQty									  	,	"출하요청수량"		,15,1,5
	SetSpreadFloatLocal		C_SoIssueQty								  	,	"출고수량"    		,15,1,5
	SetSpreadFloatLocal		C_SoCcQty									  	,	"수출통관수량"		,15,1,5
	SetSpreadFloatLocal		C_PoCcQty									  	,	"수입통관수량"		,15,1,5
	SetSpreadFloatLocal		C_PoRcptQty									  	,	"입고수량"    		,15,1,5
	SetSpreadFloatLocal		C_SoBillQty									  	,	"매출수량"    		,15,1,5
	SetSpreadFloatLocal		C_PoIvQty									  	,	"매입수량"    		,15,1,5

'	Call ggoSpread.MakePairsColumn(C_ItemCd,C_SpplSpec)
	Call ggoSpread.SSSetColHidden(C_PoSts,	C_PoSts,	True)
	Call ggoSpread.SSSetColHidden(C_SoSts,	C_SoSts,	True)

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

    .vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLock 1 , -1


    .vspdData.ReDraw = True

    End With
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
'    With frm1

'    .vspdData.ReDraw = False
'    ggoSpread.SSSetProtected		C_PlantCd, pvStartRow, pvEndRow


'    .vspdData.ReDraw = True

'    End With
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
' Function Name : InitSpreadPosVariables
' Function Desc : This method assign sequential number for Spreadsheet column position
'========================================================================================
Sub InitSpreadPosVariables()
	C_SoCompany			=	1			'수주법인 
	C_SoCompanyNm		=	2			'수주법인명 
	C_PoNo				=	3			'발주번호 
	C_PoSeqNo			=	4			'발주순번 
	C_SoNo				=	5			'수주번호 
	C_SoSeqNo			=	6			'수주순번 
	C_ItemCd			=	7			'품목 
	C_ItemNm			=	8			'품목명 
	C_Spec				=	9			'품목규격 
	C_PoSts				=	10  		'발주법인상태 
	C_PoStsNm			=	11  		'발주법인상태 
	C_SoSts				=	12  		'수주법인상태 
	C_SoStsNm			=	13  		'수주법인상태 
	C_PoUnit			=	14  		'단위 
	C_PoQty				=	15  		'발주수량 
	C_SoQty				=	16  		'수주수량 
	C_PoLcQty			=	17  		'수입L/C수량 
	C_SoLcQty			=	18  		'수출L/C수량 
	C_SoReqQty			=	19  		'출하요청수량 
	C_SoIssueQty		=	20  		'출고수량 
	C_SoCcQty			=	21  		'수출통관수량 
	C_PoCcQty			=	22  		'수입통관수량 
	C_PoRcptQty			=	23  		'입고수량 
	C_SoBillQty			=	24  		'매출수량 
	C_PoIvQty			=	25			'매입수량 
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

			C_SoCompany			=	iCurColumnPos(1)			'수주법인 
			C_SoCompanyNm		=	iCurColumnPos(2)
			C_PoNo				=	iCurColumnPos(3)			'발주번호 
			C_PoSeqNo			=	iCurColumnPos(4)			'발주순번 
			C_SoNo				=	iCurColumnPos(5)			'수주번호 
			C_SoSeqNo			=	iCurColumnPos(6)			'수주순번 
			C_ItemCd			=	iCurColumnPos(7)			'품목 
			C_ItemNm			=	iCurColumnPos(8)			'품목명 
			C_Spec				=	iCurColumnPos(9)			'품목규격 
			C_PoSts				=	iCurColumnPos(10)			'발주법인상태 
			C_PoStsNm			=	iCurColumnPos(11)			'발주법인상태 
			C_SoSts				=	iCurColumnPos(12)			'수주법인상태                
			C_SoStsNm			=	iCurColumnPos(13)			'수주법인상태                
			C_PoUnit			=	iCurColumnPos(14)			'단위                        
			C_PoQty				=	iCurColumnPos(15)			'발주수량                    
			C_SoQty				=	iCurColumnPos(16)			'수주수량                    
			C_PoLcQty			=	iCurColumnPos(17)			'수입L/C수량                 
			C_SoLcQty			=	iCurColumnPos(18)			'수출L/C수량                 
			C_SoReqQty			=	iCurColumnPos(19)			'출하요청수량                
			C_SoIssueQty		=	iCurColumnPos(20)			'출고수량                    
			C_SoCcQty			=	iCurColumnPos(21)			'수출통관수량                
			C_PoCcQty			=	iCurColumnPos(22)			'수입통관수량                
			C_PoRcptQty			=	iCurColumnPos(23)			'입고수량                    
			C_SoBillQty			=	iCurColumnPos(24)			'매출수량                    
			C_PoIvQty			=	iCurColumnPos(25)			'매입수량          
			     
	End Select
End Sub	

'------------------------------------------  OpenSoCompany()  -------------------------------------------------
' Name : OpenSoCompany()
' Description : SpreadItem PopUp
'---------------------------------------------------------------------------------------------------------- 
Function OpenSoCompany()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			 
	arrParam(0) = "수주법인" 
	arrParam(1) = "B_BIZ_PARTNER"    
	 
	arrParam(2) = Trim(frm1.txtSoCompanyCd.Value)

	arrParam(4) = "BP_TYPE In ('S','CS') And IN_OUT_FLAG = 'O'"
	arrParam(5) = "수주법인"

	arrField(0) = "BP_CD"
	arrField(1) = "BP_NM"

	arrHeader(0) = "수주법인"
	arrHeader(1) = "수주법인명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		frm1.txtSoCompanyCd.focus
		Exit Function
	Else
		frm1.txtSoCompanyCd.Value= arrRet(0)
		frm1.txtSoCompanyNm.Value= arrRet(1)
		frm1.txtSoCompanyCd.focus
	End If
End Function


'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'   Event Desc : 구매만 쓰임 그리드의 숫자 부분이 변경된면 이 함수를 변경 해야함.
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"Z"
    End Select
End Sub




'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc :
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If y<20 Then			'2003-03-01 Release 추가 
	    lgSpdHdrClicked = 1
	End If

    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
    End If
End Sub



'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)	'###그리드 컨버전 주의부분###

 	gMouseClickStatus = "SPC"

	Set gActiveSpdSheet = frm1.vspdData

 	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
	Else
		Call SetPopupMenuItemInf("0101111111")         '화면별 설정 
	End If

	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		lgSpdHdrClicked = 0		'2003-03-01 Release 추가 
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey1 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey1 = 2
 		ElseIf lgSortKey1 = 2 Then
 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
 			lgSortKey1 = 1
 		End If
	Else
 		'------ Developer Coding part (Start)

	 	'------ Developer Coding part (End)
 	End If

End Sub


'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : 그리드 해더 더블클릭시 네임 변경 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName

 	If Row <= 0 Then
		Exit Sub
 	End If

  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub



'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'=======================================================================================================
Sub Form_Load()	'###그리드 컨버전 주의부분###

	Call LoadInfTB19029                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
'	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")  'Lock  Suitable  Field
	Call InitSpreadSheet
	Call InitVariables
	Call SetDefaultVal



	set gActiveSpdSheet = frm1.vspdData
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : 그리드 위치 변경 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()	'###그리드 컨버전 주의부분###
	Dim iActiveRow
	Dim iConvActiveRow
	Dim lngRangeFrom
	Dim lngRangeTo
	Dim lRow
	Dim i
	Dim strFlag
	Dim strParentRowNo

    ggoSpread.Source = gActiveSpdSheet
    If gActiveSpdSheet.Name = "vspdData" Then
		Call ggoSpread.RestoreSpreadInf()
		Call InitSpreadSheet
		Call ggoSpread.ReOrderingSpreadData
    End If

 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)	'###그리드 컨버전 주의부분###
	If lgSpdHdrClicked = 1 Then	'2003-03-01 Release 추가 
		Exit Sub
	End If

	'/* 9월 정기패치 : 동일한 키값을 입력한 채 다른 스프레드로 옮기지 못하도록 수정관련 변수 추가 - START */
	Dim lRow
	'/* 9월 정기패치 : 동일한 키값을 입력한 채 다른 스프레드로 옮기지 못하도록 수정관련 변수 추가 - END */

	Set gActiveSpdSheet = frm1.vspdData

	frm1.vspdData.redraw = false
	If Row <> NewRow And NewRow > 0 Then
		With frm1
			.vspdData.redraw = false
			'/* 8월 정기패치 : 우측 스프레드에 필수입력 필드 체크 - START */
		'	If DefaultCheck = False Then
		'		.vspdData.Row = Row
		'		.vspdData.Col = 1
		'		.vspdData2.focus
    	'		Exit Sub
		'	End If
			'/* 8월 정기패치 : 우측 스프레드에 필수입력 필드 체크 - END */

			'/* 9월 정기패치: '다른 작업이 이루어지는 상황에서 다른 행 이동 시 조회가 이루어 지지 않도록 한다. - START */
			If CheckRunningBizProcess = True Then
				.vspdData.Row = Row
				.vspdData.Col = 1
				Exit Sub
			End If
			'/* 9월 정기패치: '다른 작업이 이루어지는 상황에서 다른 행 이동 시 조회가 이루어 지지 않도록 한다. - END */
			lgCurrRow = NewRow
			.vspdData.redraw = true
		End With

	End If
	frm1.vspdData.redraw = true
End Sub


'======================================================================================================
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'=======================================================================================================
'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'==========================================================================================
Sub txtPoFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoFrDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtPoFrDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : txtBillToDt
'   Event Desc :
'==========================================================================================
Sub txtPoToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoToDt.Action = 7
	    Call SetFocusToDocument("M")
        frm1.txtPoToDt.Focus
	End if
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc :
'==========================================================================================
Sub txtPoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPoToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() '###그리드 컨버전 주의부분###
    FncQuery = False

    Dim IntRetCD
    '-----------------------
    'Check previous data area
    '-----------------------

'    If ChangeCheck = True Then
'		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
'		If IntRetCD = vbNo Then
'			Exit Function
'		End If
'    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables															'Initializes local global variables

	ggoSpread.Source = frm1.vspdData	'###그리드 컨버전 주의부분###
    ggoSpread.ClearSpreadData

'	ggoSpread.Source = frm1.vspdData2
'    ggoSpread.ClearSpreadData

    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then											'This function check indispensable field
'	   Exit Function
'    End If

' === 2005.07.22 수정 ===========================================================

	If ValidDateCheck(frm1.txtPoFrDt, frm1.txtPoToDt) = False Then Exit Function

' === 2005.07.22 수정 ===========================================================

	'-----------------------
    'Query function call area
    '-----------------------
	If DbQuery = False then
		Exit Function
	End If
    															'☜: Query db data
	Set gActiveElement = document.activeElement
    FncQuery = True
End Function


'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave()
    FncSave = False

    Dim IntRetCD

	'-----------------------
    'Precheck area
    '-----------------------
    If ChangeCheck = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If

    '8월 정기패치: 화면에 보이는 우측 스프레드에 행추가 되었으나 Hidden 스프레드에 반영이 안된 것 체크 START
	If DefaultCheck = False Then
		Exit Function
	End If
    '8월 정기패치: 화면에 보이는 우측 스프레드에 행추가 되었으나 Hidden 스프레드에 반영이 안된 것 체크 END

'	intRetCd = DisplayMsgBox("900018", VB_YES_NO, "X", "X")   '☜ 바쾪E觀?
'	If intRetCd = VBNO Then
'		Exit Function
'	End IF


    '-----------------------
    'Check content area
    '-----------------------
'    If Not chkField(Document, "1") Then
'       		Exit Function
'    End If

    '-----------------------
    'Save function call area
    '-----------------------
	If DbSave = False then
		Exit Function
	End If

	Set gActiveElement = document.activeElement
    FncSave = True
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================
Function FncExcel()
	FncExcel = False
 	Call parent.FncExport(Parent.C_MULTI)
	Set gActiveElement = document.activeElement
 	FncExcel = True
 End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
	FncPrint = False
	Call Parent.FncPrint()
	Set gActiveElement = document.activeElement
	FncPrint = True
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================
Function FncFind()
	FncFind = False
    Call parent.FncFind(Parent.C_MULTI, False)                                         '☜:화면 유형, Tab 유무 
	Set gActiveElement = document.activeElement
    FncFind = True
End Function

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
	Set gActiveElement = document.activeElement
End Sub

'=======================================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================================
Function FncExit()
	FncExit = False

	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery()
	DbQuery = False

	Dim strVal

	Call LayerShowHide(1)
	with frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then

		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPoCompanyCd=" & Trim(.txtPoCompanyCd.value)
			strVal = strVal & "&txtSoCompanyCd=" & Trim(.txtSoCompanyCd.value)
			strVal = strVal & "&txtPoFrDt=" & Trim(.txtPoFrDt.Text)
			strVal = strVal & "&txtPoToDt=" & Trim(.txtPoToDt.Text)
			strVal = strVal & "&rdoImportFlg=" & Trim(.rdoImportFlg.Text)
		    strVal = strVal & "&lgPageNo=" & lgPageNo                  '☜: Next key tag
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		Else

		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPoCompanyCd=" & Trim(.txtPoCompanyCd.value)
			strVal = strVal & "&txtSoCompanyCd=" & Trim(.txtSoCompanyCd.value)
			strVal = strVal & "&txtPoFrDt=" & Trim(.txtPoFrDt.Text)
			strVal = strVal & "&txtPoToDt=" & Trim(.txtPoToDt.Text)
			if .rdoImportFlg(0).checked = true then
				strVal = strVal & "&rdoImportFlg=" & "%"
			elseif .rdoImportFlg(1).checked = true then
				strVal = strVal & "&rdoImportFlg=" & "N"
			else
				strVal = strVal & "&rdoImportFlg=" & "Y"
			End if
		    strVal = strVal & "&lgPageNo=" & lgPageNo                  '☜: Next key tag
		    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

		End If


	End with

	Call RunMyBizASP(MyBizASP, strVal)													'☜: 비지니스 ASP 를 가동 

	DbQuery = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbQueryOk(byVal intARow,byVal intTRow)
	DbQueryOk = False

	Dim i
	Dim lRow
	Dim TmpArrPrevKey
	Dim TmpArrHiddenRows

	Call ggoOper.LockField(Document, "Q")			'This function lock the suitable field
	Call SetToolBar("11000000000011")				'버튼 툴바 제어 

	With frm1
		'-----------------------
		'Reset variables area
		'-----------------------
		lRow = .vspdData.MaxRows

		i=0
		If lRow > 0 And intARow > 0 Then
			If intTRow<=0 Then
				ReDim lgStrPrevKeyM(intARow - 1)
				ReDim lglngHiddenRows(intARow - 1)			'lRow = .vspdData.MaxRows	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
			Else
				TmpArrPrevKey=lgStrPrevKeyM
				TmpArrHiddenRows=lglngHiddenRows

				ReDim lgStrPrevKeyM(intTRow+intARow - 1)
				ReDim lglngHiddenRows(intTRow+intARow - 1)			'lRow = .vspdData.MaxRows	'ex) 첫번째 그리드의 특정Row에 해당하는 두번째 그리드의 Row 갯수를 저장하는 배열.
				For i = 0 To intTRow-1
					lgStrPrevKeyM(i) = TmpArrPrevKey(i)
					lglngHiddenRows(i) = TmpArrHiddenRows(i)
				Next
			End If

			For i = intTRow To intTRow+intARow-1
				lglngHiddenRows(i) = 0
			Next

'		if lgIntFlgModeM = Parent.OPMD_CMODE then
'		    If DbQuery2(1, False) = False Then	Exit Function
'	    End If
		    lgIntFlgModeM = Parent.OPMD_UMODE
		    lgIntFlgMode = Parent.OPMD_UMODE
		End If
	End With

    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtSoCompanyCd.focus
	End If

	Set gActiveElement = document.activeElement
    DbQueryOk = true

    call InitVariables()




End Function

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_Check And Row > 0 Then
		ggoSpread.Source = frm1.vspdData
	    Select Case ButtonDown
	    Case 1
			ggoSpread.UpdateRow Row
	    Case 0
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row
			frm1.vspdData.text = ""
			lgBlnFlgChgValue = False
	    End Select
	ElseIf Col = C_CfmFlg And Row > 0 Then
		ggoSpread.Source = frm1.vspdData
	    Select Case ButtonDown
	    Case 1
			ggoSpread.UpdateRow Row
	    Case 0
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row
			frm1.vspdData.text = ""
			lgBlnFlgChgValue = False
	    End Select
	End If
	lgBlnFlgChgValue = True

End Sub

'--------------------------------------------------------------------
'		Cookie 사용함수 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD


	If Kubun = 0 Then

'		strTemp = ReadCookie("PoNo")

'		If strTemp = "" then Exit Function

'		frm1.txtPoNo.value = strTemp

'		WriteCookie "PoNo" , ""

'		Call MainQuery()

	elseIf Kubun = 1 Then

'	    If lgIntFlgMode <> Parent.OPMD_UMODE Then
'	        Call DisplayMsgBox("900002", "X", "X", "X")
'	        Exit Function
'	    End If

'	    If lgBlnFlgChgValue = True Then
'			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
'			If IntRetCD = vbNo Then
'				Exit Function
'			End If
'	    End If

'		WriteCookie "PoNo" , frm1.txtPoNo.value

		Call PgmJump(BIZ_PGM_JUMP_ID_PO_DTL)
	End IF
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>
<!--########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
'######################################################################################################## -->
<BODY TABINDEX="-1" SCROLL="no">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>멀티컴퍼니수발주진행조회</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
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
								<TD CLASS="TD5" NOWRAP>발주법인</TD>
								<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT ALT="발주법인"   NAME="txtPoCompanyCd" SIZE=10 MAXLENGTH=10 tag="24X" >
														<INPUT TYPE=TEXT ALT="발주법인" NAME="txtPoCompanyNm" SIZE=20 MAXLENGTH=50 tag="24X"></TD>
								<TD CLASS="TD5" NOWRAP>수주법인</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="수주법인"   NAME="txtSoCompanyCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenSoCompany" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSoCompany()" >
														<INPUT TYPE=TEXT ALT="수주법인" NAME="txtSoCompanyNm" SIZE=20 MAXLENGTH=50 tag="24X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발주일</TD>
								<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/kmm112qa1_fpDateTime2_txtPoFrDt.js'></script>~
										<script language =javascript src='./js/kmm112qa1_fpDateTime2_txtPoToDt.js'></script>
								</TD>
								<TD CLASS="TD5" NOWRAP>내외자처리구분</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="내외자처리구분" NAME="rdoImportFlg" id = "rdoImportFlg1" Value="A" checked tag="11"><label for="rdoImportFlg1">&nbsp;전체&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="내외자처리구분" NAME="rdoImportFlg" id = "rdoImportFlg2" Value="Y" tag="11"><label for="rdoImportFlg2">&nbsp;내자&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="내외자처리구분" NAME="rdoImportFlg" id = "rdoImportFlg3" Value="N" tag="11"><label for="rdoImportFlg3">&nbsp;외자&nbsp;</label></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<script language =javascript src='./js/kmm112qa1_OBJECT1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" align="left">&nbsp;</TD>
					<td WIDTH="*" align="right"></td>
					<TD WIDTH=10>&nbsp;</TD>
			 </TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>