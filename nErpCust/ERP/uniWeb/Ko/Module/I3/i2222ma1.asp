
<%@ LANGUAGE="VBSCRIPT" %> 
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2222ma1.asp
'*  4. Program Name         : Query Inventory Shortage
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2005/01/26
'*  9. Modifier (First)     : Chen, Jae Hyun
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'* 12. History              : 
'*                          :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################
'******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                      

On Error Resume Next
Err.Clear
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

Const BIZ_PGM_QRY1_ID						= "I2222mb1.asp"		'☆: Inventory Query ASP명 
Const BIZ_PGM_QRY2_ID						= "I2222mb2.asp"		'☆: ETC Query ASP명 

'-------------------------------
' Column Constants : Spread 1 
'-------------------------------
Dim C_ItemCd1			
Dim C_ItemNm1
Dim C_Spec1 		
Dim C_SLCD1		
Dim C_SLNm1		
Dim C_Tracking1				
Dim C_Block1
Dim C_Unit1	
Dim C_GoodQty1	
Dim C_BadQty1
Dim C_InspQty1
Dim C_TransitQty1
Dim C_SchedRcptQty1
Dim C_SchedIssueQty1
Dim C_PrevGoodQty1	
Dim C_PrevBadQty1
Dim C_PrevInspQty1
Dim C_PrevTransitQty1
Dim C_AllocQty1


'-------------------------------
' Column Constants : Spread 2 
'-------------------------------
Dim C_ProcMthd2
Dim C_OrderNo2
Dim C_OrderStatus2
Dim C_EndDt2
Dim C_OrderQty2
Dim C_Unit2
Dim C_ResultQty2
Dim C_SchedRecieptQty2
Dim C_RecieptQty2
Dim C_Manager2
Dim C_Mthd2

'-------------------------------
' Column Constants : Spread 3
'-------------------------------
Dim C_ProcMthd3
Dim C_OrderNo3
Dim C_OprNo3
Dim C_Seq3
Dim C_ReqDt3
Dim C_DestCd3
Dim C_DestNm3
Dim C_ReqQty3
Dim C_Unit3
Dim C_IssueQty3
Dim C_ConsumeQty3
Dim C_RemainQty3

'-------------------------------
' Column Constants : Spread 4 
'-------------------------------
Dim C_InspFlag4
Dim C_InspReqNo4
Dim C_InspStatus4
Dim C_InspReqDt4
Dim C_LotQty4
Dim C_LotUnit4
Dim C_GoodQty4
Dim C_DefectQty4

'-------------------------------
' Column Constants : Spread 5
'-------------------------------
Dim C_SLCd5
Dim C_SLNm5
Dim C_Tracking5
Dim C_LotNo5
Dim C_LotSubNo5		
Dim C_Block5		
Dim C_Unit5	
Dim C_GoodQty5
Dim C_BadQty5
Dim C_InspQty5
Dim C_TransitQty5
Dim C_PrevGoodQty5
Dim C_PrevBadQty5
Dim C_PrevInspQty5
Dim C_PrevTransitQty5

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrPrevKey1
Dim lgStrPrevKey2		
Dim lgStrPrevKey3

Dim lgSortKey1
Dim lgSortKey2
Dim lgSortKey3
Dim lgSortKey4
Dim lgSortKey5

Dim lgOldRow1

Dim iDBSYSDate
Dim LocSvrDate
Dim StartDate
Dim EndDate

	iDBSYSDate = "<%=GetSvrDate%>"			'⊙: DB의 현재 날짜를 받아와서 시작날짜에 사용한다.
	LocSvrDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
	StartDate = UNIDateAdd("D",-14,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 처음 날짜 
	EndDate = UNIDateAdd("D", 7,LocSvrDate, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 마지막 날짜 

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim IsOpenPop					 'Popup
Dim gSelframeFlg


'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================

'========================================================================================
' Function Name : InitVariables
' Function Desc : This method initializes general variables
'========================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgStrPrevKey3 = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	lgOldRow1 = 0
	lgSortKey1 = 1
	lgSortKey2 = 1
	lgSortKey3 = 1
	lgSortKey4 = 1
	lgSortKey5 = 1
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

    frm1.txtReqStartDt.Text = StartDate
	frm1.txtReqEndDt.Text = EndDate

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()   
   <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
   <% Call loadInfTB19029A("Q","I","NOCOOKIE","MA") %>
End Sub
    
'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
       
    Call InitSpreadPosVariables(pvSpdNo)
            
    With frm1
		If pvSpdNo = "A" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 1 Setting
			'-------------------------------------------
			
			ggoSpread.Source = .vspdData1
			ggoSpread.Spreadinit "V20041207", ,Parent.gAllowDragDropSpread
					
			.vspdData1.ReDraw = false
     		.vspdData1.MaxCols = C_AllocQty1 + 1
			.vspdData1.MaxRows = 0
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit		C_ItemCd1,		"품목", 18
			ggoSpread.SSSetEdit		C_ItemNm1,		"품목명", 20
			ggoSpread.SSSetEdit		C_Spec1,		"규격", 20
			ggoSpread.SSSetEdit		C_SLCD1,		"창고", 10
			ggoSpread.SSSetEdit		C_SLNm1,		"창고명", 10
			ggoSpread.SSSetEdit		C_Tracking1,	"Tracking No.", 20
			ggoSpread.SSSetEdit		C_Block1,		"Block", 8
			ggoSpread.SSSetEdit		C_Unit1,		"단위", 8
			ggoSpread.SSSetFloat	C_GoodQty1,		"양품재고량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_BadQty1,		"불량재고량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InspQty1,		"검사중수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_TransitQty1,	"이동중수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SchedRcptQty1,"입고예정량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SchedIssueQty1,"출고예정량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQty1,	"전월양품재고량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_PrevBadQty1,	"전월불량재고량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevInspQty1,	"전월검사중수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevTransitQty1,"전월이동중수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_AllocQty1,	"재고할당량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
			Call ggoSpread.SSSetColHidden(.vspdData1.MaxCols, .vspdData1.MaxCols, True)
			
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("A")
			
			.vspdData1.ReDraw = True
		End If
		
		If pvSpdNo = "B" Or pvSpdNo = "*" Then
			'-------------------------------------------
			' Spread 2 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData2
			ggoSpread.Spreadinit "V20041209", ,Parent.gAllowDragDropSpread
			.vspdData2.ReDraw = false
 
    		.vspdData2.MaxCols = C_Mthd2 + 1
			.vspdData2.MaxRows = 0

			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit		C_ProcMthd2,		"조달구분", 10
			ggoSpread.SSSetEdit		C_OrderNo2,			"오더번호", 18
			ggoSpread.SSSetEdit		C_OrderStatus2,		"오더상태", 10
			ggoSpread.SSSetDate 	C_EndDt2,			"납기일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_OrderQty2,		"오더수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_Unit2,			"단위", 7
			ggoSpread.SSSetFloat	C_ResultQty2,		"실적양품수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_SchedRecieptQty2,	"입고대기수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RecieptQty2,		"입고수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_Manager2	,		"담당자", 10
			ggoSpread.SSSetEdit		C_Mthd2,			"조달구분", 10

			Call ggoSpread.SSSetColHidden(C_Mthd2, C_Mthd2, True)
			Call ggoSpread.SSSetColHidden(.vspdData2.MaxCols, .vspdData2.MaxCols, True)

			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("B")
			
			.vspdData2.ReDraw = True
		End If
		
		If pvSpdNo = "C" Or pvSpdNo = "*" Then	
			'-------------------------------------------
			' Spread 3 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData3
			ggoSpread.Spreadinit "V20041206", ,Parent.gAllowDragDropSpread
			.vspdData3.ReDraw = false
			.vspdData3.MaxCols = C_RemainQty3 + 1
			.vspdData3.MaxRows = 0
	
			Call GetSpreadColumnPos("C")
		
			ggoSpread.SSSetEdit		C_ProcMthd3,	"모듈구분", 10
			ggoSpread.SSSetEdit		C_OrderNo3,		"오더번호", 18	
			ggoSpread.SSSetEdit		C_OprNo3,		"공정", 7
			ggoSpread.SSSetEdit		C_Seq3,			"순번", 7
			ggoSpread.SSSetDate 	C_ReqDt3,		"필요일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_DestCd3,		"현재위치", 10
			ggoSpread.SSSetEdit		C_DestNm3,		"현재위치명", 10
			ggoSpread.SSSetFloat	C_ReqQty3,		"필요수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_Unit3,		"단위", 7
			ggoSpread.SSSetFloat	C_IssueQty3,	"출고수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ConsumeQty3,	"소비수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RemainQty3,	"반입가능수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		
			Call ggoSpread.SSSetColHidden(.vspdData3.MaxCols, .vspdData3.MaxCols, True)
		
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("C")
			
			.vspdData3.ReDraw = True
		End If	
		
		
		If pvSpdNo = "D" Or pvSpdNo = "*" Then	
			'-------------------------------------------
			' Spread 4 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData4
			ggoSpread.Spreadinit "V20041206", ,Parent.gAllowDragDropSpread
			.vspdData4.ReDraw = false
			.vspdData4.MaxCols = C_DefectQty4 + 1
			.vspdData4.MaxRows = 0
	
			Call GetSpreadColumnPos("D")
		
			ggoSpread.SSSetEdit		C_InspFlag4,		"검사구분", 10
			ggoSpread.SSSetEdit		C_InspReqNo4,		"검사의뢰번호", 18	
			ggoSpread.SSSetEdit		C_InspStatus4,		"검사진행상태", 12
			ggoSpread.SSSetDate 	C_InspReqDt4,		"검사요구일", 11, 2, parent.gDateFormat
			ggoSpread.SSSetFloat	C_LotQty4,			"로트크기",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_LotUnit4,			"단위", 10
			ggoSpread.SSSetFloat	C_GoodQty4,			"검사양품수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_DefectQty4,		"검사불량수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
		
			Call ggoSpread.SSSetColHidden(.vspdData4.MaxCols, .vspdData4.MaxCols, True)
		
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("D")
			
			.vspdData4.ReDraw = True
		End If
		
		
		If pvSpdNo = "E" Or pvSpdNo = "*" Then	
			'-------------------------------------------
			' Spread 3 Setting
			'-------------------------------------------
			ggoSpread.Source = .vspdData5
			ggoSpread.Spreadinit "V20041106", ,Parent.gAllowDragDropSpread
			.vspdData5.ReDraw = false
			.vspdData5.MaxCols = C_PrevTransitQty5 + 1
			.vspdData5.MaxRows = 0
	
			Call GetSpreadColumnPos("E")
		
			ggoSpread.SSSetEdit		C_SLCD5,		"창고", 10
			ggoSpread.SSSetEdit		C_SLNm5,		"창고명", 10
			ggoSpread.SSSetEdit		C_Tracking5,	"Tracking No.", 20
			ggoSpread.SSSetEdit		C_LotNo5,		"Lot No.", 10
			ggoSpread.SSSetEdit		C_LotSubNo5,	"Lot Sub No.", 10
			ggoSpread.SSSetEdit		C_Block5,		"Block", 8
			ggoSpread.SSSetEdit		C_Unit5,		"단위", 8
			ggoSpread.SSSetFloat	C_GoodQty5,		"양품재고량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_BadQty5,		"불량재고량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InspQty5,		"검사중수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_TransitQty5,	"이동중수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevGoodQty5,	"전월양품재고량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat	C_PrevBadQty5,	"전월불량재고량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevInspQty5,	"전월검사중수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_PrevTransitQty5,"전월이동중수량",15,parent.ggQtyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			
			Call ggoSpread.SSSetColHidden(.vspdData5.MaxCols, .vspdData5.MaxCols, True)
		
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("E")
			
			.vspdData5.ReDraw = True
		End If
		
    End With
        
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
		If pvSpdNo = "A" Then
			'-------------------------
			' Set Lock Prop :Spread 1 		
			'-------------------------
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		
		If pvSpdNo = "B" Then		
			'-------------------------
			' Set Lock Prop :Spread 2 		
			'-------------------------
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If
		
		If pvSpdNo = "C" Then
			'-------------------------
			' Set Lock Prop :Spread 3 		
			'-------------------------
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If	
		
		If pvSpdNo = "D" Then
			'-------------------------
			' Set Lock Prop :Spread 4		
			'-------------------------
			ggoSpread.Source = frm1.vspdData4
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If	
		
		If pvSpdNo = "E" Then
			'-------------------------
			' Set Lock Prop :Spread 5		
			'-------------------------
			ggoSpread.Source = frm1.vspdData5
			ggoSpread.SpreadLockWithOddEvenRowColor()
		End If	

    End With
End Sub

'================================== 2.2.6 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : 
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	Dim strCboCd

	 '****************************
	'List Compare Sign(<,<=,=, =>, >)
	'****************************
	strCboCd =  "<" & Chr(11) & "<="  & Chr(11) & "="  & Chr(11) & ">="  & Chr(11) & ">"& Chr(11)
    Call SetCombo2(frm1.cboCompareSign, strCboCd, strCboCd, Chr(11))
	frm1.cboCompareSign.value = "<"
End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
    
    If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'grid1
		C_ItemCd1 = 1		
		C_ItemNm1 = 2
		C_Spec1 = 3 		
		C_SLCD1	= 4	
		C_SLNm1	= 5	
		C_Tracking1 = 6				
		C_Block1 = 7
		C_Unit1	= 8
		C_GoodQty1 = 9	
		C_BadQty1 = 10
		C_InspQty1 = 11
		C_TransitQty1 = 12
		C_SchedRcptQty1 = 13
		C_SchedIssueQty1 = 14
		C_PrevGoodQty1 = 15
		C_PrevBadQty1 = 16
		C_PrevInspQty1 = 17
		C_PrevTransitQty1 = 18
		C_AllocQty1 = 19 
	End If
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then	
		'grid2
		C_ProcMthd2 = 1
		C_OrderNo2 = 2
		C_OrderStatus2 = 3
		C_EndDt2 = 4
		C_OrderQty2 = 5
		C_Unit2 = 6
		C_ResultQty2 = 7
		C_SchedRecieptQty2 = 8
		C_RecieptQty2 = 9
		C_Manager2 = 10
		C_Mthd2 = 11
	End If
	
	If pvSpdNo = "C" Or pvSpdNo = "*" Then	
		'grid3
		C_ProcMthd3 = 1
		C_OrderNo3 = 2
		C_OprNo3 = 3
		C_Seq3 = 4
		C_ReqDt3 = 5
		C_DestCd3 = 6
		C_DestNm3 = 7
		C_ReqQty3 = 8
		C_Unit3 = 9
		C_IssueQty3 = 10
		C_ConsumeQty3 = 11
		C_RemainQty3 = 12
	End If	
	
	If pvSpdNo = "D" Or pvSpdNo = "*" Then	
		'grid4
		C_InspFlag4 = 1
		C_InspReqNo4 = 2
		C_InspStatus4 = 3
		C_InspReqDt4 = 4
		C_LotQty4 = 5
		C_LotUnit4 = 6
		C_GoodQty4 = 7
		C_DefectQty4 = 8

	End If	
	
	If pvSpdNo = "E" Or pvSpdNo = "*" Then	
		'grid5
		C_SLCd5 = 1
		C_SLNm5 = 2
		C_Tracking5 = 3
		C_LotNo5 = 4
		C_LotSubNo5	= 5	
		C_Block5 = 6		
		C_Unit5	 = 7
		C_GoodQty5 = 8
		C_BadQty5 = 9
		C_InspQty5 = 10
		C_TransitQty5 = 11
		C_PrevGoodQty5 = 12
		C_PrevBadQty5 = 13
		C_PrevInspQty5 = 14
		C_PrevTransitQty5 = 15

	End If	

End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==========
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'=================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		 	ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
					
			C_ItemCd1 = iCurColumnPos(1)
			C_ItemNm1 = iCurColumnPos(2)
			C_Spec1 = iCurColumnPos(3)	
			C_SLCD1	= iCurColumnPos(4)	
			C_SLNm1	= iCurColumnPos(5)	
			C_Tracking1 = iCurColumnPos(6)				
			C_Block1 = iCurColumnPos(7)
			C_Unit1	= iCurColumnPos(8)
			C_GoodQty1 = iCurColumnPos(9)	
			C_BadQty1 = iCurColumnPos(10)
			C_InspQty1 = iCurColumnPos(11)
			C_TransitQty1 = iCurColumnPos(12)
			C_SchedRcptQty1 = iCurColumnPos(13)
			C_SchedIssueQty1 = iCurColumnPos(14)
			C_PrevGoodQty1 = iCurColumnPos(15)
			C_PrevBadQty1 = iCurColumnPos(16)
			C_PrevInspQty1 = iCurColumnPos(17)
			C_PrevTransitQty1 = iCurColumnPos(18)
			C_AllocQty1 = iCurColumnPos(19)

		Case "B"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ProcMthd2 = iCurColumnPos(1)
			C_OrderNo2 = iCurColumnPos(2)
			C_OrderStatus2 = iCurColumnPos(3)
			C_EndDt2 = iCurColumnPos(4)
			C_OrderQty2 = iCurColumnPos(5)
			C_Unit2 = iCurColumnPos(6)
			C_ResultQty2 = iCurColumnPos(7)
			C_SchedRecieptQty2 = iCurColumnPos(8)
			C_RecieptQty2 = iCurColumnPos(9)
			C_Manager2 = iCurColumnPos(10)
			C_Mthd2 = iCurColumnPos(11)
					
		Case "C"
		 	ggoSpread.Source = frm1.vspdData3
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_ProcMthd3 = iCurColumnPos(1)
			C_OrderNo3 = iCurColumnPos(2)
			C_OprNo3 = iCurColumnPos(3)
			C_Seq3 = iCurColumnPos(4)
			C_ReqDt3 = iCurColumnPos(5)
			C_DestCd3 = iCurColumnPos(6)
			C_DestNm3 = iCurColumnPos(7)
			C_ReqQty3 = iCurColumnPos(8)
			C_Unit3 = iCurColumnPos(9)
			C_IssueQty3 = iCurColumnPos(10)
			C_ConsumeQty3 = iCurColumnPos(11)
			C_RemainQty3 = iCurColumnPos(12)
			
		Case "D"
			ggoSpread.Source = frm1.vspdData4
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_InspFlag4 = iCurColumnPos(1)
			C_InspReqNo4 = iCurColumnPos(2)
			C_InspStatus4 = iCurColumnPos(3)
			C_InspReqDt4 = iCurColumnPos(4)
			C_LotQty4 = iCurColumnPos(5)
			C_LotUnit4 = iCurColumnPos(6)
			C_GoodQty4 = iCurColumnPos(7)
			C_DefectQty4 = iCurColumnPos(8)
		
		
		Case "E"	
			ggoSpread.Source = frm1.vspdData5
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_SLCd5 = iCurColumnPos(1)
			C_SLNm5 = iCurColumnPos(2)
			C_Tracking5 = iCurColumnPos(3)
			C_LotNo5 = iCurColumnPos(4)
			C_LotSubNo5	= iCurColumnPos(5)	
			C_Block5 = iCurColumnPos(6)		
			C_Unit5	 = iCurColumnPos(7)
			C_GoodQty5 = iCurColumnPos(8)
			C_BadQty5 = iCurColumnPos(9)
			C_InspQty5 = iCurColumnPos(10)
			C_TransitQty5 = iCurColumnPos(11)
			C_PrevGoodQty5 = iCurColumnPos(12)
			C_PrevBadQty5 = iCurColumnPos(13)
			C_PrevInspQty5 = iCurColumnPos(14)
			C_PrevTransitQty5 = iCurColumnPos(15)
					
    End Select    
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenCondPlant()  --------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetConPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call displaymsgbox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분: From To를 입력할 것 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenProdOrderNo()  -------------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If frm1.txtPlantCd.value = "" Then
		Call displaymsgbox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("P4111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdFromDt.Text
	arrParam(2) = frm1.txtProdToDt.Text
	arrParam(3) = "RL"
	arrParam(4) = "RL"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value) 
	arrParam(6) = Trim(frm1.txtTrackingNo.value)
	arrParam(7) = Trim(frm1.txtItemCd.value)
	arrParam(8) = Trim(frm1.cboOrderType.value)
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
		
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTrackingInfo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
    
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName

	If IsOpenPop = True  Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = ""
	arrParam(3) = ""
	arrParam(4) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
		
End Function

'------------------------------------------  OpenItemGroup()  --------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "품목그룹"
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
    arrField(3) = "LEAF_FLG"	
    arrField(4) = "UPPER_ITEM_GROUP_CD"	
    
    arrHeader(0) = "품목그룹"		
    arrHeader(1) = "품목그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
	
End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
'	Name : OpenSLCd()	입고창고 
'	Description : Storage Location PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSLCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtSLCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X" , "공장", "X")
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "창고팝업"											' 팝업 명칭 
	arrParam(1) = "B_STORAGE_LOCATION"										' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtSLCd.Value)									' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")  	' Where Condition
	arrParam(5) = "창고"												' TextBox 명칭 
	
    arrField(0) = "SL_CD"													' Field명(0)
    arrField(1) = "SL_NM"													' Field명(1)
    
    arrHeader(0) = "창고"												' Header명(0)
    arrHeader(1) = "창고명"												' Header명(1)
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSLCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtSLCd.focus
	
End Function

'------------------------------------------  OpenUnit()  ----------------------------------------------
'	Name : OpenMfgUnit()	오더단위 
'	Description : Unit PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenUnit()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtUnit.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "단위팝업"	
	arrParam(1) = "B_UNIT_OF_MEASURE"				
	arrParam(2) = Trim(frm1.txtUnit.Value)
	arrParam(3) = ""
	arrParam(4) = "Dimension <> " & FilterVar("TM", "''", "S") & "  "			
	arrParam(5) = "단위"
	
    arrField(0) = "UNIT"	
    arrField(1) = "UNIT_NM"	
    
    arrHeader(0) = "단위"		
    arrHeader(1) = "단위명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetUnit(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtUnit.focus
	
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)
	With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
	End With
End Function

'------------------------------------------  SetConPlant()  ----------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
End Function    

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
    frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetUnit()  --------------------------------------------------
'	Name : SetUnit()
'	Description : MfgUnit Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetUnit(byval arrRet)
	frm1.txtUnit.Value    = arrRet(0)		
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetSLCd()  --------------------------------------------------
'	Name : SetSLCd()
'	Description : Ware House Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSLCd(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)
	frm1.txtSlNm.value	  = arrRet(1)	
	lgBlnFlgChgValue = True
End Function

'------------------------------------------  SetItemGroup()  ---------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value	= arrRet(0)		
	frm1.txtItemGroupNm.value   = arrRet(1)
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
	Call InitSpreadSheet("*")															'⊙: Setup the Spread sheet
	Call InitComboBox
	Call SetDefaultVal
	Call InitVariables																'⊙: Initializes local global variables
	
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("11000000000011")
		
	If frm1.txtPlantCd.value = "" Then
		If parent.gPlant <> "" Then
			frm1.txtPlantCd.value = parent.gPlant
			frm1.txtPlantNm.value = parent.gPlantNm
			frm1.txtReqStartDt.focus 
			Set gActiveElement = document.activeElement 
		Else
			frm1.txtPlantCd.focus 
			Set gActiveElement = document.activeElement 
		End If
	End If
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'*********************************************************************************************************

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtReqStartDt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtReqStartDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtReqEndDt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtReqEndDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtBaseQty_KeyDown ------------------------------------------
'	Name : txtBaseQty_KeyDown
'	Description : 
'----------------------------------------------------------------------------------------------------------
Sub txtBaseQty_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================

Sub txtReqStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqStartDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqStartDt.Focus
    End If
End Sub

Sub txtReqEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtReqEndDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtReqEndDt.Focus
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1
	
	If frm1.vspdData1.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey1
            lgSortKey1 = 1
        End If
		
		frm1.vspdData2.MaxRows = 0
		frm1.vspdData3.MaxRows = 0
		frm1.vspdData4.MaxRows = 0
		frm1.vspdData5.MaxRows = 0
		
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_Tracking1
		frm1.txtTrackingNo2.value = frm1.vspdData1.Text
		frm1.vspdData1.Col = C_SLCD1
		frm1.txtSLCd2.value = frm1.vspdData1.Text
		frm1.vspdData1.Col = C_SLNm1
		frm1.txtSlNm2.value = frm1.vspdData1.Text
			
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
        
	Else
		If lgOldRow1 <> Row Then
			
			frm1.vspdData2.MaxRows = 0
			frm1.vspdData3.MaxRows = 0
			frm1.vspdData4.MaxRows = 0
			frm1.vspdData5.MaxRows = 0
			
			frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
			frm1.vspdData1.Col = C_Tracking1
			frm1.txtTrackingNo2.value = frm1.vspdData1.Text
			frm1.vspdData1.Col = C_SLCD1
			frm1.txtSLCd2.value = frm1.vspdData1.Text
			frm1.vspdData1.Col = C_SLNm1
			frm1.txtSlNm2.value = frm1.vspdData1.Text
			
			If DbQuery2 = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
					
			lgOldRow1 = Row
		End If
	End If	
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================

Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2
	
	If frm1.vspdData2.MaxRows = 0 Then
 		Exit Sub
 	End If
	
	If Row <= 0 Then
	
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey2 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey2 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey2
            lgSortKey2 = 1
        End If   
		
	End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData3_Click
'   Event Desc :
'==========================================================================================
Sub vspdData3_Click(ByVal Col , ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP3C"
	
	Set gActiveSpdSheet = frm1.vspdData3
	
	If frm1.vspdData3.MaxRows = 0 Then
 		Exit Sub
 	End If
	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData3
        If lgSortKey3 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey3 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey3
            lgSortKey3 = 1
        End If
        Exit Sub
    End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData4_Click
'   Event Desc :
'==========================================================================================
Sub vspdData4_Click(ByVal Col , ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP4C"
	
	Set gActiveSpdSheet = frm1.vspdData4
	
	If frm1.vspdData4.MaxRows = 0 Then
 		Exit Sub
 	End If
	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData4
        If lgSortKey4 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey4 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey4
            lgSortKey4 = 1
        End If
        Exit Sub
    End If
	
End Sub


'==========================================================================================
'   Event Name : vspdData4_Click
'   Event Desc :
'==========================================================================================
Sub vspdData5_Click(ByVal Col , ByVal Row)
	
	Call SetPopupMenuItemInf("0000111111")
	
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP5C"
	
	Set gActiveSpdSheet = frm1.vspdData5
	
	If frm1.vspdData5.MaxRows = 0 Then
 		Exit Sub
 	End If
	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData5
        If lgSortKey5 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey5 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey5
            lgSortKey5 = 1
        End If
        Exit Sub
    End If
	
End Sub


'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData3_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData3_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP3C" Then
       gMouseClickStatus = "SP3CR"
    End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData4_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData4_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP4C" Then
       gMouseClickStatus = "SP4CR"
    End If
	
End Sub

'==========================================================================================
'   Event Name : vspdData5_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData5_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP5C" Then
       gMouseClickStatus = "SP5CR"
    End If
	
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData5_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData5
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

Sub vspdData3_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("C")
End Sub

Sub vspdData4_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData4
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("D")
End Sub

Sub vspdData5_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData5
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("E")
End Sub



'==========================================================================================
'   Event Name : vspdData1_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData1
	
    If Row >= NewRow Then
        Exit Sub
    End If

	'----------  Coding part  -------------------------------------------------------------

    End With

End Sub


'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
       Exit Sub
	End If  
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '----------  Coding part  -------------------------------------------------------------
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1 ,NewTop) Then
		If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" And lgStrPrevKey3 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If		
		End If
    End if
    
End Sub


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
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If
		
	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If ValidDateCheck(frm1.txtReqStartDt, frm1.txtReqEndDt) = False Then Exit Function	
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    Call InitVariables															'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
   
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
   
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
     On Error Resume Next
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next    
End Function

'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	If frm1.vspdData1.MaxRows <= 0 Then Exit Function	

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.EditUndo            
	
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    On Error Resume Next
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
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '☜: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)									'☜: 화면 유형 
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

End Sub

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI, False)                               '☜:화면 유형, Tab 유무 
End Function

'========================================================================================
' Function Name : FncScreenSave
' Function Desc : This function is related to FncScreenSave menu item of Main menu
'========================================================================================
Function FncScreenSave() 
    Call ggoSpread.SaveLayout
End Function

'========================================================================================
' Function Name : FncScreenRestore
' Function Desc : This function is related to FncScreenRestore menu item of Main menu
'========================================================================================
Function FncScreenRestore() 
    If ggoSpread.AllClear = True Then
       ggoSpread.LoadLayout
    End If
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 

End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete가 성공적일때 수행 
'========================================================================================
Function DbDeleteOk()												'☆: 삭제 성공후 실행 로직 

End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : Spread 1 조회 및 Scroll
'========================================================================================
Function DbQuery() 
    
    DbQuery = False
    
    Call LayerShowHide(1)
 
    Dim strVal
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCd.value)
		strVal = strVal & "&txtSLCd=" & Trim(frm1.hSLCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCd.value)
		strVal = strVal & "&txtBaseQty=" & Trim(frm1.hBaseQty.value)
		strVal = strVal & "&txtSign=" & Trim(frm1.hcboCompareFlag.value)
		strVal = strVal & "&txtBaseUnit=" & Trim(frm1.hBaseUnit.value)
		strVal = strVal & "&rdoIssueFlag=" & Trim(frm1.hSchedIssueFlg.value)
		strVal = strVal & "&rdoInventoryFlag=" & Trim(frm1.hInventoryFlg.value)
	Else
		strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgStrPrevKey3=" & lgStrPrevKey3
		strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCd.value)
		strVal = strVal & "&txtSLCd=" & Trim(frm1.txtSlCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(frm1.txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCd.value)
		strVal = strVal & "&txtBaseQty=" & Trim(frm1.txtBaseQty.text)
		strVal = strVal & "&txtSign=" & Trim(frm1.cboCompareSign.value)
		strVal = strVal & "&txtBaseUnit=" & Trim(frm1.txtUnit.value)
		Select Case frm1.rdoIssueFlg1.checked
			Case True
				strVal = strVal & "&rdoIssueFlag=" & frm1.rdoIssueFlg1.value 
			Case False
				strVal = strVal & "&rdoIssueFlag=" & frm1.rdoIssueFlg2.value 
		End Select	
		
		If frm1.rdoInventoryFlg1.checked Then
			strVal = strVal & "&rdoInventoryFlag=" & frm1.rdoInventoryFlg1.value
		ElseIf frm1.rdoInventoryFlg2.checked Then
			strVal = strVal & "&rdoInventoryFlag=" & frm1.rdoInventoryFlg2.value
		Else
			strVal = strVal & "&rdoInventoryFlag=" & frm1.rdoInventoryFlg3.value
		End If
		
	End If    

    Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()
    
    ggoSpread.Source = frm1.vspdData1

	frm1.vspdData1.ReDraw = False
	
	frm1.vspdData1.ReDraw = True
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then
		
		lgOldRow1 = 1
		
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
		
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_Tracking1
		frm1.txtTrackingNo2.value = frm1.vspdData1.Text
		frm1.vspdData1.Col = C_SLCD1
		frm1.txtSLCd2.value = frm1.vspdData1.Text
		frm1.vspdData1.Col = C_SLNm1
		frm1.txtSlNm2.value = frm1.vspdData1.Text
		
		If DbQuery2 = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
		
    End If

    lgIntFlgMode = parent.OPMD_UMODE													'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")										'⊙: This function lock the suitable field
    
    Call SetToolBar("11000000000111")
    
    lgBlnFlgChgValue = False
    
End Function

'========================================================================================
' Function Name : DbQuery2
' Function Desc : Spread 2 
'========================================================================================

Function DbQuery2() 
	
	Dim strVal
	Dim strItemCd
	Dim strSLCd
	Dim strTrackingNo
	
	With frm1.vspdData1
		.Row = .ActiveRow
		
		.Col = C_ItemCd1
		strItemCd = .Value
		
		.Col = C_SLCD1
		strSLCd = .Value
		 
		.Col = C_Tracking1
		strTrackingNo = .Value
	
	End With
    
    DbQuery2 = False                                    
    
    Call LayerShowHide(1)
	
	strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001	
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCd.value)			'☜: 조회 조건 데이타 
	strVal = strVal & "&txtReqStartDt=" & Trim(frm1.txtReqStartDt.text)
	strVal = strVal & "&txtReqEndDt=" & Trim(frm1.txtReqEndDt.text)
	strVal = strVal & "&txtItemCd=" & Trim(strItemCd)			'☜: 조회 조건 데이타 
	strVal = strVal & "&txtTrackingNo=" & Trim(strTrackingNo)		'☜: 조회 조건 데이타 
	strVal = strVal & "&txtSLCd=" & Trim(strSLCd)					'☜: 조회 조건 데이타 
	
    Call RunMyBizASP(MyBizASP, strVal)											

    DbQuery2 = True                                                          	

End Function

'========================================================================================
' Function Name : DbQuery2Ok
' Function Desc : Spread 2 And Spread 3 Data 조회 
'========================================================================================

Function DbQuery2Ok() 
	Dim LngRow
	
    With frm1.vspdData2
		.ReDraw = False
		If .MaxRows > 0 Then
			For LngRow = 1 To .MaxRows
				.Row = LngRow
				.Col = C_Mthd2
				Select Case Trim(.Text)
					Case "M"
						.Col = C_SchedRecieptQty2
						If uniCDbl(.Text) > 0 Then
							.ForeColor = vbRed
							.Col = C_OrderNo2
							.ForeColor = vbRed
						End If
					
					Case "O", "P"
						.Col = C_SchedRecieptQty2
						If uniCDbl(.Text) > 0 Then
						.Col = C_EndDt2
							If CompareDateByFormat(.Text, LocSvrDate,"납기일","현재일","970025",parent.gDateFormat,parent.gComDateType,False) = True Then
								.Col = C_SchedRecieptQty2
								.ForeColor = vbRed
								.Col = C_OrderNo2
								.ForeColor = vbRed
							End If 
						End If					
				End Select	
			Next
		End If
		.ReDraw = True
	End With
    
    With frm1.vspdData3
		.ReDraw = False
		If .MaxRows > 0 Then
			For LngRow = 1 To .MaxRows
				.Row = LngRow
				.Col = C_RemainQty3
				If uniCDbl(.Text) > 0 Then
					.ForeColor = vbRed
					.Col = C_OrderNo3
					.ForeColor = vbRed
				End If
			Next
		End If
		.ReDraw = True
	End With
        
    frm1.vspdData1.Focus
   
	
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 

End Function    

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
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData()
End Sub 

'========================================================================================
' Function Name : ViewHidden
' Function Desc : Show Detail Field
'========================================================================================
Function ViewHidden(StrMnuID, MnuCount, StrImageSize )
    Dim ii

    For ii = 1 To MnuCount
        If document.all(StrMnuID & ii).style.display = "" Then 
           document.all(StrMnuID & ii).style.display = "none"
           Select Case StrImageSize
				Case 1
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/Smallplus.gif"
				Case 2
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddlePlus.gif"
				Case 3
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/BigPlus.gif"
				Case Else
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddlePlus.gif"
			End Select		
        Else
           document.all(StrMnuID & ii).style.display = ""
           Select Case StrImageSize
				Case 1
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/SmallMinus.gif"
				Case 2
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddleMinus.gif"
				Case 3
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/BigMinus.gif"
				Case Else
					document.all("IMG_" & StrMnuID).src= "../../../CShared/image/MiddleMinus.gif"
			End Select
        End If
    Next    

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고관련오더조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="공장"></TD>
									<TD CLASS=TD5 NOWRAP>기준일</TD>
								    <TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/i2222ma1_fpDateTime1_txtReqStartDt.js'></script>
									&nbsp;~&nbsp;
									<script language =javascript src='./js/i2222ma1_fpDateTime1_txtReqEndDt.js'></script></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>재고수량</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>													
											<TR>
												<TD><SELECT NAME="cboCompareSign" ALT="재고수량" STYLE="Width: 40px;" tag="12"><OPTION VALUE=""></OPTION></SELECT>
												</TD>
												<TD>																
													<script language =javascript src='./js/i2222ma1_OBJECT1_txtBaseQty.js'></script>				
												</TD>
												<TD valign=bottom>
													&nbsp;<INPUT TYPE=TEXT NAME="txtUnit" SIZE=5 MAXLENGTH=7 tag="11xxxU" ALT="재고단위"  ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUnit1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit()">
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>창고</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING= 0>
											<TR>
												<TD>
													<INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSlCd" SIZE=10 MAXLENGTH= 7 tag="11xxxU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSlCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSlCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" SIZE=25 tag="14" ALT="창고">
												</TD>
												<TD WIDTH="*">
													&nbsp;
												</TD>
												<TD  WIDTH="20" STYLE="TEXT-ALIGN: RIGHT" ><IMG SRC="../../../CShared/image/BigPlus.gif" Style="CURSOR: hand" ALT="DetailCondition" ALIGN= "TOP" ID = "IMG_DetailCondition" NAME="pop1" ONCLICK= 'vbscript:viewHidden "DetailCondition" ,3, 3' ></IMG></TD>
											</TR>
										</TABLE>	
									</TD>
								</TR>
								<TR ID="DetailCondition1" style="display: none">
									<TD CLASS=TD5 NOWRAP>기준재고량</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInventoryFlg" tag="11" CHECKED ID="rdoInventoryFlg3" VALUE="A"><LABEL FOR="rdoInventoryFlg">모두</LABEL>&nbsp;
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInventoryFlg" tag="11" ID="rdoInventoryFlg1" VALUE="Y"><LABEL FOR="rdoInventoryFlg">현재고</LABEL>&nbsp;
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoInventoryFlg" tag="11" ID="rdoInventoryFlg2" VALUE="N"><LABEL FOR="rdoInventoryFlg">전월재고</LABEL>
														</TD>
									<TD CLASS=TD5 NOWRAP>출고예정여부</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSchedIssueFlg" tag="11" CHECKED ID="rdoIssueFlg1" VALUE="Y"><LABEL FOR="rdoSchedIssueFlg">예</LABEL>
														<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSchedIssueFlg" tag="11" ID="rdoIssueFlg2" VALUE="N"><LABEL FOR="rdoSchedIssueFlg">아니오</LABEL></TD>
								</TR>
								<TR ID="DetailCondition2" style="display: none">
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=25 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
								</TR>
								<TR ID="DetailCondition3" style="display: none">
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>						
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>창고</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtSLCd2" SIZE=10 tag="24" ALT="창고">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm2" SIZE=25 tag="24" ALT="창고"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
								    <TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtTrackingNo2" SIZE=25 MAXLENGTH=25 tag="24" ALT="Tracking No."></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD WIDTH="30%">
									<script language =javascript src='./js/i2222ma1_A_vspdData1.js'></script>
								</TD>
								<TD WIDTH="70%">
									
									<TABLE <%=LR_SPACE_TYPE_20%> BORDER=1>
										<TR>
											<TD CLASS=TD656 <%=HEIGHT_TYPE_02%>>
												<IMG SRC="../../../CShared/image/SmallMinus.gif" Style="CURSOR: hand" align=top ALT="검사대기" ID = "IMG_InspQueue" NAME="pop4" ONCLICK= 'vbscript:viewHidden "InspQueue" ,1, 1' >
												검사대기
											</TD>
										</TR>
										<TR ID="InspQueue1" style="">	
											<TD>
												<script language =javascript src='./js/i2222ma1_D_vspdData4.js'></script>
											</TD>	
										</TR>
										<TR>
											<TD CLASS=TD656 <%=HEIGHT_TYPE_02%>>
												<IMG SRC="../../../CShared/image/SmallMinus.gif" Style="CURSOR: hand" align=top ALT="입고예정오더" ID = "IMG_RecieptDetail" NAME="pop2" ONCLICK= 'vbscript:viewHidden "RecieptDetail" ,1, 1'></IMG>
												입고현황
											</TD>
										</TR>	
										<TR ID="RecieptDetail1" style="">
											<TD>
												<script language =javascript src='./js/i2222ma1_B_vspdData2.js'></script>
											</TD>
										</TR>
										<TR>
											<TD CLASS=TD656 <%=HEIGHT_TYPE_02%>>
												<IMG SRC="../../../CShared/image/SmallMinus.gif" Style="CURSOR: hand" align=top ALT="반입가능오더" ID = "IMG_IssueDetail" NAME="pop3" ONCLICK= 'vbscript:viewHidden "IssueDetail" ,1, 1' >
												출고현황
											</TD>
										</TR>
										<TR ID="IssueDetail1" style="">	
											<TD>
												<script language =javascript src='./js/i2222ma1_C_vspdData3.js'></script>
											</TD>	
										</TR>
										
										<TR>
											<TD CLASS=TD656 <%=HEIGHT_TYPE_02%>>
												<IMG SRC="../../../CShared/image/SmallMinus.gif" Style="CURSOR: hand" align=top ALT="이동가능재고" ID = "IMG_TransitInv" NAME="pop5" ONCLICK= 'vbscript:viewHidden "TransitInv" ,1, 1' >
												이동가능재고
											</TD>
										</TR>
										<TR ID="TransitInv1" style="">	
											<TD>
												<script language =javascript src='./js/i2222ma1_E_vspdData5.js'></script>
											</TD>	
										</TR>
									</TABLE>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TABINDEX = "-1" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" TABINDEX = "-1" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hProdOrderNo" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hBaseQty" TABINDEX = "-1" tag="24">
<INPUT TYPE=HIDDEN NAME="hBaseUnit" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hItemGroupCd" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hSLCd" TABINDEX = "-1" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hReqStartDt" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hReqEndDt" TABINDEX = "-1" tag="24">
<INPUT TYPE=HIDDEN NAME="hReqEndDt" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hSchedIssueFlg" TABINDEX = "-1" tag="24"><INPUT TYPE=HIDDEN NAME="hInventoryFlg" TABINDEX = "-1" tag="24">
<INPUT TYPE=HIDDEN NAME="hcboCompareFlag" TABINDEX = "-1" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
