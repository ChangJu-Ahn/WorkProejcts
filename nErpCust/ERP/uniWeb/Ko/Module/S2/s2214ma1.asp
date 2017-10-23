<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2214MA1
'*  4. Program Name         : 고객별품목판매계획등록 
'*  5. Program Desc         : 고객별품목판매계획등록 
'*  6. Comproxy List        : PS2G241.dll
'*  7. Modified date(First) : 2002/12/31
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Seongbae Hwang
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= 
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.Inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                             '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "s2214mb1.asp"            '☆: Head Query 비지니스 로직 ASP명 
Const BIZ_JUMP_ID = "s2214ba1"				 '☆: JUMP시 비지니스 로직 ASP명 

Const C_PopSalesGrp		= 1
Const C_PopFrSpPeriod	= 2
Const C_PopToSpPeriod	= 3
Const C_PopSoldToParty	= 4
Const C_PopItemCd		= 5

<!-- #Include file="../../inc/lgvariables.inc" --> 

'========================================================================================================
'=                       Variables For spreadsheet
'========================================================================================================
'☆: Spread Sheet의 Column
'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
Dim C_SpType
Dim C_LocExpFlag
Dim C_SoBillFlag
Dim C_SpPeriod
Dim C_SpPeriodPopup
Dim C_SpPeriodDesc
Dim C_SpSeq
Dim C_SalesGrp
Dim C_SoldToParty
Dim C_SoldToPartyPopup
Dim C_SoldToPartyNm
Dim C_Cur
Dim C_CurPopup
Dim C_XchgRate
Dim C_ItemCd
Dim C_ItemCdPopup
Dim C_ItemNm
Dim C_Spec
Dim C_Qty
Dim C_Unit
Dim C_UnitPopup
Dim C_Price
Dim C_Amt
Dim C_AmtLoc
Dim C_CfmFlag
Dim C_DistrFlag
Dim C_FromDt
Dim C_ToDt
Dim C_SpMonth
Dim C_SpWeek
Dim C_XchgRateOp
Dim C_Pointer
Dim C_OldSpPeriod
Dim C_OldSoldToParty
Dim C_OldQty
Dim C_OldAmt

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim C_SpPeriod2
Dim C_SpPeriodDesc2
Dim C_SoldToParty2
Dim C_SoldToPartyNm2
Dim C_TotQty
Dim C_Cur2
Dim C_TotAmt

'========================================================================================================
'=                       User-defind Variables
'========================================================================================================
Dim lgBlnOpenPop
Dim lgStrWhere					' Scrollbar를 조회조건 
Dim	lgStrPriceRule				' 단가 적용 규칙 
Dim	lgXchgRateFg				' 환율 적용기준 
Dim	lgXPmNonXchgRate			' 환율 처리방법 
Dim	lgStrLastCfmSpPeriod		' 최종확정기간 
Dim	lgStrLastCfmSpPeriodDesc
Dim	lgDtLastCfmToDt
Dim lgBlnExistsSpConfig			' 판매계획환경설정 존재여부 
Dim lgBlnDisplayMsg				' 환율이 없는 경우 경고 메세지 Display 여부 
Dim lgBlnOpenedFlag				' 화면 Load 여부 

Dim lgLngStartRow		' Start row to be queryed

'========================================================================================================
Sub initSpreadPosVariables()  
	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SPCRP" Then
		C_SpType			= 1
		C_LocExpFlag		= 2
		C_SoBillFlag		= 3
		C_SpPeriod			= 4
		C_SpPeriodPopup		= 5
		C_SpPeriodDesc		= 6
		C_SpSeq				= 7
		C_SalesGrp			= 8
		C_SoldToParty		= 9
		C_SoldToPartyPopup	= 10
		C_SoldToPartyNm		= 11
		C_Cur				= 12
		C_CurPopup			= 13
		C_XchgRate			= 14
		C_ItemCd			= 15
		C_ItemCdPopup		= 16
		C_ItemNm			= 17
		C_Spec				= 18
		C_Qty				= 19
		C_Unit				= 20
		C_UnitPopup			= 21
		C_Price				= 22
		C_Amt				= 23
		C_AmtLoc			= 24
		C_CfmFlag			= 25
		C_DistrFlag			= 26
		C_FromDt			= 27
		C_ToDt				= 28
		C_SpMonth			= 29
		C_SpWeek			= 30
		C_XchgRateOp		= 31
		C_Pointer			= 32
		C_OldSpPeriod		= 33
		C_OldSoldToParty	= 34
		C_OldQty			= 35
		C_OldAmt			= 36
	End If
	
	'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SP2CRP" Then
		C_SpPeriod2			= 1
		C_SpPeriodDesc2		= 2
		C_SoldToParty2		= 3
		C_SoldToPartyNm2	= 4
		C_TotQty			= 5
		C_Cur2				= 6
		C_TotAmt			= 7
	End If
	
End Sub

'========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  
    lgBlnOpenPop = False
    lgStrLastCfmSpPeriod = ""
    lgStrLastCfmSpPeriodDesc = ""
    lgDtLastCfmToDt = ""
    lgBlnDisplayMsg = True    
End Sub

'=========================================================================================================
Sub SetDefaultVal()
	'영업그룹 Default값처리 
	If Parent.gSalesGrp <> "" And Trim(frm1.txtConSalesGrp.value) = "" Then
		frm1.txtConSalesGrp.value = parent.gSalesGrp
		Call GetCfmPeriod(0)
	End If

'	Call GetSpConfig()

	frm1.cboConSpType.focus
End Sub

'=========================================================================================================
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetRowDefaultVal(ByVal pvRowCnt)
	Dim iIntRow
	
	With frm1.vspdData
		For iIntRow = 0 To pvRowCnt - 1
			.Row = .ActiveRow + iIntRow
	
			.Col = C_SpType		:	.Text = frm1.cboSpType.value
			.Col = C_LocExpFlag	:	.Text = frm1.cboLocExpFlag.value
			.Col = C_SoBillFlag :	.Text = "2"
			.Col = C_SpSeq		:	.Text = "0"
			.Col = C_SalesGrp	:	.Text = frm1.txtSalesGrp.value
			.Col = C_Qty		:	.Text = "0"
			.Col = C_Price		:	.Text = "0"
			.Col = C_Amt		:	.Text = "0"
			.Col = C_CfmFlag	:	.Text = "N"
			.Col = C_DistrFlag	:	.Text = "N"
		Next
	End With

End Sub

' Copy row
Sub SetRowCopyDefaultVal(ByVal pvRowCnt)

	With frm1.vspdData
	
		.Row = pvRowCnt
	
		.Col = C_ItemCd		:	.Text = ""
		.Col = C_ItemNm		:	.Text = ""
		.Col = C_Spec		:	.Text = ""
		.Col = C_Qty		:	.Text = "0"
		.Col = C_Unit		:	.Text = ""
		.Col = C_Price		:	.Text = "0"
		.Col = C_Amt		:	.Text = "0"
		.Col = C_AmtLoc		:	.Text = "0"
		.Col = C_CfmFlag	:	.Text = "N"
		.Col = C_DistrFlag	:	.Text = "N"
		.Col = C_OldQty		:	.Text = "0"
		.Col = C_OldAmt		:	.Text = "0"

		.Row = pvRowCnt - 1
		.Col = C_CfmFlag
		If .Text = "N" Then
			.Row = pvRowCnt
			.Col = C_ItemCd
		Else
			.Row = pvRowCnt
			.Col = C_FromDt			:	.Text = ""
			.Col = C_SpPeriodDesc	:	.Text = ""
			.Col = C_SpPeriod		:	.Text = ""

			If .ColHidden Then	.Col = C_FromDt
		End If

		.Action = 0
	End With

End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>
End Sub
'==========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
   	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
   	' gMouseClickStatus = "N" : when the form is loaded
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SPCRP" Then
		With frm1.vspdData		
			
		   	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
			ggoSpread.Source = frm1.vspdData
			'patch version
		    ggoSpread.Spreadinit "V20051208",,parent.gAllowDragDropSpread    		
			.ReDraw = false
			
			.MaxRows = 0 : .MaxCols = 0
			.MaxCols = C_OldAmt + 1            '☜: 최대 Columns의 항상 1개 증가시킴 
		    
		    Call GetSpreadColumnPos("A")
		    
		    ' SSSetEdit(Col, Header, ColWidth, HAlign, Row, Length, CharCase)
		    
			ggoSpread.SSSetEdit		C_SpType,		"계획구분", 10,,,1,2
			ggoSpread.SSSetEdit		C_LocExpFlag,	"거래구분", 10,,,1,2
			ggoSpread.SSSetEdit		C_SoBillFlag,	"수주매출구분", 10,,,1,2
			ggoSpread.SSSetEdit		C_SpPeriod,		"계획기간", 10,2,,8
		    ggoSpread.SSSetButton	C_SpPeriodPopup
			ggoSpread.SSSetEdit		C_SpPeriodDesc,"계획기간설명", 18,,,30
			Call AppendNumberPlace("6","3","0")
			ggoSpread.SSSetFloat	C_SpSeq,		"계획차수" ,10,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"  
			ggoSpread.SSSetEdit		C_SalesGrp,		"영업그룹", 18,,,4,2
			ggoSpread.SSSetEdit		C_SoldToParty, "고객", 18,,,10,2
		    ggoSpread.SSSetButton	C_SoldToPartyPopup
			ggoSpread.SSSetEdit		C_SoldToPartyNm, "고객명", 18
			ggoSpread.SSSetEdit		C_Cur,			"화폐", 8,2,,3,2
		    ggoSpread.SSSetButton	C_CurPopup
			ggoSpread.SSSetFloat	C_XchgRate,		"환율",15,parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ItemCd,		"품목", 18,,,18,2 
		    ggoSpread.SSSetButton	C_ItemCdPopup
			ggoSpread.SSSetEdit		C_ItemNm,		"품목명", 18
			ggoSpread.SSSetEdit		C_Spec,			"규격",			20
			ggoSpread.SSSetFloat	C_Qty,			"수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_Unit,			"단위", 8,2,,3,2
		    ggoSpread.SSSetButton	C_UnitPopup
			ggoSpread.SSSetFloat	C_Price,		"단가",15,"C",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Amt,			"금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_AmtLoc,		"자국금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
			ggoSpread.SSSetEdit		C_CfmFlag,		"확정여부", 10,2,,1,2
			ggoSpread.SSSetEdit		C_DistrFlag,	"배분여부", 10,2,,1,2
			ggoSpread.SSSetDate		C_FromDt,		"시작일", 10, 2, parent.gDateFormat
			ggoSpread.SSSetDate		C_ToDt,			"종료일", 10, 2, parent.gDateFormat
			ggoSpread.SSSetEdit		C_SpMonth,		"월", 10,2,,2
			ggoSpread.SSSetEdit		C_SpWeek,		"주", 10,2,,2
			ggoSpread.SSSetEdit		C_XChgRateOp,	"환율연산자", 10,,,1

			ggoSpread.SSSetEdit		C_Pointer,			"", 1
			ggoSpread.SSSetEdit		C_OldSpPeriod,		"", 1
			ggoSpread.SSSetEdit		C_OldSoldToParty,	"", 1
			ggoSpread.SSSetEdit		C_OldQty,			"", 1
			ggoSpread.SSSetEdit		C_OldAmt,			"", 1

			Call ggoSpread.MakePairsColumn(C_SpPeriod,C_SpPeriodPopup)
			Call ggoSpread.MakePairsColumn(C_SoldToParty,C_SoldToPartyPopup)
			Call ggoSpread.MakePairsColumn(C_Cur,C_CurPopup)
			Call ggoSpread.MakePairsColumn(C_ItemCd,C_ItemCdPopup)
			Call ggoSpread.MakePairsColumn(C_Unit,C_UnitPopup)
		    
		    Call ggoSpread.SSSetColHidden(C_SpType,C_SoBillFlag,True)
		    Call ggoSpread.SSSetColHidden(C_SpSeq,C_SalesGrp,True)
		    Call ggoSpread.SSSetColHidden(C_AmtLoc,C_AmtLoc,True)
		    Call ggoSpread.SSSetColHidden(C_XchgRateOp,.MaxCols,True)   '☜: 공통콘트롤 사용 Hidden Column
		    
   		    Call SetSpreadLock()

			.ReDraw = True
		End With
	End If
    
   	'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SP2CRP" Then
		With frm1.vspdData2		
			
		   	'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
			ggoSpread.Source = frm1.vspdData2
			'patch version
		    ggoSpread.Spreadinit "V20021214",,parent.gAllowDragDropSpread    		
			.ReDraw = false
			
			.MaxRows = 0 : .MaxCols = 0
			.MaxCols = C_TotAmt + 1            '☜: 최대 Columns의 항상 1개 증가시킴 
		    
		    Call GetSpreadColumnPos("B")
		    
		    ' SSSetEdit(Col, Header, ColWidth, HAlign, Row, Length, CharCase)
			ggoSpread.SSSetEdit		C_SpPeriod2,		"계획기간", 18,,,8
			ggoSpread.SSSetEdit		C_SpPeriodDesc2,	"계획기간설명", 18,,,30
			ggoSpread.SSSetEdit		C_SoldToParty2,		"고객", 18,,,10,2
			ggoSpread.SSSetEdit		C_SoldToPartyNm2,	"고객명", 18
			ggoSpread.SSSetFloat	C_TotQty,			"수량" ,15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_Cur2,				"화폐", 8,2,,3,2
			ggoSpread.SSSetFloat	C_TotAmt,			"금액",15,"A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		    Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)   '☜: 공통콘트롤 사용 Hidden Column
		    
		    ' Lock the sheet
		    Call SetSpreadLock2()
		    .OperationMode = 3
			.ReDraw = True
		End With
	End If
End Sub

'==========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock C_SpPeriodDesc, -1, C_SpPeriodDesc
	ggoSpread.SpreadLock C_SoldToPartyNm, -1, C_SoldToPartyNm
	ggoSpread.SpreadLock C_ItemNm, -1, C_ItemNm
	ggoSpread.SpreadLock C_Spec, -1, C_Spec
	ggoSpread.SpreadLock C_CfmFlag, -1, C_DistrFlag
	ggoSpread.SpreadLock C_ToDt, -1	
End Sub

Sub SetSpreadLock2()
	ggoSpread.SpreadLock 1, -1
End Sub

'==========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	' 새로이 등록한 경우 
	With frm1.vspdData
		.Col = C_FromDt
		If .ColHidden Then
			ggoSpread.SSSetRequired  C_SpPeriod		, pvStartRow, pvEndRow
		Else
			ggoSpread.SSSetRequired  C_FromDt		, pvStartRow, pvEndRow
			.Col = C_SpPeriod
			If Not .ColHidden Then
				ggoSpread.SSSetRequired  C_SpPeriod		, pvStartRow, pvEndRow
			End If
		End If
	End With
	
	ggoSpread.SSSetProtected C_SpPeriodDesc , pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_SoldToParty	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SoldToPartyNm, pvStartRow, pvEndRow

	ggoSpread.SSSetRequired  C_Cur			, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_XchgRate		, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_ItemCd		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ItemNm		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_Spec			, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_Qty			, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_Unit			, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_Price		, pvStartRow, pvEndRow
	ggoSpread.SSSetRequired  C_Amt			, pvStartRow, pvEndRow

	ggoSpread.SSSetProtected C_CfmFlag		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_DistrFlag	, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_ToDt			, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SpMonth		, pvStartRow, pvEndRow
	ggoSpread.SSSetProtected C_SpWeek		, pvStartRow, pvEndRow
End Sub

' Afetr query
Sub SetQuerySpreadColor(ByVal pvStartRow)
	Dim iLngLastCfmRow, iLngFirstNotCfmRow
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		ggoSpread.SSSetProtected C_SpPeriod			, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_SpPeriodPopup	, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_SoldToParty		, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_SoldToPartyPopup	, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_Cur				, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_CurPopup			, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_ItemCd			, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_ItemCdPopup		, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_Unit				, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_UnitPopup		, pvStartRow, .MaxRows
		ggoSpread.SSSetProtected C_FromDt			, pvStartRow, .MaxRows

		iLngLastCfmRow = 0
		iLngFirstNotCfmRow = 0
		.Row = pvStartRow	: .Col = C_CfmFlag
		If .Text = "Y" Then
			For iLngLastCfmRow = pvStartRow to .MaxRows
				.Row = iLngLastCfmRow
				If .Text = "N" Then
					iLngFirstNotCfmRow = iLngLastCfmRow
					iLngLastCfmRow = iLngLastCfmRow - 1
					Exit For
				End If
			Next
		Else
			iLngFirstNotCfmRow = pvStartRow
		End If

		If iLngLastCfmRow > 0 Then
			ggoSpread.SSSetProtected  C_XchgRate		, pvStartRow, iLngLastCfmRow
			ggoSpread.SSSetProtected  C_Qty				, pvStartRow, iLngLastCfmRow
			ggoSpread.SSSetProtected  C_Price			, pvStartRow, iLngLastCfmRow
			ggoSpread.SSSetProtected  C_Amt				, pvStartRow, iLngLastCfmRow
		End If

		If iLngFirstNotCfmRow > 0 Then
			ggoSpread.SSSetRequired  C_XchgRate			, iLngFirstNotCfmRow, .MaxRows
			ggoSpread.SSSetRequired  C_Qty				, iLngFirstNotCfmRow, .MaxRows
			ggoSpread.SSSetRequired  C_Price			, iLngFirstNotCfmRow, .MaxRows
			ggoSpread.SSSetRequired  C_Amt				, iLngFirstNotCfmRow, .MaxRows
		End If
	End With
	
End Sub

'==========================================================================================================
' Desc : This method set focus to position of error
'      : This method is called in MB area
'==========================================================================================================
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
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub
'==========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_SpType			= iCurColumnPos(1)
			C_LocExpFlag		= iCurColumnPos(2)
			C_SoBillFlag		= iCurColumnPos(3)
			C_SpPeriod			= iCurColumnPos(4)
			C_SpPeriodPopup		= iCurColumnPos(5)
			C_SpPeriodDesc		= iCurColumnPos(6)
			C_SpSeq				= iCurColumnPos(7)
			C_SalesGrp			= iCurColumnPos(8)
			C_SoldToParty		= iCurColumnPos(9)
			C_SoldToPartyPopup	= iCurColumnPos(10)
			C_SoldToPartyNm		= iCurColumnPos(11)
			C_Cur				= iCurColumnPos(12)
			C_CurPopup			= iCurColumnPos(13)
			C_XchgRate			= iCurColumnPos(14)
			C_ItemCd			= iCurColumnPos(15)
			C_ItemCdPopup		= iCurColumnPos(16)
			C_ItemNm			= iCurColumnPos(17)
			C_Spec				= iCurColumnPos(18)
			C_Qty				= iCurColumnPos(19)
			C_Unit				= iCurColumnPos(20)
			C_UnitPopup			= iCurColumnPos(21)
			C_Price				= iCurColumnPos(22)
			C_Amt				= iCurColumnPos(23)
			C_AmtLoc			= iCurColumnPos(24)
			C_CfmFlag			= iCurColumnPos(25)
			C_DistrFlag			= iCurColumnPos(26)
			C_FromDt			= iCurColumnPos(27)
			C_ToDt				= iCurColumnPos(28)
			C_SpMonth			= iCurColumnPos(29)
			C_SpWeek			= iCurColumnPos(30)
			C_XchgRateOp		= iCurColumnPos(31)
			C_Pointer			= iCurColumnPos(32)
			C_OldSpPeriod		= iCurColumnPos(33)
			C_OldSoldToParty	= iCurColumnPos(34)
			C_OldQty			= iCurColumnPos(35)
			C_OldAmt			= iCurColumnPos(36)

       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_SpPeriod2			= iCurColumnPos(1)
			C_SpPeriodDesc2		= iCurColumnPos(2)
			C_SoldToParty2		= iCurColumnPos(3)
			C_SoldToPartyNm2	= iCurColumnPos(4)
			C_TotQty			= iCurColumnPos(5)
			C_Cur2				= iCurColumnPos(6)
			C_TotAmt			= iCurColumnPos(7)
    End Select    
End Sub

'==========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' 판매계획유형 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0023", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConSpType,lgF0,lgF1,parent.gColSep)
	Call SetCombo2(frm1.cboSpType,lgF0,lgF1,parent.gColSep)

	'거래구분 
	Call CommonQueryRs(" MINOR_CD,MINOR_NM", " B_MINOR ", " MAJOR_CD=" & FilterVar("S4225", "''", "S") & " ", _
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	Call SetCombo2(frm1.cboConLocExpFlag, lgF0,lgF1, parent.gColSep)
	Call SetCombo2(frm1.cboLocExpFlag, lgF0,lgF1, parent.gColSep)

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
'==========================================================================================================
Sub InitData()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'==========================================================================================================
Function CookiePage(Byval pvKubun)

	On Error Resume Next
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim iStrTemp, iArrVal

	With frm1
		If pvKubun = 1 Then
			WriteCookie CookieSplit , .txtConSalesGrp.value & Parent.gColSep & .txtConSalesGrpNm.value & Parent.gColSep & .cboConSpType.value
		ElseIf pvKubun = 0 Then
			iStrTemp = ReadCookie(CookieSplit)
			
			If Trim(Replace(iStrTemp, parent.gColSep, "")) = "" then
				' 판매계획유형을 실행판매계획으로 Default 설정 
				.cboConSpType.value = "E"
				 Exit Function
			End If
			
			iArrVal = Split(iStrTemp, Parent.gColSep)

			.txtConSalesGrp.value	= iArrVal(0)
			.txtConSalesGrpNm.value = iArrVal(1)
			.txtConFrSPPeriod.value = iArrVal(2)
			.txtConFrSPPeriodDesc.value = iArrVal(3)
			.txtConToSPPeriod.value		= iArrVal(4)
			.txtConToSPPeriodDesc.value = iArrVal(5)
			.cboConSpType.value = iArrVal(6)
			WriteCookie CookieSplit , ""
		End If
	End With
End Function
'==========================================================================================================
Function JumpChgCheck(byVal pvStrJumpPgmId)

	Dim IntRetCD

	'************ 멀티인 경우 **************
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call CookiePage(1)
	Call PgmJump(pvStrJumpPgmId)

End Function

'==========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029             '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	'----------  Coding part  -------------------------------------------------------------
	Call InitComboBox()
	Call InitSpreadSheet
	Call CookiePage(0)
	Call SetDefaultVal    
	Call InitVariables              '⊙: Initializes local global variables

	lgBlnOpenedflag = True
	Call SetToolbar("11000000000011")          '⊙: 버튼 툴바 제어 
End Sub

'==========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    FncQuery = False                                  <%'⊙: Processing is NG%>
    
    Err.Clear             
                                                      <%'☜: Protect system from crashing%>
    '-----------------------
    'Check previous data area
    '-----------------------
	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
    End If

	'-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")          <%'⊙: Clear Contents  Field%>
    Call ggoSpread.ClearSpreadData()
    Call InitVariables               <%'⊙: Initializes local global variables%>
	'-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery                <%'☜: Query db data%>

    FncQuery = True                <%'⊙: Processing is OK%>
        
End Function

'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		'Call MsgBox("No data changed!!", vbInformation)
		Exit Function
	End If
    
<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Not chkField(Document, "2") Then     <%'⊙: Check contents area%>
       Exit Function
    End If
    If ggoSpread.SSDefaultCheck = False Then     <%'⊙: Check contents area%>
       Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll  DbSave                                                   <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function

'========================================================================================================
Function FncCopy() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	FncCopy = False
	
	With frm1.vspdData
		.ReDraw = False
		.focus
			 
		ggoSpread.Source = frm1.vspdData 
		ggoSpread.CopyRow
		
		SetSpreadColor .ActiveRow, .ActiveRow

		Call FormatSpreadCellByCurrency(.ActiveRow, .ActiveRow, "I")
		
		Call SetRowCopyDefaultVal(.ActiveRow)
		.ReDraw = True
	End With

	lgBlnFlgChgValue = True

	If Err.number = 0 Then  FncCopy = True				                                '☜: Processing is OK
	
    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================================
Function FncCancel() 
	On Error Resume Next
	Dim iDblNewQty, iDblNewAmt, iDblOldQty, iDblOldAmt
	Dim iStrFlag
	
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
	frm1.vspdData.ReDraw = False
	
		ggoSpread.Source = frm1.vspdData 
	
		With frm1.vspdData
			.Row = .ActiveRow
			.Col = C_OldQty	: iDblOldQty = UNICDbl(.Text)
			.Col = C_OldAmt	: iDblOldAmt = UNICDbl(.Text)
			.Col = 0		: iStrFlag = .Text
		    
			Select Case	iStrFlag
				Case ggoSpread.InsertFlag
					If iDblOldQty + iDblOldAmt > 0 Then
						Call ReCalcSpread2(.ActiveRow,-iDblOldQty,-iDblOldAmt,0,"")
					End If
				    ggoSpread.EditUndo
				    
				Case ggoSpread.UpdateFlag
				    ggoSpread.EditUndo
					.Col = C_Qty	:	iDblNewQty = UNICDbl(.Text)
					.Col = C_Amt	:	iDblNewAmt = UNICDbl(.Text)
					
					Call ReCalcSpread2(.ActiveRow, iDblNewQty - iDblOldQty, iDblNewAmt - iDblOldAmt,0,"")

				Case ggoSpread.DeleteFlag
				    ggoSpread.EditUndo
					.Col = C_Qty	:	iDblNewQty = UNICDbl(.Text)
					.Col = C_Amt	:	iDblNewAmt = UNICDbl(.Text)
					
					Call ReCalcSpread2(.ActiveRow, iDblNewQty, iDblNewAmt,0,"")
			End Select
			
			Call FormatSpreadCellByCurrency(.ActiveRow, ActiveRow, "I")

		End With

	frm1.vspdData.ReDraw = True

End Function

'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	On Error Resume Next                                                          '☜: If process fails

    Dim iIntInsertedRows
	Dim iIntActiveRow

    Err.Clear
    
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        iIntInsertedRows = CInt(pvRowCnt)
    Else
        iIntInsertedRows = AskSpdSheetAddRowCount()
        If iIntInsertedRows = "" Then
            Exit Function
        End If
    End If
   
	With frm1.vspdData
		.ReDraw = False
		.focus
		ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow, iIntInsertedRows
				
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		iIntActiveRow = .ActiveRow
		
		SetSpreadColor .ActiveRow, .ActiveRow + iIntInsertedRows - 1
		
		' 추가된 Row의 Default 값 설정 
		Call SetRowDefaultVal(iIntInsertedRows)
		
		' set the focus
		Call SubSetErrPos(iIntActiveRow)
		
		'------ Developer Coding part (End )   --------------------------------------------------------------	
		.ReDraw = True
						 
		lgBlnFlgChgValue = True
	End With
	
	If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If    
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function FncDeleteRow() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function

	Dim iLngRow, iLngFirstRow, iLngLastRow
	Dim iDblQty, iDblAmt
	With frm1.vspdData  
		.focus
		ggoSpread.Source = frm1.vspdData 
		<% '----------  Coding part  -------------------------------------------------------------%>   
		iLngFirstRow = .SelBlockRow
		If iLngFirstRow = -1 Then
			iLngFirstRow = 1
			iLngLastRow = .MaxRows
			Exit Function
		Else
			iLngLastRow = .SelBlockRow2
		End If
		
		.Col = 0
		For	iLngRow = iLngFirstRow To iLngLastRow
			.Row = iLngRow
			If .Text <> ggoSpread.DeleteFlag And .Text <> ggoSpread.InsertFlag Then
				.Col = C_Qty	: iDblQty = UNICDbl(.Text)
				.Col = C_Amt	: iDblAmt = UNICDbl(.Text)
				Call ReCalcSpread2(iLngRow,-iDblQty,-iDblAmt,0,"")
			End If
		Next

		Call ggoSpread.DeleteRow
		
		lgBlnFlgChgValue = True
	End With
    
End Function

'========================================================================================================
Function FncPrint() 
 Call parent.FncPrint()
End Function

'========================================================================================================
Function FncExcel() 
 On Error Resume Next                                                             '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call parent.FncExport(Parent.C_SINGLEMULTI)	                     			  '☜: 화면 유형 

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function FncFind()
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

    Call parent.FncFind(Parent.C_SINGLEMULTI, False)                              '☜:화면 유형, Tab 유무 
    
    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Sub FncSplitColumn()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Sub
	End If

	ggoSpread.Source = gActiveSpdSheet
	ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'========================================================================================================
Function FncExit()
 Dim IntRetCD
 FncExit = False

	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
Function DbQuery() 
	Err.Clear                                                               <%'☜: Protect system from crashing%>
	    
	DbQuery = False                                                         <%'⊙: Processing is NG%>
	If  LayerShowHide(1) = False Then
		Exit Function 
	End If
	    
	Dim iStrVal
	With Frm1
		iStrVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         <%'☜: 비지니스 처리 ASP의 상태 %>

	    If lgIntFlgMode = Parent.OPMD_UMODE Then    
			' Scroll
			iStrVal = iStrVal & lgStrWhere
		Else
			' Initial query
			lgStrWhere = "&txtWhere="
			lgStrWhere = lgStrWhere & .cboConSpType.value & parent.gColSep							' Sales Plan Type
			lgStrWhere = lgStrWhere & .cboConLocExpFlag.value & parent.gColSep		' Sales Order/Billing Flag
			lgStrWhere = lgStrWhere & "2" & parent.gColSep							' Local/Export Flag
			lgStrWhere = lgStrWhere & .txtConFrSPPeriod.value & parent.gColSep		' Sales Plan Period
			lgStrWhere = lgStrWhere & .txtConToSPPeriod.value & parent.gColSep		' Sales Plan Period
			lgStrWhere = lgStrWhere & "0" & parent.gColSep							' Sales Plan Sequence
			lgStrWhere = lgStrWhere & .txtConSalesGrp.value & parent.gColSep		' Sales Group
			lgStrWhere = lgStrWhere & .txtConSoldToParty.value & parent.gColSep		' Slod to party
			lgStrWhere = lgStrWhere & .txtConItemCd.value '& parent.gColSep			' Item Code

			iStrVal = iStrVal & lgStrWhere
		End If 
		iStrVal = iStrVal & "&lgStrPrevKey=" & lgStrPrevKey
		iStrVal = iStrVal & "&txtLastRow=" & frm1.vspdData.MaxRows

		lgLngStartRow = frm1.vspdData.MaxRows + 1
	End With
	Call RunMyBizASP(MyBizASP, iStrVal)            <%'☜: 비지니스 ASP 를 가동 %>

	DbQuery = True   
                          <%'⊙: Processing is NG%>
End Function

'========================================================================================================
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then
		lgIntFlgMode = Parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>
		' Set the defaut data for input
		Call GetLastCfmSpPeriod()
		Call GetSpConfig()
	End If
    
	If Trim(lgStrPrevKey) = "" Then
		lgStrWhere = ""
    End If

	Call SetQuerySpreadColor(lgLngStartRow)
	
	Call SetToolbar("11001111001111")
	 
	frm1.vspdData.focus
End Function

'========================================================================================================
Function DbSave() 

	Err.Clear                <%'☜: Protect system from crashing%>
	 
	Dim iStrIns, iStrUpd, iStrDel, iStrKey
	Dim iLngRow
		 
	DbSave = False                                                          '⊙: Processing is NG
		    
	On Error Resume Next                                                   '☜: Protect system from crashing
		   
	If LayerShowHide(1) = False Then
		Exit Function 
	End If

  '-----------------------
  'Data manipulate area
  '-----------------------
  iStrInt = ""
  iStrUpd = ""
  iStrDel = ""
   
  '-----------------------
  'Data manipulate area
  '-----------------------
	With frm1.vspdData
		For iLngRow = 1 To .MaxRows
    
			.Row = iLngRow
			.Col = 0

			if .Text <> "" Then
				iStrKey = CStr(iLngRow) & Parent.gColSep		' Row No.
				.Col = C_SpType			' Sales planning type
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_LocExpFlag		' Local/Export Flag(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_SoBillFlag		' Sales Order/Billing Flag(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_SpPeriod		' Sales Plan Period(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_SpSeq			' Sales Plan Sequence(PK)
				iStrKey = iStrKey & UNIConvNum(.Text,0) & Parent.gColSep
				.Col = C_SalesGrp       ' Sales Group(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_SoldToParty    ' Slod to party(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				.Col = C_ItemCd			' Item Code(PK)
				iStrKey = iStrKey & .Text & Parent.gColSep
				
				.Col = 0
				Select Case .Text
					Case ggoSpread.InsertFlag       '☜: 신규 
						iStrIns = iStrIns & iStrKey

						.Col = C_Cur		' Currency
						iStrIns = iStrIns & .Text & Parent.gColSep
						
						.Col = C_XchgRate	' Exchange rate
						iStrIns = iStrIns & UNIConvNum(.Text,0) & Parent.gColSep
						
						.Col = C_XchgRateOp	' Exchange rate Operator
						iStrIns = iStrIns & .Text & Parent.gColSep

						.Col = C_Unit		' Item unit
						iStrIns = iStrIns & .Text & Parent.gColSep
						
						.Col = C_Qty		' Quantity
						iStrIns = iStrIns & UNIConvNum(.Text,0) & Parent.gColSep
						
						.Col = C_Price		' Pirce
						iStrIns = iStrIns & UNIConvNum(.Text,0) & Parent.gColSep
						
						.Col = C_Amt		' Amount
						iStrIns = iStrIns & UNIConvNum(.Text,0) & Parent.gColSep
						
						.Col = C_AmtLoc		' Amount(local)
						iStrIns = iStrIns & parent.gCurrency & Parent.gColSep
						iStrIns = iStrIns & "0" & Parent.gColSep
						
						.Col = C_CfmFlag	' Confirmed flag
						iStrIns = iStrIns & .Text & Parent.gColSep
			
						.Col = C_DistrFlag	' Deleted flag
						iStrIns = iStrIns & .Text & Parent.gColSep

						iStrIns = iStrIns & Parent.gUsrID & Parent.gColSep & Parent.gRowSep

					Case ggoSpread.UpdateFlag       '☜: 수정 
						iStrUpd = iStrUpd & iStrKey
						
						.Col = C_XchgRate	' Exchange rate
						iStrUpd = iStrUpd & UNIConvNum(.Text,0) & Parent.gColSep
						
						.Col = C_XchgRateOp	' Exchange rate Operator
						iStrUpd = iStrUpd & .Text & Parent.gColSep

						.Col = C_Qty		' Quantity
						iStrUpd = iStrUpd & UNIConvNum(.Text,0) & Parent.gColSep
						
						.Col = C_Price		' Pirce
						iStrUpd = iStrUpd & UNIConvNum(.Text,0) & Parent.gColSep
						
						.Col = C_Amt		' Amount
						iStrUpd = iStrUpd & UNIConvNum(.Text,0) & Parent.gColSep

						.Col = C_AmtLoc		' Amount(local)
						iStrUpd = iStrUpd & parent.gCurrency & Parent.gColSep
						iStrUpd = iStrUpd & 0 & Parent.gColSep
						
						iStrUpd = iStrUpd & Parent.gUsrID & Parent.gColSep & Parent.gRowSep

					Case ggoSpread.DeleteFlag       '☜: 삭제 
						iStrDel = iStrDel & iStrKey & Parent.gRowSep
				End Select
			End If
		Next
	End With
 
	With frm1
	  .txtMode.value = Parent.UID_M0002
	  .txtSpreadIns.value = iStrIns
	  .txtSpreadUpd.value = iStrUpd
	  .txtSpreadDel.value = iStrDel
	End With

 	Call ExecMyBizASP(frm1, BIZ_PGM_ID)         '☜: 비지니스 ASP 를 가동 
 
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================================
Function DbSaveOk()               <%'☆: 저장 성공후 실행 로직 %>
	With frm1
		.txtConSalesGrp.value = .txtSalesGrp.value
		.txtConSalesGrpNm.value = .txtSalesGrpNm.value
		.cboConLocExpFlag.value = .cboLocExpFlag.value
	End With
	Call ggoOper.ClearField(Document, "2")
    Call MainQuery()
End Function

'==========================================================================================
'   Event Desc : 판매계획유형 
'==========================================================================================
Function cboConSpType_OnChange()
	If frm1.txtConSalesGrp.value <> "" Then
		Call GetCfmPeriod(0)
	End If
End Function

'==========================================================================================
'   Event Desc : 영업그룹 
'==========================================================================================
Function txtConSalesGrp_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtConSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				.txtConSalesGrp.value = ""
				.txtConSalesGrpNm.value = ""
				.txtConSalesGrp.focus
			End If
			txtConSalesGrp_OnChange = False
		Else
			.txtConSalesGrpNm.value = ""
		End If
	End With
End Function

'========================================================================================================
' Popup
'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	Select Case pvIntWhere
	Case C_PopSalesGrp												
		iArrParam(1) = "B_SALES_GRP"						<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConSalesGrp.value)		<%' Code Condition%>
		iArrParam(3) = ""									<%' Name Cindition%>
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						<%' Where Condition%>
		iArrParam(5) = "영업그룹"						<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"			<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"			<%' Field명(1)%>
    
	    iArrHeader(0) = "영업그룹"						<%' Header명(0)%>
	    iArrHeader(1) = "영업그룹명"					<%' Header명(1)%>

		frm1.txtConSalesGrp.focus 

	Case C_PopFrSpPeriod
		OpenConPopup = OpenConSpPeriodPopup(C_PopFrSpPeriod, frm1.txtConFrSPPeriod.value)
		frm1.txtConFrSPPeriod.focus
		Exit Function
	
	Case C_PopToSpPeriod
		OpenConPopup = OpenConSpPeriodPopup(C_PopToSpPeriod, frm1.txtConToSPPeriod.value)
		frm1.txtConToSPPeriod.focus
		Exit Function

	Case C_PopSoldToParty												
		iArrParam(1) = "B_BIZ_PARTNER BP"					<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConSoldToParty.value)	<%' Code Condition%>
		iArrParam(3) = ""									<%' Name Cindition%>
		iArrParam(4) = "BP.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND BP.BP_TYPE LIKE " & FilterVar("C%", "''", "S") & " "	<%' Where Condition%>
		iArrParam(5) = "고객"							<%' TextBox 명칭 %>
			
		iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
		iArrField(2) = "ED8" & Parent.gColSep & "BP.CURRENCY"
		    
		iArrHeader(0) = "고객"
		iArrHeader(1) = "고객명"
		iArrHeader(2) = "화폐"

		frm1.txtConSoldToParty.focus

	Case C_PopItemCd
		OpenConPopup = OpenConItemPopup(C_PopItemCd, frm1.txtConItemCd.value)
		frm1.txtConItemCd.focus
		Exit Function

	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

' Sales planning period Popup
Function OpenConSpPeriodPopup(ByVal pvIntWhere, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(4)
	Dim iCalledAspName

	OpenConSpPeriodPopup = False

	iCalledAspName = AskPRAspName("s2211pa3")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2211pa3", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	iArrParam(4) = frm1.cboConSpType.value
	
	iArrRet = window.showModalDialog(iCalledAspName & "?txtDisplayFlag=N", Array(window.parent,iArrParam), _
	 "dialogWidth=690px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConSpPeriodPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
End Function

' Item Popup
Function OpenConItemPopup(ByVal pvIntWhere, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(3)
	Dim iCalledAspName

	OpenConItemPopup = False

	iCalledAspName = AskPRAspName("s2210pa1")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2210pa1", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
	 "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConItemPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	
End Function

' Spread button popup
Function OpenSpreadPopup(ByVal pvLngCol, ByVal pvLngRow, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenSpreadPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	Select Case pvLngCol
		Case C_SpPeriodPopup
			OpenSpreadPopup = OpenSpreadSpPeriodPopup(pvLngRow, pvStrData)
			Exit Function
	
		Case C_SoldToPartyPopup
			iArrParam(1) = "B_BIZ_PARTNER BP"			<%' TABLE 명칭 %>
			iArrParam(2) = pvStrData					<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = "BP.USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND BP.BP_TYPE LIKE " & FilterVar("C%", "''", "S") & " "	<%' Where Condition%>
			iArrParam(5) = "고객"						<%' TextBox 명칭 %>
				
			iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
			iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
			iArrField(2) = "ED8" & Parent.gColSep & "BP.CURRENCY"
			    
			iArrHeader(0) = "고객"
			iArrHeader(1) = "고객명"
			iArrHeader(2) = "화폐"

		Case C_ItemCdPopup
			OpenSpreadPopup = OpenSpreadItemPopup(pvLngRow, pvStrData)
			Exit Function

		Case C_CurPopup
			iArrParam(1) = "dbo.B_CURRENCY "				<%' TABLE 명칭 %>
			iArrParam(2) = pvStrData						<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = ""								<%' Where Condition%>
			iArrParam(5) = "화폐"						<%' TextBox 명칭 %>
				
			iArrField(0) = "ED15" & Parent.gColSep & "CURRENCY"
			iArrField(1) = "ED30" & Parent.gColSep & "CURRENCY_DESC"
			    
			iArrHeader(0) = "화폐"
			iArrHeader(1) = "화폐명"

		Case C_UnitPopup
			iArrParam(1) = "dbo.B_UNIT_OF_MEASURE "			<%' TABLE 명칭 %>
			iArrParam(2) = pvStrData						<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = " DIMENSION <> " & FilterVar("TM", "''", "S") & " "			<%' Where Condition%>
			iArrParam(5) = "단위"						<%' TextBox 명칭 %>
				
			iArrField(0) = "ED15" & Parent.gColSep & "UNIT"
			iArrField(1) = "ED30" & Parent.gColSep & "UNIT_NM"
			    
			iArrHeader(0) = "단위"
			iArrHeader(1) = "단위명"
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSpreadPopup = SetSpreadPopup(iArrRet,pvLngCol, pvLngRow)
	End If	

End Function

' Sales planning period Popup
Function OpenSpreadSpPeriodPopup(ByVal pvLngRow, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(4)
	Dim iCalledAspName

	OpenSpreadSpPeriodPopup = False

	iCalledAspName = AskPRAspName("s2211pa3")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2211pa3", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	iArrParam(1) = lgStrLastCfmSpPeriod				' 최종확정 기간 
	iArrParam(2) = lgStrLastCfmSpPeriodDesc
	iArrParam(3) = lgDtLastCfmToDt
	iArrParam(4) = frm1.cboSpType.value
	
	iArrRet = window.showModalDialog(iCalledAspName & "?txtDisplayFlag=Y", Array(window.parent,iArrParam), _
	 "dialogWidth=690px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSpreadSpPeriodPopup = SetSpreadPopup(iArrRet, C_SpPeriodPopup, pvLngRow)
	End If	
End Function

' Item Popup
Function OpenSpreadItemPopup(ByVal pvLngRow, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(2)
	Dim iCalledAspName

	OpenSpreadItemPopup = False

	iCalledAspName = AskPRAspName("s2210pa1")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2210pa1", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = pvStrData
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_FromDt		:	iArrParam(1) = .Text
		.Col = C_ToDt		:	iArrParam(2) = .Text
	End With
		
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
	 "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSpreadItemPopup = SetSpreadPopup(iArrRet, C_ItemCdPopup, pvLngRow)
	End If	
End Function

'===========================================================================================================
' POP-UP (Set)
'===========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopSalesGrp
			.txtConSalesGrp.value = pvArrRet(0) 
			.txtConSalesGrpNm.value = pvArrRet(1)   
			Call GetCfmPeriod(0)
			
		Case C_PopFrSpPeriod
			.txtConFrSPPeriod.value = pvArrRet(0)
			.txtConFrSPPeriodDesc.value = pvArrRet(1)

		Case C_PopToSpPeriod
			.txtConToSPPeriod.value = pvArrRet(0)
			.txtConToSPPeriodDesc.value = pvArrRet(1)

		Case C_PopSoldToParty
			frm1.txtConSoldToParty.value = pvArrRet(0) 
			frm1.txtConSoldToPartyNm.value = pvArrRet(1)   

		Case C_PopItemCd
			frm1.txtConItemCd.value = pvArrRet(0) 
			frm1.txtConItemNm.value = pvArrRet(1)   
		End Select
	End With

	SetConPopup = True
End Function

Function SetSpreadPopup(Byval pvArrRet,ByVal pvLngCol, ByVal pvLngRow)
	SetSpreadPopup = False

	With frm1.vspdData
		.Row = pvLngRow
		
		Select Case pvLngCol
			Case C_SpPeriodPopup
				.Col = C_SpPeriod		: .Text = pvArrRet(0)
				.Col = C_SpPeriodDesc	: .Text = pvArrRet(1)
				.Col = C_SpMonth		: .Text = pvArrRet(6)
				.Col = C_FromDt			: .Text = pvArrRet(2)
				.Col = C_ToDt			: .Text = pvArrRet(3)
				.Col = C_SpWeek			: .Text = pvArrRet(4)

			Case C_SoldToPartyPopup
				.Col = C_SoldToParty	: .Text = pvArrRet(0)
				.Col = C_SoldToPartyNm	: .Text = pvArrRet(1)
				.Col = C_Cur			: .Text = pvArrRet(2)

			Case C_ItemCdPopup
				.Col = C_ItemCd			: .Text = pvArrRet(0)
				.Col = C_ItemNm			: .Text = pvArrRet(1)
				.Col = C_Spec			: .Text = pvArrRet(2)
				.Col = C_Unit			: .Text = pvArrRet(3)

			Case C_CurPopup
				.Col = C_Cur			: .Text = pvArrRet(0)

			Case C_UnitPopup
				.Col = C_Unit			: .Text = pvArrRet(0)

		End Select
	End With

	SetSpreadPopup = True
End Function

'===========================================================================================================
'   Event Desc : Update the Row Status
'===========================================================================================================
Sub SetRowStatus(intRow)
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow intRow

 lgBlnFlgChgValue = True
End Sub


'========================================================================================================
' Spread Popup Method
'========================================================================================================
'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	 '------ Developer Coding part (Start ) --------------------------------------------------------------
	If gMouseClickStatus = "SPCRP" Then	SetQuerySpreadColor(1)
	
	Call FormatSpreadCellByCurrency(-1, -1, "Q")
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
' Spread OCX Tag Event
'========================================================================================================
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

 <% '----------  Coding part  -------------------------------------------------------------%>   
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		   
		If Row > 0 Then
			Select Case Col
				CASE C_SpPeriodPopup
					.Col = C_SpPeriod
					If OpenSpreadPopup(Col, Row, .Text) Then
						Call GetXchgRate(Row)
						Call ReCalcSpread2BySpPeriod(Row)
						' 품목의 유효성 Check
						.Col = C_ItemCd
						If Not GetItemInfo(Row, .Text, True) Then
							.Col = C_ItemCd : .Text = ""
						End If
					End If
				     
				CASE C_SoldToPartyPopup
					.Col = C_SoldToParty
					If OpenSpreadPopup(Col, Row, .Text) Then
						Call GetXchgRate(Row)
						Call FormatSpreadCellByCurrency(Row, Row, "I")
					End If

				CASE C_ItemCdPopup
					.Col = C_ItemCd
					If OpenSpreadPopup(Col, Row, .Text) Then
						Call GetItemPrice(Row)
					End If

				CASE C_CurPopup
					.Col = C_Cur
					If OpenSpreadPopup(Col, Row, .Text) Then
						Call GetXchgRate(Row)
						Call FormatSpreadCellByCurrency(Row, Row, "I")
					End If

				CASE C_UnitPopup
					.Col = C_Unit
					If OpenSpreadPopup(Col, Row, .Text) Then
						Call GetItemPrice(Row)
					End If
			End Select
			
			Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")

		End If
	End With

End Sub
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
	
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

Sub vspdData2_Click(ByVal Col, ByVal Row)

	gMouseClickStatus = "SP2C"	'Split 상태코드 
	   
    Set gActiveSpdSheet = frm1.vspdData2
    
    ' spread1에서 spread2의 Pointer 갖고 있어 spread2의 정렬은 disalbe 시킴 
    Exit Sub

    If frm1.vspdData2.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData2
		
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

'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim iStrData, iStrOldSpPeriod
	Dim iDblOldAmt, iDblQty, lDblAmt
	
	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData
		.Row = Row
		.Col = 0
		If .Text = ggoSpread.DeleteFlag Then Exit Sub
		
		CALL SetRowStatus(Row)

		.Col = Col	: iStrData = .Text
		
		If iStrData = "" Then Exit Sub
		
		Select Case Col
			Case C_SpPeriod
				If Len(iStrData) = Len(lgStrLastCfmSpPeriod) AND iStrData <= lgStrLastCfmSpPeriod Then
					Call DisplayMsgBox("202406", "X", "", "")
					.Text = ""
				ElseIf GetSpPeriodInfo(Row, iStrData, True) Then
					Call GetXchgRate(Row)
					Call ReCalcSpread2BySpPeriod(Row)
					' Check validity of item code
					.Col = C_ItemCd
					If Not GetItemInfo(Row, .Text, True) Then
						.Col = C_ItemCd : .Text = ""
					End If
				Else
					.Text = ""
				End If

			Case C_SoldToParty
				If GetSoldToPartyInfo(Row, iStrData) Then
					Call GetXchgRate(Row)
					Call ReCalcSpread2BySoldToParty(Row)
					Call FormatSpreadCellByCurrency(Row, Row, "I")
				Else
					.Text = ""
				End If
				
			Case C_Cur
				Call GetXchgRate(Row)

				Call FormatSpreadCellByCurrency(Row, Row, "I")
				
			Case C_XchgRate
				' If the currency is the local currency, exchange rate must be '1'
				.Col = C_Cur
				If .Text = parent.gCurrency Then
					.Col = C_XchgRate	: .Text = "1"
				End If

			Case C_ItemCd
				If GetItemInfo(Row, iStrData, False) Then
					Call GetItemPrice(Row)
				Else
					.Col = C_ItemCd : .Text = ""
				End If
				
			Case C_Qty
				Call CalcAmt(Row, C_Qty)
				.Col = C_OldQty	: .Text = iStrData
				
			Case C_Unit
				Call GetItemPrice(Row)
				
			Case C_Price
				Call CalcAmt(Row, C_Price)
                
			Case C_Amt
				.Col = C_OldAmt	: iDblOldAmt = UNICDbl(.Text)
				.Text = iStrData
				Call ReCalcSpread2(Row, 0, UNICDbl(iStrData) - iDblOldAmt, C_Amt, iStrData)
				
			Case C_FromDt
				If UniConvDateToYYYYMMDD(iStrData, Parent.gDateFormat,"") <= UniConvDateToYYYYMMDD(lgDtLastCfmToDt, Parent.gDateFormat,"") Then
					Call DisplayMsgBox("202406", "X", "", "")
					.Text = ""
				ElseIf GetSpPeriodInfo(Row, iStrData, False) Then
					Call GetXchgRate(Row)
					Call ReCalcSpread2BySpPeriod(Row)
					' Check validity of item code
					.Col = C_ItemCd
					If Not GetItemInfo(Row, .Text, True) Then
						.Col = C_ItemCd : .Text = ""
					End If
				Else
					.Text = ""
				End If
							
		End Select
	End With

End Sub

'========================================================================================================
Sub vspdData_EditMode(ByVal Col, ByVal Row, ByVal Mode, ByVal ChangeMade)
    Select Case Col
        Case C_Price
            Call EditModeCheck(frm1.vspdData, Row, C_Cur, C_Price, "C" ,"I", Mode, "X", "X")        
		Case C_Amt
			Call EditModeCheck(frm1.vspdData, Row, C_Cur, C_Amt, "A" ,"I", Mode, "X", "X")        
		Case C_AmtLoc
			Call EditModeCheck(frm1.vspdData, Row, C_Cur, C_AmtLoc, "A" ,"I", Mode, "X", "X")        
    End Select
End Sub

'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

Sub vspdData2_ColWidthChange(ByVal Col1, ByVal Col2)
   ggoSpread.Source = frm1.vspdData2
  Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

Sub vspdData2_MouseDown(Button , Shift , x , y)

 If Button = 2 And gMouseClickStatus = "SP2C" Then
  gMouseClickStatus = "SP2CR"
 End If

End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	If OldLeft <> NewLeft Then	Exit Sub
	    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) And lgStrPrevKey <> "" Then
		If CheckRunningBizProcess = True Then Exit Sub
	    
		Call DisableToolBar(Parent.TBC_QUERY)
		If DBQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if    

End Sub

'========================================================================================================
' ETC
'========================================================================================================
<%'=============================================================================================
'	Description : 판매계획환경정보를 Fetch한다.
'==================================================================================================== %>
Sub GetSpConfig()

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	iStrSelectList = " PRICE_RULE, XCHG_RATE_FG, PM_NON_XCHG_RATE "
	iStrFromList = " dbo.S_SP_CONFIG "
	iStrWhereList = " SP_TYPE =  " & FilterVar(frm1.cboSpType.value , "''", "S") & ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, parent.gColSep)
		lgStrPriceRule = iArrRs(1)
		lgXchgRateFg = iArrRs(2)
		lgXPmNonXchgRate = iArrRs(3)
		lgBlnExistsSpConfig = True
	Else
		'판매계획환경설정 정보가 없습니다.
		Call DisplayMsgBox("202403", "X", "", "")
		Call SetToolbar("11001011000111")          '⊙: 버튼 툴바 제어 
	End if
End Sub

' 최종확정 정보를 Fetch한다.
Sub GetLastCfmSpPeriod()

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	iStrSelectList = " SP.SP_PERIOD, SP.SP_PERIOD_DESC, SP.TO_DT "
	iStrFromList = " dbo.S_SP_PERIOD_INFO SP, "
	iStrFromList = iStrFromList & "(SELECT MAX(TO_SP_PERIOD ) AS " & FilterVar("SP_PERIOD", "''", "S") & " " & _
								  "FROM dbo.S_SP_CFM_INFO_BY_SALES_GRP " & _
								  "WHERE SP_STEP = " & FilterVar("S2214BA1", "''", "S") & " AND SP_TYPE =  " & FilterVar(frm1.cboSpType.value , "''", "S") & "" & _
								  "AND SALES_GRP =  " & FilterVar(frm1.txtSalesGrp.value , "''", "S") & ") T"
	iStrWhereList = " SP.SP_PERIOD = T.SP_PERIOD AND SP.SP_TYPE =  " & FilterVar(frm1.cboSpType.value , "''", "S") & ""

	Err.Clear
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, parent.gColSep)
		lgStrLastCfmSpPeriod = iArrRs(1)
		lgStrLastCfmSpPeriodDesc = iArrRs(2)
		lgDtLastCfmToDt = UNIDateClientFormat(iArrRs(3))
	End if
End Sub

' 확정할 기간정보를 Fetch한다.
Function GetCfmPeriod(ByVal pvIntSpPeriodSeq)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	GetCfmPeriod = False
	
	With frm1
		iStrSelectList = " * "
		iStrFromList = "  dbo.ufn_s_GetCfmPeriod(" & FilterVar("S2214BA1", "''", "S") & ",  " & FilterVar(.txtConSalesGrp.value, "''", "S") & ", " & FilterVar("1", "''", "S") & " ,  " & FilterVar(.cboConSpType.value, "''", "S") & ", " & FilterVar("Y", "''", "S") & " , " & pvIntSpPeriodSeq & ") "
		iStrWhereList = ""
	
		Err.Clear
	
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrRs = Split(iStrRs, Chr(11))
			.txtConFrSPPeriod.value = iArrRs(1)
			.txtConFrSPPeriodDesc.value = iArrRs(2)
			
			GetCfmPeriod = True
		Else
			.txtConFrSPPeriod.value = ""
			.txtConFrSPPeriodDesc.value = ""
		End if
	End With
End Function

'===========================================================================================================
' Description : 판매계획기간 정보를 Fetch한다.
'===========================================================================================================
Function GetSpPeriodInfo(ByVal pvLngRow, ByVal pvStrData, ByVal pvBlnSpFlag)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrSpPeriodInfo
	
	GetSpPeriodInfo = False
	
	iStrSelectList = " SP_PERIOD, SP_PERIOD_DESC, FROM_DT, TO_DT, SP_MONTH, SP_WEEK "
	iStrFromList	  = " dbo.S_SP_PERIOD_INFO "
	If pvBlnSpFlag Then
		iStrWhereList  = " SP_TYPE =  " & FilterVar(frm1.cboSpType.value , "''", "S") & " AND SP_PERIOD =  " & FilterVar(pvStrData , "''", "S") & ""
	Else
		iStrWhereList  = " SP_TYPE =  " & FilterVar(frm1.cboSpType.value , "''", "S") & " AND FROM_Dt <=  " & FilterVar(UNIConvDate(pvStrData), "''", "S") & " AND TO_Dt >=  " & FilterVar(UNIConvDate(pvStrData), "''", "S") & ""
	End If

	Err.Clear
	    
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrSpPeriodInfo = Split(iStrRs, Chr(11))
		With frm1.vspdData
			.Row = pvLngRow
			.Col = C_SpPeriod		: .text = Trim(iArrSpPeriodInfo(1))
			.Col = C_SpPeriodDesc	: .text = Trim(iArrSpPeriodInfo(2))
			.Col = C_FromDt			: .text = UNIDateClientFormat(Trim(iArrSpPeriodInfo(3)))
			.Col = C_ToDt			: .text = UNIDateClientFormat(Trim(iArrSpPeriodInfo(4)))
			.Col = C_SpMonth		: .text = Trim(iArrSpPeriodInfo(5))
			.Col = C_SpWeek			: .text = Trim(iArrSpPeriodInfo(6))
		End With
		GetSpPeriodInfo = True
		Exit Function
	Else
		If Err.number = 0 Then
			If pvBlnSpFlag Then
				GetSpPeriodInfo = OpenSpreadPopup(C_SpPeriodPopup, pvLngRow, pvStrData)
			Else
				' 계획기간 정보가 없습니다.
				Call DisplayMsgBox("202402", "X", "", "")
			End If
		Else
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If

End Function

'===========================================================================================================
' Description : 거래처 정보를 Fetch한다.
'===========================================================================================================
Function GetSoldToPartyInfo(ByVal pvLngRow, ByVal pvStrData)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrSoldToPartyInfo
	
	GetSoldToPartyInfo = False
	
	iStrSelectList = " BP_CD, BP_NM, CURRENCY "
	iStrFromList   = " dbo.B_BIZ_PARTNER "
	iStrWhereList  = " BP_TYPE LIKE " & FilterVar("C%", "''", "S") & " AND USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND BP_CD =  " & FilterVar(pvStrData , "''", "S") & ""

	Err.Clear
	    
	'거래처정보 Fetch
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrSoldToPartyInfo = Split(iStrRs, Chr(11))
		With frm1.vspdData
			.Row = pvLngRow
			.Col = C_SoldToParty	: .text = Trim(iArrSoldToPartyInfo(1))
			.Col = C_SoldToPartyNm	: .text = Trim(iArrSoldToPartyInfo(2))
			.Col = C_Cur			: .text = Trim(iArrSoldToPartyInfo(3))
		End With
		GetSoldToPartyInfo = True
		Exit Function
	Else
		If Err.number = 0 Then
			GetSoldToPartyInfo = OpenSpreadPopup(C_SoldToPartyPopup, pvLngRow, pvStrData)
		Else
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If

End Function

'===========================================================================================================
' Description : 환율 Fetch
'===========================================================================================================
Function GetXchgRate(ByVal pvLngRow)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrCur, iStrFromDt, iStrSpMonth, iStrYYYYMM
	Dim iStrRs
	Dim iArrXchgRate
	
	GetXchgRate = False

	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_Cur			:	iStrCur = Trim(.Text)
		
		If iStrCur = parent.gCurrency Then
			.Col = C_XchgRate	:	.Text = "1"
			.Col = C_XchgRateOp	:	.Text = "*"
			GetXchgRate = True
			Exit Function
		End If
		.Col = C_FromDt			:	iStrFromDt = Trim(.Text)
		.Col = C_SpMonth		:	iStrSpMonth = Trim(.Text)
		
		If iStrCur = "" Or iStrFromDt = "" Then	Exit Function
		
		iStrYYYYMM = Left(UniConvDateToYYYYMMDD(iStrFromDt, Parent.gDateFormat,""), 4) & Right("0" & iStrSpMonth, 2)
		
		iStrSelectList  = "*"
		iStrFromList = " dbo.ufn_s_GetXchgRateForSalesPlanning( " & FilterVar(iStrCur, "''", "S") & "," _
																  & FilterVar(parent.gCurrency, "''", "S") & ", '" _
																  & UNIConvDate(iStrFromDt)	& "', '" _
																  & iStrYYYYMM			& "', '" _
																  & lgXchgRateFg	& "', '" _
																  & lgXPmNonXchgRate & "')"
		iStrWhereList = ""
		
		Err.Clear
	    
		'환율 Fetch
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrXchgRate = Split(iStrRs, parent.gColSep)
			
			.Col = C_XchgRate
			.Text = UNIFormatNumber(iArrXchgRate(1),ggExchRate.DecPoint,-2,0,ggExchRate.RndPolicy,ggExchRate.RndUnit)	
			.Col = C_XchgRateOp
			.Text = iArrXchgRate(2)
			GetXchgRate = True
			Exit Function
		Else
			If Err.number = 0 Then
				.Col = C_XchgRate	:	.Text = "0"
				.Col = C_XchgRateOp :	.Text = "*"
				' 환율정보가 존재하지 않습니다.
				If lgBlnDisplayMsg Then
					Call DisplayMsgBox("202407", "X", "", "")
					lgBlnDisplayMsg = False
				End If
			Else
				MsgBox Err.description 
				Err.Clear
				Exit Function
			End If
		End If
	End With

End Function

'===========================================================================================================
' Description : 품목정보를 Fetch한다.
'===========================================================================================================
Function GetItemInfo(ByVal pvLngRow, ByVal pvStrData, ByVal pvBlnChkFlag)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrItemInfo
	
	Err.Clear
	    
	GetItemInfo = False
	
	If pvStrData = "" Then Exit Function
	
	iStrSelectList = " ITEM_CD, ITEM_NM, BASIC_UNIT, SPEC "
	iStrFromList   = " dbo.B_ITEM "
	iStrWhereList  = " VALID_FLG = " & FilterVar("Y", "''", "S") & "  AND ITEM_CD =  " & FilterVar(pvStrData , "''", "S") & " "
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_FromDt
		If .Text <> "" Then
			iStrWhereList = iStrWhereList & " AND VALID_FROM_DT <=  " & FilterVar(UNIConvDate(.Text), "''", "S") & " "
		End If

		.Col = C_ToDt
		If .Text <> "" Then
			iStrWhereList = iStrWhereList & " AND VALID_TO_DT >=  " & FilterVar(UNIConvDate(.Text), "''", "S") & " "
		End If

		'품목정보 Fetch
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			If Not pvBlnChkFlag Then
				iArrItemInfo = Split(iStrRs, parent.gColSep)
				.Col = C_ItemCd	: .text = Trim(iArrItemInfo(1))
				.Col = C_ItemNm	: .text = Trim(iArrItemInfo(2))
				.Col = C_Unit	: .text = Trim(iArrItemInfo(3))
				.Col = C_Spec	: .text = Trim(iArrItemInfo(4))
			End If

			GetItemInfo = True
			Exit Function
		Else
			If Err.number = 0 Then
				If Not pvBlnChkFlag Then GetItemInfo = OpenSpreadPopup(C_ItemCdPopUP, pvLngRow, pvStrData)
			Else
				MsgBox Err.description 
				Err.Clear
				Exit Function
			End If
		End If
	End With
		
End Function

'===========================================================================================================
' Description : 품목단가 Fetch
'===========================================================================================================
Function GetItemPrice(ByVal pvLngRow)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrSoldToParty, iStrItemCd, iStrUnit, iStrCur, iStrFromDt
	Dim iStrRs
	Dim iArrPrice
	Dim iDblOldPrice
	
	GetItemPrice = False
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_SoldToParty	:	iStrSoldToParty = Trim(.Text)
		.Col = C_ItemCd			:	iStrItemCd = Trim(.Text)
		.Col = C_Unit			:	iStrUnit = Trim(.Text)
		.Col = C_Cur			:	iStrCur = Trim(.Text)
		.Col = C_FromDt			:	iStrFromDt = Trim(.Text)

		If iStrSoldToParty = "" Or iStrItemCd = "" Or iStrUnit = "" Or iStrCur = "" Or iStrFromDt = "" Then
			Exit Function
		End If

		iStrSelectList = " dbo.ufn_s_GetItemSalesPlanningPrice( " & FilterVar(iStrSoldToParty, "''", "S") & "," _
																  & FilterVar(iStrItemCd, "''", "S") & ", " & FilterVar("*", "''", "S") & " , " & FilterVar("*", "''", "S") & " , " _
																  & FilterVar(iStrUnit, "''", "S")	& ", " _
																  & FilterVar(iStrCur, "''", "S")	& ", '" _
																  & UNIConvDate(iStrFromDt)	& "', " _
																  & FilterVar(lgStrPriceRule , "''", "S") & ")"
		iStrFromList  = ""
		iStrWhereList = ""
		
		Err.Clear
	    
		'품목정보 Fetch
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrPrice = Split(iStrRs, Chr(11))
			
			.Col = C_Price
			iDblOldPrice = UNICDbl(.text)
			.Text = UNIConvNumPCToCompanyByCurrency(iArrPrice(1), iStrCur, Parent.ggUnitCostNo, "X" , "X")
			
			' 금액 재계산 
			If iDblOldPrice <> Cdbl(iArrPrice(1)) Then
				Call CalcAmt(pvLngRow, C_Price)
			End If
			
			GetItemPrice = True
			Exit Function
		Else
			If Err.number <> 0 Then
				MsgBox Err.description 
				Err.Clear
				Exit Function
			End If
		End If
	End With

End Function

'===========================================================================================================
' Description : 수량, 단가 변경시 금액을 재계산한다.
'===========================================================================================================
Sub CalcAmt(ByVal pvLngRow, ByVal pvLngCol)
	Dim iStrCur, iStrNewAmt
	Dim iDblQty, iDblOldQty, iDblPrice, iDblOldAmt, iDblNewAmt
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_Qty	:	iDblQty = UNICDbl(.Text)
		.Col = C_OldQty	:	iDblOldQty = UNICDbl(.Text)
		.Col = C_Price	:	iDblPrice = UNICDbl(.Text)
		
		If pvLngCol = C_Qty And iDblPrice = 0 Then
			If iDblOldQty <> iDblQty Then
				Call ReCalcSpread2(pvLngRow, iDblQty - iDblOldQty, 0, 0, "")
			End If
			Exit Sub
		End If

		If pvLngCol = C_Price And iDblQty = 0 Then Exit Sub
		
		.Col = C_Cur	:	iStrCur = .Text
		.Col = C_Amt
		iDblOldAmt = UNICDbl(.Text)
		iDblNewAmt = iDblQty * iDblPrice
		
		iStrNewAmt = UNIConvNumPCToCompanyByCurrency(iDblNewAmt,iStrCur,Parent.ggAmtOfMoneyNo, "X" , "X")
		.Text = iStrNewAmt
		
		If (iDblOldAmt <> iDblNewAmt) Or (iDblOldQty <> iDblQty) Then
			Call ReCalcSpread2(pvLngRow, iDblQty - iDblOldQty,iDblNewAmt - iDblOldAmt, 0, "")
			.Col = C_OldAmt	: .Text = iStrNewAmt
		End If
	End With

End Sub

'===========================================================================================================
' Description : 집계 Spread 금액 재계산 
'===========================================================================================================
Sub ReCalcSpread2BySpPeriod(ByVal pvLngRow)
	Dim iStrNewSpPeriod, iStrOldSpPeriod
	Dim	iDblQty, lDblAmt
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_SpPeriod		:	iStrNewSpPeriod = .Text
		.Col = C_OldSpPeriod	:	iStrOldSpPeriod = .Text
		If iStrNewSpPeriod <> iStrOldSpPeriod Then
			.Col = C_Qty	:	iDblQty = UNICDbl(.Text)
			.Col = C_Amt	:	lDblAmt = UNICDbl(.Text)
			.Col = C_Pointer
			If .Text <> "" Then
				Call ReCalcSpread2(pvLngRow, -iDblQty, -lDblAmt, C_SpPeriod, iStrOldSpPeriod)
				.Text = ""
			End If
			Call ReCalcSpread2(pvLngRow, iDblQty, lDblAmt, C_SpPeriod, iStrNewSpPeriod)
			.Col = C_OldSpPeriod	: .Text = iStrNewSpPeriod
		End If
	End With
	
End Sub

'===========================================================================================================
' Description : 집계 Spread 금액 재계산 
'===========================================================================================================
Sub ReCalcSpread2BySoldToParty(ByVal pvLngRow)
	Dim iStrNewSoldToParty, iStrOldSoldToParty
	Dim	iDblQty, lDblAmt
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_SoldToParty	:	iStrNewSoldToParty = .Text
		.Col = C_OldSoldToParty	:	iStrOldSoldToParty = .Text
		If iStrNewSoldToParty <> iStrOldSoldToParty Then
			.Col = C_Qty	:	iDblQty = UNICDbl(.Text)
			.Col = C_Amt	:	lDblAmt = UNICDbl(.Text)
			.Col = C_Pointer
			If .Text <> "" Then
				Call ReCalcSpread2(pvLngRow, -iDblQty, -lDblAmt, C_SoldToParty, iStrOldSoldToParty)
				.Text = ""
			End If
			Call ReCalcSpread2(pvLngRow, iDblQty, lDblAmt, C_SoldToParty, iStrNewSoldToParty)
			.Col = C_OldSoldToParty	: .Text = iStrNewSoldToParty
		End If
	End With
	
End Sub

'===========================================================================================================
' Description : 집계 Spread 금액 재계산 
'===========================================================================================================
Sub ReCalcSpread2(ByVal pvLngRow, ByVal pvDblQty, ByVal pvDblAmt, ByVal pvLngCol, ByVal pvStrData)
	Dim iStrPointer, iStrCur, iStrSpPeriod, iStrSpPeriod2, iStrSoldToParty, iStrSoldToParty2
	Dim iLngRow
	Dim iBlnFound
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_Pointer	: iStrPointer = Trim(.Text)
		
		If iStrPointer <> "" Then
			With frm1.vspdData2
				.Row = CLng(iStrPointer)
				.Col = C_Cur2	: iStrCur = .Text
				
				.Col = C_TotQty
				.Text = UNIFormatNumber(UNICDbl(.Text) + pvDblQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)	
				.Col = C_TotAmt
				.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(.Text) + pvDblAmt,iStrCur,Parent.ggAmtOfMoneyNo, "X" , "X")
			End With
		Else
			iBlnFound = False
			If pvLngCol = C_SpPeriod Then
				iStrSpPeriod = pvStrData
			Else
				.Col = C_SpPeriod		:	iStrSpPeriod = .Text
			End If
			
			If pvLngCol = C_SoldToParty Then
				iStrSoldToParty = pvStrData
			Else
				.Col = C_SoldToParty	:	iStrSoldToParty = .Text
			End If
			
			If iStrSpPeriod = "" Or iStrSoldToParty = "" Then Exit Sub

			With frm1.vspdData2
				For iLngRow = 1 To .MaxRows
					.Row = iLngRow
					.Col = C_SpPeriod2		: 	iStrSpPeriod2 = .Text
					.Col = C_SoldToParty2	:	iStrSoldToParty2 = .Text
					
					If iStrSpPeriod = iStrSpPeriod2 And iStrSoldToParty = iStrSoldToParty2 Then
						iBlnFound = True
						Exit For
					End If
				Next
				
				If iBlnFound Then
					.Col = C_Cur2	: iStrCur = .Text
				
					.Col = C_TotQty
					.Text = UNIFormatNumber(UNICDbl(.Text) + pvDblQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)	
					
					.Col = C_TotAmt
					.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(.Text) + pvDblAmt,iStrCur,Parent.ggAmtOfMoneyNo, "X" , "X")
					
					.Col = .MaxCols
					iStrPointer = CStr(CLng(.Text) + 1)
				Else
					.MaxRows = .MaxRows + 1
					.Row = .MaxRows
					.Col = C_SpPeriod2		:	.Text = iStrSpPeriod
					.Col = C_SpPeriodDesc2
					frm1.vspdData.Col = C_SpPeriodDesc	: .Text = frm1.vspdData.Text
					
					.Col = C_SoldToParty2	:	.Text = iStrSoldToParty
					.Col = C_SoldToPartyNm2
					frm1.vspdData.Col = C_SoldToPartyNm	: .Text = frm1.vspdData.Text
					
					.Col = C_TotQty
					frm1.vspdData.Col = C_Qty	: .Text = frm1.vspdData.Text
					.Col = C_Cur2
					frm1.vspdData.Col = C_Cur	: .Text = frm1.vspdData.Text
					.Col = C_TotAmt
					frm1.vspdData.Col = C_Amt	: .Text = frm1.vspdData.Text
					Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData2, .Row, .Row,C_Cur2,C_TotAmt,"A" ,"I","X","X")
					.Col = .MaxCols
					.Text = CStr(.MaxRows - 1)
					iStrPointer = CStr(.MaxRows)
				End If
			End With
			
			' Set the Pointer
			.Col = C_Pointer
			.Text = iStrPointer
		End If
	End With
End Sub

<%'===========================================================================================================
'	Description : 코드값에 해당하는 명을 Display한다.
'====================================================================================================
%>
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(5), iArrTemp
	
	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		If err.number = 0 Then
			If lgBlnOpenedFlag Then
				GetCodeName = OpenConPopup(pvIntWhere)
			End If
		Else
			MsgBox err.Description
		End If
	End if
End Function

' 화폐별로 Cell Formating을 재설정한다.
Sub FormatSpreadCellByCurrency(ByVal pvLngStartRow, ByVal pvLngEndRow, ByVal pvStrEditMode)
	Dim iLngPointer
	Dim iStrCur
	
	' 입력인 경우 
	If pvStrEditMode = "I" Then
		Call FixDecimalPlaceByCurrency(frm1.vspdData,pvLngStartRow,C_Cur,C_Price,"C" ,"X","X")				
		Call FixDecimalPlaceByCurrency(frm1.vspdData,pvLngStartRow,C_Cur,C_Amt,"A" ,"X","X")
		Call FixDecimalPlaceByCurrency(frm1.vspdData,pvLngStartRow,C_Cur,C_AmtLoc,"A" ,"X","X")
		
		With frm1.vspdData
			.Col = C_cur	: iStrCur = .Text
			.Col = C_Pointer: iLngPointer = .Text
		End With
		
		frm1.vspdData2.Col = C_Cur2	: frm1.vspdData2.Text = iStrCur
		If iLngPointer <> "" Then	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData2, iLngPointer, iLngPointer,C_Cur2,C_TotAmt,"A" ,"I","X","X")
	End If
	
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,pvLngStartRow,pvLngEndRow,C_Cur,C_Price,"C" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,pvLngStartRow,pvLngEndRow,C_Cur,C_Amt,"A" ,"I","X","X")         
	Call ReFormatSpreadCellByCellByCurrency(frm1.vspdData,pvLngStartRow,pvLngEndRow,C_Cur,C_AmtLoc,"A" ,"I","X","X")         
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>고객별품목판매계획등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
							</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>판매계획유형</TD>
									<TD CLASS="TD6"><SELECT Name="cboConSpType" ALT="판매계획유형" tag="12XXXU"></SELECT></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp)">&nbsp;<INPUT NAME="txtConSalesGrpNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5">거래구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboConLocExpFlag" tag="12X" STYLE="WIDTH: 150px;"></SELECT></TD>									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>계획기간</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConFrSPPeriod" ALT="계획기간" TYPE="Text" MAXLENGTH=8 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConFrSPPeriod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopFrSpPeriod)">&nbsp;<INPUT NAME="txtConFrSPPeriodDesc" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>고객</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConSoldToParty" ALT="고객" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSoldToParty)">&nbsp;<INPUT NAME="txtConSoldToPartyNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6">&nbsp;~&nbsp;<INPUT NAME="txtConToSPPeriod" ALT="계획기간" TYPE="Text" MAXLENGTH=8 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConToSPPeriod" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopToSpPeriod)">&nbsp;<INPUT NAME="txtConToSPPeriodDesc" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtConItemCd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopItemCd)">&nbsp;<INPUT NAME="txtConItemNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>판매계획유형</TD>
								<TD CLASS="TD6"><SELECT Name="cboSpType" ALT="판매계획유형" tag="24XXXU"><OPTION Value=""></OPTION></SELECT></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>영업그룹</TD>
								<TD CLASS="TD6"><INPUT NAME="txtSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" Disabled=True align=top TYPE="BUTTON">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="24"></TD>
								<TD CLASS="TD5">거래구분</TD>
								<TD CLASS="TD6"><SELECT NAME="cboLocExpFlag" tag="24X" STYLE="WIDTH: 150px;"><OPTION value=""></OPTION></SELECT></TD>									
							</TR>
							<TR>
								<TD HEIGHT="68%" WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT3> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10></TD>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* Align=RIGHT><a href = "VBSCRIPT:JumpChgCheck(BIZ_JUMP_ID)">판매계획확정</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" src="../../blank.htm"  HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA class=hidden name=txtSpreadIns tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadUpd tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpreadDel tag="24" TABINDEX="-1"></TEXTAREA>

<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows2" tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


