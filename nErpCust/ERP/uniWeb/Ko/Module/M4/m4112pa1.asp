<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : 구매																		*
'*  2. Function Name        : L/C관리																	*
'*  3. Program ID           : M4112PA1																	*
'*  4. Program Name         : 입고내역팝업																*
'*  5. Program Desc         : 수입진행현황조회를 위한 입고내역팝업 *
'*  7. Modified date(First) : 2003/07/01																*
'*  8. Modified date(Last)  :           																*
'*  9. Modifier (First)     : Lee Eun hee																*
'* 10. Modifier (Last)      :           
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 												*
'*				            : 												*
'*				            : 												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE>입고내역팝업</TITLE>
<!--
'******************************************  1.1 Inc 선언   **********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

Const BIZ_PGM_ID 		= "m4112pb1.asp"                              '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgPopUpR                       
Dim IsOpenPop  

Dim arrReturn
Dim arrParam	
Dim arrParent

'--------------
Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec   
Dim C_TrackingNo
Dim C_InspFlg	
Dim C_GrQty	
Dim C_StockQty
Dim C_GRUnit
Dim C_Cur	
Dim C_MvmtPrc
Dim C_DocAmt
Dim C_LocAmt
Dim C_SlCd	
Dim C_SlNm		
Dim C_InspSts	
Dim C_GRMeth	
Dim C_LotNo		
Dim C_LotSeqNo	
Dim C_MakerLotNo
Dim C_MakerLotSeqNo
Dim C_GRNo
Dim C_GRSeqNo
Dim C_InspReqNo
Dim C_InspResultNo
Dim C_PoNo
Dim C_PoSeqNo
Dim C_CCNo
Dim C_CCSeqNo
Dim C_LLCNo
Dim C_LLCSeqNo

'---------------

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
    lgStrPrevKeyIndex	= ""
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
			
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
	Dim arrParam
		
	arrParam = arrParent(1)

	With frm1
		.txtSupplierCd.value = arrParam(0)
		.txtMvmtNo.value 	 = arrParam(1)
	End With
		
End Sub

'=================================  LoadInfTB19029()  ======================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "PA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "PA") %>
End Sub
'=======================================  initSpreadPosVariables()  ========================================
Sub InitSpreadPosVariables() 
	C_PlantCd		= 1
	C_PlantNm		= 2
	C_ItemCd		= 3
	C_ItemNm		= 4
	C_Spec			= 5 
	C_TrackingNo	= 6
	C_InspFlg		= 7

	C_GrQty			= 8
	C_StockQty		= 9
	C_GRUnit		= 10
	C_Cur		    = 11
	C_MvmtPrc	    = 12
	C_DocAmt		= 13
	C_LocAmt		= 14
	C_SlCd			= 15

	C_SlNm			= 16
	C_InspSts		= 17
	C_GRMeth		= 18
	C_LotNo			= 19
	C_LotSeqNo		= 20
	C_MakerLotNo	= 21
	C_MakerLotSeqNo	= 22
	C_GRNo			= 23
	C_GRSeqNo		= 24
	C_InspReqNo		= 25
	C_InspResultNo	= 26
	C_PoNo			= 27
	C_PoSeqNo		= 28
	C_CCNo			= 29
	C_CCSeqNo		= 30
	C_LLCNo			= 31
	C_LLCSeqNo		= 32

End Sub
'=======================================  GetSpreadColumnPos()  ========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_PlantCd		= iCurColumnPos(1)
			C_PlantNm		= iCurColumnPos(2)
			C_ItemCd		= iCurColumnPos(3)
			C_ItemNm		= iCurColumnPos(4)
			C_Spec			= iCurColumnPos(5) 
			C_TrackingNo	= iCurColumnPos(6)
			C_InspFlg		= iCurColumnPos(7)

			C_GrQty			= iCurColumnPos(8)
			C_StockQty		= iCurColumnPos(9)
			C_GRUnit		= iCurColumnPos(10)
			C_Cur		    = iCurColumnPos(11)
			C_MvmtPrc	    = iCurColumnPos(12)
			C_DocAmt		= iCurColumnPos(13)
			C_LocAmt		= iCurColumnPos(14)
			C_SlCd			= iCurColumnPos(15)

			C_SlNm			= iCurColumnPos(16)
			C_InspSts		= iCurColumnPos(17)
			C_GRMeth		= iCurColumnPos(18)
			C_LotNo			= iCurColumnPos(19)
			C_LotSeqNo		= iCurColumnPos(20)
			C_MakerLotNo	= iCurColumnPos(21)
			C_MakerLotSeqNo	= iCurColumnPos(22)
			C_GRNo			= iCurColumnPos(23)
			C_GRSeqNo		= iCurColumnPos(24)
			C_InspReqNo		= iCurColumnPos(25)
			C_InspResultNo  = iCurColumnPos(26)
			C_PoNo			= iCurColumnPos(27)
			C_PoSeqNo		= iCurColumnPos(28)
			C_CCNo			= iCurColumnPos(29)
			C_CCSeqNo		= iCurColumnPos(30)
			C_LLCNo			= iCurColumnPos(31)
			C_LLCSeqNo		= iCurColumnPos(32)

	End Select
End Sub
<!--
'==========================================  2.2.3 InitSpreadSheet()  ===================================
-->
Sub InitSpreadSheet()
    With frm1
		Call InitSpreadPosVariables()

		ggoSpread.Source = .vspdData
		ggoSpread.SpreadInit "V20030701",,PopupParent.gAllowDragDropSpread
			
		.vspdData.ReDraw = False

		.vspdData.MaxCols = C_LLCSeqNo + 1
		.vspdData.MaxRows = 0
		
		Call AppendNumberPlace("6", "5", "0")	
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit 		C_PlantCd,	"공장", 7
		ggoSpread.SSSetEdit 		C_PlantNm,	"공장명", 20
		ggoSpread.SSSetEdit 		C_ItemCd,	"품목", 15
		ggoSpread.SSSetEdit 		C_ItemNm,	"품목명", 20 
		ggoSpread.SSSetEdit 		C_Spec,	    "품목규격", 20 	
		ggoSpread.SSSetEdit 		C_TrackingNo,	"Tracking No.", 15 	
		ggoSpread.SSSetCheck 		C_InspFlg,	"검사품여부",10,,,true
		ggoSpread.SSSetFloat		C_GrQty,	"입고수량", 10, PopupParent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, 1,,"Z"
		ggoSpread.SSSetFloat		C_StockQty,	"재고처리수량", 13, PopupParent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, 1,,"Z"
		ggoSpread.SSSetEdit 		C_GRUnit,	"단위", 7
		ggoSpread.SSSetEdit 		C_Cur,	    "화폐", 7
		
		ggoSpread.SSSetFloat		C_MvmtPrc,	"입고단가"		, 10	,"C" ,ggStrIntegeralPart ,ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
		ggoSpread.SSSetFloat 		C_DocAmt,	"입고금액"		, 15	,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
		ggoSpread.SSSetFloat 		C_LocAmt,	"입고자국금액"	, 15	,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec		    
		
		ggoSpread.SSSetEdit			C_SlCd,		"창고", 8
		ggoSpread.SSSetEdit 		C_SlNm,		"창고명", 20	    
		ggoSpread.SSSetEdit 		C_InspSts,	"검사상태", 10
		ggoSpread.SSSetEdit 		C_GRMeth,	"납입시검사방법", 20
		ggoSpread.SSSetEdit 		C_LotNo,	"Lot No.", 20, , , 12, 2    
		ggoSpread.SSSetFloat 		C_LotSeqNo,	"LOT NO 순번",12,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		
		ggoSpread.SSSetEdit 		C_MakerLotNo,	"MAKER LOT NO.", 20,,,,2    
		ggoSpread.SSSetFloat 		C_MakerLotSeqNo,"Maker Lot 순번",15,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_GRNo,		"재고처리번호", 20
		ggoSpread.SSSetFloat 		C_GRSeqNo,	"재고처리순번",15,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		
		ggoSpread.SSSetEdit 		C_InspReqNo,"검사요청번호", 20
		ggoSpread.SSSetEdit 		C_InspResultNo,"검사결과등록번호", 20
		ggoSpread.SSSetEdit 		C_PoNo,		"발주번호", 20
		ggoSpread.SSSetFloat 		C_PoSeqNo,	"발주순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		ggoSpread.SSSetEdit 		C_CCNo,		"통관번호", 20
		ggoSpread.SSSetFloat 		C_CCSeqNo,	"통관순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0    
		ggoSpread.SSSetEdit 		C_LLCNo,	"LOCAL L/C번호", 20
		ggoSpread.SSSetFloat 		C_LLCSeqNo,	"LOCAL L/C순번",13,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		
		Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
		Call SetSpreadLock()
			
		.vspdData.ReDraw = True
	End With
End Sub
'=================================  SetSpreadLock()  ======================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow
	with frm1
	If .vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(.vspdData.SelModeSelCount - 1, .vspdData.MaxCols - 1)

		For intRowCnt = 0 To .vspdData.MaxRows - 1

			.vspdData.Row = intRowCnt + 1

			If .vspdData.SelModeSelected Then
				For intColCnt = 0 To .vspdData.MaxCols - 1
					'.vspdData.Col = intColCnt + 1
					'arrReturn(intInsRow, intColCnt) = .vspdData.Text
					frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
				Next

				intInsRow = intInsRow + 1

			End IF
		Next
	End If			
	End With
	Self.Returnvalue = arrReturn
	Self.Close()
End Function	

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
	
'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>

	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	If DbQuery = False Then
		Exit Sub
	End if
	
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
'=========================================  3.3.1 vspdData_DblClick()  ==================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
      Exit Function
    End If
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================  3.3.2 vspdData_KeyPress()  ===================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	 gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("1101111111")

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	frm1.vspdData.Row = Row
End Sub
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'======================================  3.3.3 vspdData_TopLeftChange()  ================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If DbQuery = False Then
				Exit Sub
			End if
		End if
	End if
    
    
End Sub

'===================================  FncQuery()  ============================================
Function FncQuery() 
    FncQuery = False                                                 
    Err.Clear                                                        

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    						
	Call InitVariables												

    If DbQuery = False Then Exit Function							

    FncQuery = True									
    Set gActiveElement = document.activeElement    
End Function

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery()
	Err.Clear															<%'☜: Protect system from crashing%>

	DbQuery = False														<%'⊙: Processing is NG%>

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
		
	With frm1

	   If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001		'☜: 비지니스 처리 ASP의 상태 
			strVal = strVal & "&txtMvmtNo=" & Trim(.txtMvmtNo.value)			'L/C관리번호 
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001		'☜: 비지니스 처리 ASP의 상태 
			strVal = strVal & "&txtMvmtNo=" & Trim(.txtMvmtNo.value)			'L/C관리번호 
		End If
	End With
		strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
		strVal = strVal		& "&lgPageNo="       & lgPageNo                  '☜: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

		DbQuery = True														<%'⊙: Processing is NG%>
End Function
'===================================  DbQueryOk()  ============================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtMvmtNo.focus
	End If
	Set gActiveElement = document.activeElement
End Function
'===================================  OpenOrderBy()  ============================================
Function OpenOrderByPopup()
	Dim arrRet
	
	On Error Resume Next
	
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS="TD5" NOWRAP>입고번호</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="입고번호" NAME="txtMvmtNo" MAXLENGTH=18 SIZE=32 tag="14XXXU"></TD>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>입고형태</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="24NXXU">
												   <INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtTypeNm" SIZE=20 tag="24X"></TD>
							<TD CLASS="TD5" NOWRAP>입고일</TD>
							<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m4112pa1_fpDateTime1_txtGmDt.js'></script></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>공급처</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="24XXXU">
												   <INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="24X"></TD>
							<TD CLASS="TD5" NOWRAP>구매그룹</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="24XXXU">
												   <INPUT TYPE=TEXT Alt="구매그룹" ID="txtGroupNm" SIZE=20 NAME="arrCond" tag="24X"></TD>								
						</TR>
							
					</TABLE>
				</FIELDSET>
			</TD>
		</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<script language =javascript src='./js/m4112pa1_vaSpread1_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
											</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

