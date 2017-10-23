<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : 구매																		*
'*  2. Function Name        : L/C관리																	*
'*  3. Program ID           : M3212PA1																	*
'*  4. Program Name         : L/C 내역팝업																*
'*  5. Program Desc         : 수입진행현황조회를 위한 L/C 내역 팝업 *
'*  7. Modified date(First) : 2003/06/27																*
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
<TITLE>L/C내역팝업</TITLE>
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

Const BIZ_PGM_ID 		= "m3212pb1.asp"                              '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgPopUpR                       
Dim IsOpenPop  

Dim arrReturn
Dim arrParam	
Dim arrParent

'--------------
Dim C_LcSeq		
Dim C_ItemCd 			
Dim C_ItemNm 
Dim C_Spec			
Dim C_Unit	 		
Dim C_LcQty 		
Dim C_Price 		
Dim C_DocAmt	 	
Dim C_PORemainQty 	
Dim C_HsCd 			
Dim C_HsNm 			
Dim C_PoNo 			
Dim C_PoSeq 		
Dim C_OverTolerance 
Dim C_UnderTolerance
Dim C_TrackingNo
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
		.txtBeneficiaryCd.value = arrParam(0)
		.txtBeneficiaryNm.value = arrParam(1)
		.txtLCDocNo.value 		= arrParam(2)
		.txtLCAmendSeq.value 	= arrParam(3)
		.txtLCNo.Value			= arrParam(4)
		.txtCurrency.Value		= arrParam(5)
		'.vspdData.OperationMode = 3
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
	C_LcSeq 		 = 1
	C_ItemCd 		 = 2	
	C_ItemNm 		 = 3	
	C_Spec			 = 4
	C_Unit	 		 = 5
	C_LcQty 	     = 6	
	C_Price 		 = 7
	C_DocAmt	 	 = 8
	C_PORemainQty 	 = 9
	C_HsCd 			 = 10
	C_HsNm 			 = 11
	C_PoNo 			 = 12
	C_PoSeq 		 = 13
	C_OverTolerance  = 14
	C_UnderTolerance = 15
	C_TrackingNo	 = 16

End Sub
'=======================================  GetSpreadColumnPos()  ========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_LcSeq 		 = iCurColumnPos(1)
			C_ItemCd 		 = iCurColumnPos(2)	
			C_ItemNm 		 = iCurColumnPos(3)	
			C_Spec			 = iCurColumnPos(4)	
			C_Unit	 		 = iCurColumnPos(5)
			C_LcQty 	     = iCurColumnPos(6)	
			C_Price 		 = iCurColumnPos(7)
			C_DocAmt	 	 = iCurColumnPos(8)
			C_PORemainQty 	 = iCurColumnPos(9)
			C_HsCd 			 = iCurColumnPos(10)
			C_HsNm 			 = iCurColumnPos(11)
			C_PoNo 			 = iCurColumnPos(12)
			C_PoSeq 		 = iCurColumnPos(13)
			C_OverTolerance  = iCurColumnPos(14)
			C_UnderTolerance = iCurColumnPos(15)
			C_TrackingNo	 = iCurColumnPos(16)	
	End Select
End Sub
<!--
'==========================================  2.2.3 InitSpreadSheet()  ===================================
-->
Sub InitSpreadSheet()
    With frm1
		Call InitSpreadPosVariables()

		ggoSpread.Source = .vspdData
		ggoSpread.SpreadInit "V20030627",,PopupParent.gAllowDragDropSpread
			
		.vspdData.ReDraw = False

		.vspdData.MaxCols = C_TrackingNo + 1
		.vspdData.MaxRows = 0
		
		Call AppendNumberPlace("6", "5", "0")	
		Call GetSpreadColumnPos("A")
			
		ggoSpread.SSSetFloat 	C_LcSeq,		"L/C순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		ggoSpread.SSSetEdit		C_ItemCd,		"품목", 18, 0
		ggoSpread.SSSetEdit		C_ItemNm,		"품목명", 20, 0
		ggoSpread.SSSetEdit		C_Spec,			"품목규격", 20, 0
		ggoSpread.SSSetEdit		C_Unit,			"단위", 10, 2
		ggoSpread.SSSetFloat	C_LcQty,		"L/C수량", 9, PopupParent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, 1,,"Z"
		ggoSpread.SSSetFloat 	C_Price,		"단가", 15,"C" ,ggStrIntegeralPart ,ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec		
		ggoSpread.SSSetFloat 	C_DocAmt,		"금액", 15,"A" ,ggStrIntegeralPart ,ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
		ggoSpread.SSSetFloat	C_PORemainQty,	"발주잔량", 10, PopupParent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, 1,,"Z"
		ggoSpread.SSSetEdit		C_HsCd,			"HS부호", 20, 0
		ggoSpread.SSSetEdit		C_HsNm,			"HS명", 20, 0
		ggoSpread.SSSetEdit		C_PoNo,			"발주번호", 18, 0
		ggoSpread.SSSetFloat 	C_PoSeq,		"발주순번",10,"6",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,0
		ggoSpread.SSSetFloat	C_OverTolerance, "과부족허용율(+)", 15, "D"       ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, 1,,"Z"
		ggoSpread.SSSetFloat	C_UnderTolerance,"과부족허용율(-)", 15, "D"       ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, 1,,"Z"
		ggoSpread.SSSetEdit		C_TrackingNo, "Tracking No.",  25,,,25,2

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
'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	ggoOper.FormatFieldByObjectOfCur frm1.txtDocAmt, frm1.txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
	ggoOper.FormatFieldByObjectOfCur frm1.txtTotItemAmt, frm1.txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec

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
			strVal = strVal & "&txtLCNo=" & Trim(.txtLCNo.value)			'L/C관리번호 
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001		'☜: 비지니스 처리 ASP의 상태 
			strVal = strVal & "&txtLCNo=" & Trim(.txtLCNo.value)			'L/C관리번호 
		End If
		strVal = strVal		& "&txtCurrency=" & Trim(.txtCurrency.value)
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
		frm1.txtLCDocNo.focus
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
							<TD CLASS="TD5" NOWRAP>L/C 관리번호</TD>
							<TD CLASS="TD6"><INPUT NAME="txtLCNo" ALT="L/C 관리번호" TYPE="Text" SIZE=30 MAXLENGTH=18  TAG="14XXXU"></TD>
							<TD CLASS=TD5>L/C번호</TD>
							<TD CLASS=TD6><INPUT NAME="txtLCDocNo" ALT="L/C번호" TYPE=TEXT MAXLENGTH=18 SIZE=20  TAG="14XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>수출자</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiaryCd" SIZE=10  MAXLENGTH=10 TAG="24XXXU" ALT="수출자">&nbsp;
												 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
							<TD CLASS=TD5 NOWRAP>개설일</TD>
							<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m3212pa1_fpDateTime1_txtOpenDt.js'></script></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>총개설금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<Table cellpadding=0 cellspacing=0>
									<TR>
										<TD NOWRAP>
											<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10  MAXLENGTH=3 TAG="24XXXU" ALT="통화">&nbsp;&nbsp;
										</TD>
										<TD NOWRAP>
											<script language =javascript src='./js/m3212pa1_fpDoubleSingle5_txtDocAmt.js'></script></TD>
										</TD>
									</TR>
								</Table>
							</TD>
							<TD CLASS=TD5 NOWRAP>총품목금액</TD>
							<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/m3212pa1_fpDoubleSingle5_txtTotItemAmt.js'></script></TD>
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
						<script language =javascript src='./js/m3212pa1_vaSpread1_vspdData.js'></script>
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

