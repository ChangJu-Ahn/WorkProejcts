<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Prucurement																		*
'*  2. Function Name        : 																	*
'*  3. Program ID           : m4114pa2																	*
'*  4. Program Name         : 월별매입가계정현황-반제금액(IR) 팝업																*
'*  5. Program Desc         :  																			*
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/10/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

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

Const BIZ_PGM_ID 		= "m4114pb2.asp"                              '☆: Biz Logic ASP Name

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgPopUpR                       
Dim IsOpenPop  

Dim arrReturn
Dim arrParam	
Dim arrParent

'--------------
Dim C_IV_TYPE_CD	'매입유형 
Dim C_IV_TYPE_NM	'매입유형명 
Dim C_IV_LOC_AMT	'반제금액(IR)
'---------------

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'==========================================  2.1.1 InitVariables()  =====================================
Function InitVariables()
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
		.txtSearchDt.text 		= arrParam(0)&"01"
		.txtBpCd.value 			= arrParam(1)
		.txtBpNm.Value			= arrParam(2)
	End With

	ggoOper.SetReqAttr	frm1.txtSearchDt, "Q"	
	
		
End Sub

'=================================  LoadInfTB19029()  ======================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "PA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "PA") %>
End Sub
'=======================================  initSpreadPosVariables()  ========================================
Sub InitSpreadPosVariables() 
	C_IV_TYPE_CD	= 1		'매입유형 
	C_IV_TYPE_NM	= 2		'매입유형명 
	C_IV_LOC_AMT		= 3		'반제금액(IR)
End Sub
'=======================================  GetSpreadColumnPos()  ========================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_IV_TYPE_CD	= iCurColumnPos(1)		'매입유형 
			C_IV_TYPE_NM	= iCurColumnPos(2)		'매입유형명 
			C_IV_LOC_AMT		= iCurColumnPos(3)		'반제금액(IR)
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

		.vspdData.MaxCols = C_IV_LOC_AMT + 1
		.vspdData.MaxRows = 0
		
		Call AppendNumberPlace("6", "5", "0")	
		Call GetSpreadColumnPos("A")

	
		ggoSpread.SSSetEdit		C_IV_TYPE_CD,		"매입유형", 18, 0
		ggoSpread.SSSetEdit		C_IV_TYPE_NM,		"매입유형명", 20, 0
		SetSpreadFloatLocal 	C_IV_LOC_AMT,		"반제금액(IR)",15,1,2

	    Call ggoSpread.SSSetColHidden(.vspdData.MaxCols, .vspdData.MaxCols, True)
		.vspdData.ReDraw = True

		Call SetSpreadLock()
	End With
End Sub

'=========================  SetSpreadFloatLocal() ==================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	     
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, HAlign,,"Z"
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, PopupParent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, HAlign,,"Z"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, PopupParent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, HAlign,,"Z"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, PopupParent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec, HAlign,,"Z"
        Case 6                                                              '과부족허용율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, "6" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,,"0","999"
    End Select
         
End Sub


'=================================  SetSpreadLock()  ======================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'===================================== CurFormatNumericOCX()  =======================================
Sub CurFormatNumericOCX()

	ggoOper.FormatFieldByObjectOfCur frm1.txtDocAmt, frm1.txtCurrency.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec

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
	Call ggoOper.FormatDate(frm1.txtSearchDt,PopupParent.gDateFormat,"2")

	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	If DbQuery = False Then
		Exit Sub
	End if
	
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

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
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
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
	Dim strYear,strMonth,strSearchDT
	
	'2006.11.22 검색년월 세팅 변경 KSJ
	strYear = Trim(frm1.txtSearchDt.Year)
	strMonth = Trim(frm1.txtSearchDt.Month)
	
	if Cint(strMonth) < 10 Then strMonth =  "0" & strMonth
	strSearchDT = strYear & strMonth
		
	With frm1

		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001		'☜: 비지니스 처리 ASP의 상태 
		'strVal = strVal & "&txtSearchDt=" & Trim(frm1.txtSearchDt.Year) & Trim(frm1.txtSearchDt.Month)
		strVal = strVal & "&txtSearchDt=" & strSearchDT
		strVal = strVal & "&txtBpCd=" & Trim(frm1.txtBpCd.value)
		 
	End With
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
							<TD CLASS="TD5" NOWRAP>조회년월</TD>
							<TD CLASS="TD6" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="조회년월" NAME="txtSearchDt" CLASS=FPDTYYYYMM CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 tag="11N" Title="FPDATETIME" STYLE="READONLY:TRUE;"></OBJECT>');</SCRIPT>
							</TD>
							<TD CLASS="TD5" NOWRAP>공급처</TD>
							<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT ALT="공급처" NAME="txtBpCd"  SIZE=10 MAXLENGTH=10 tag="14">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14">
							</TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
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

<INPUT TYPE=HIDDEN NAME="hdnCurrency" TAG="14">
<INPUT TYPE=HIDDEN NAME="hdnBlNo" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHBLDocNo" TAG="14">
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

