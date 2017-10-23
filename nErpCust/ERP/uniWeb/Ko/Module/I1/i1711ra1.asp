<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : I Goods Movement Detail List
'*  2. Function Name        : 
'*  3. Program ID           : I1711ra1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 수불현황 상세 조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/05/14
'*  8. Modified date(Last)  : 2003/06/03
'*  9. Modifier (First)     : Hae Ryong Lee
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_ID = "i1711rb1.asp"

DIm C_ItemCode 
DIm C_ItemName 
DIm C_BaseUnit 
DIm C_Qty      
DIm C_Price    
DIm C_Amount   
DIm C_Storage  
DIm C_BizCd    

Dim arrReturn
Dim arrParam

Dim arrBizCd
Dim arrBizNm
Dim arrItemDocumentNo
Dim arrDocumentDt
Dim arrPosDt
Dim arrTrnsType
Dim arrMovType
  
arrParam        = window.dialogArguments
Set PopupParent = arrParam(0)

arrBizCd          = arrParam(1)
arrBizNm          = arrParam(2)
arrItemDocumentNo = arrParam(3)
arrDocumentDt     = arrParam(4)
arrPosDt          = arrParam(5)
arrTrnsType       = arrParam(6)
arrMovType        = arrParam(7)

top.document.title = PopupParent.gActivePRAspName

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgStrPrevSubKey
Dim lgUserFlag   

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntGrpCount		= 0                      
    lgLngCurRows		= 0                      
    Self.Returnvalue	= Array("")
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
 txtBizCd.value           = arrBizCd 
 txtBizNm.value           = arrBizNm 
 txtItemDocumentNo.value  = arrItemDocumentNo
 txtDocumentDt.Text       = arrDocumentDt
 txtPosDt.text			  = arrPosDt 
 txtTrnsType.value        = arrTrnsType 
 txtMovType.value         = arrMovType 
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I","NOCOOKIE","RA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20021106", , PopupParent.gAllowDragDropSpread
	
	With  vspdData
		.ReDraw = false
		.MaxCols = C_BizCd+1   
		.MaxRows = 0

		Call GetSpreadColumnPos("A")
				     
		ggoSpread.SSSetEdit		C_ItemCode,	"품목",		18
		ggoSpread.SSSetEdit		C_ItemName, "품목명",	40
		ggoSpread.SSSetEdit		C_BaseUnit, "단위",		6,2 
		ggoSpread.SSSetFloat	C_Qty,		"수불수량", 15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat	C_Price,	"수불단가", 15, PopupParent.ggUnitCostNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloat	C_Amount,	"수불금액", 15, PopupParent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetEdit		C_Storage,	"창고",		30 
		ggoSpread.SSSetEdit		C_BizCd,	"사업장",	10
		    
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetSplit2(2)
		.ReDraw = true
		
	End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_ItemCode = 1
	C_ItemName = 2
	C_BaseUnit = 3
	C_Qty      = 4
	C_Price    = 5
	C_Amount   = 6
	C_Storage  = 7
	C_BizCd    = 8
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_ItemCode = iCurColumnPos(1)
		C_ItemName = iCurColumnPos(2)
		C_BaseUnit = iCurColumnPos(3)
		C_Qty      = iCurColumnPos(4)
		C_Price    = iCurColumnPos(5)
		C_Amount   = iCurColumnPos(6)
		C_Storage  = iCurColumnPos(7)
		C_BizCd    = iCurColumnPos(8)		
	End Select
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim intColCnt
  
	If vspdData.ActiveRow > 0 Then 
		Redim arrReturn(vspdData.MaxCols - 1)
  		vspdData.Row = vspdData.ActiveRow
     
		For intColCnt = 0 To vspdData.MaxCols - 1
			vspdData.Col = intColCnt + 1
			arrReturn(intColCnt) = vspdData.Text
		Next
		Self.Returnvalue = arrReturn
	End If
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Self.Close()
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)                                                  '⊙: Setup the Spread sheet
    Call ggoOper.LockField(Document, "N")
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    

    Call InitSpreadSheet
    Call InitVariables 
    Call SetDefaultVal()
    Call DbQuery()
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = vspdData
   
	If vspdData.MaxRows = 0 Then Exit Sub
	If Row <= 0 Then
		ggoSpread.Source = vspdData 
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	
			lgSortKey = 1
		End If
		Exit Sub
	End If
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
   
	If Row <= 0 Then Exit Sub
	If vspdData.MaxRows = 0 Then Exit Sub

	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If	
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   ggoSpread.Source = vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 


Function vspdData_KeyPress(KeyAscii)
  On Error Resume Next
  If KeyAscii = 27 Then
     Call CancelClick()
  End IF 
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    If  vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) and lgStrPrevKey <> "" and lgStrPrevSubKey <> "" Then       '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		DbQuery
	End If
End Sub

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 

    FncQuery = False       
    Err.Clear              

    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables
    
    If DbQuery() = False Then Exit Function
       
    FncQuery = True  
    
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    Dim strVal    
    Dim strYear
    Dim strMonth
    Dim strDay

    DbQuery = False
    
    Call LayerShowHide(1)

    Call ExtractDateFrom(txtDocumentDt.Text,txtDocumentDt.UserDefinedFormat,PopupParent.gComDateType,strYear,strMonth,strDay)

	Call CommonQueryRs(" MINOR_CD ", " B_MINOR ", _
			" MAJOR_CD =" & FilterVar("I0002", "''", "S") & " AND MINOR_NM = " & FilterVar(txtTrnsType.value, "''", "S") & " ORDER BY MINOR_NM ",_
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    lgF0 = Split(lgF0, Chr(11))
	                                                     
    Err.Clear                                            
    
    strVal = BIZ_PGM_ID &	"?txtItemDocumentNo=" & Trim(txtItemdocumentNo.value)	& _
							"&txtDocumentYear="   & strYear							& _
							"&txtTrnsType="       & lgF0(0)							& _
							"&lgStrPrevKey="      & lgStrPrevKey					& _
							"&lgStrPrevSubKey="   & lgStrPrevSubKey					& _
							"&txtMaxRows="        & vspdData.MaxRows

    Call RunMyBizASP(MyBizASP, strVal)
  
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    txtTotAmount.text	= FncSumSheet(vspdData, C_Amount, 1, vspdData.MaxRows, False, 0, 0, "V")
    txtTotQty.text		= FncSumSheet(vspdData, C_Qty, 1, vspdData.MaxRows, False, 0, 0, "V")
    vspdData.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
	<TABLE <%=LR_SPACE_TYPE_00%>>
		<TR>
			<TD HEIGHT=40>
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> >
						</TD>
					</TR>
					<TR>
						<TD HEIGHT=20>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%> > 
									<TR>
										<TD CLASS="TD5" NOWRAP>사업장</TD>
										<TD CLASS="TD6" NOWRAP><input NAME="txtBizCd" TYPE="Text" MAXLENGTH="10" tag="14XXXU" ALT = "사업장" size=10>&nbsp;<input NAME="txtBizNm" TYPE="Text" MAXLENGTH="40" tag="14NXXU"></TD>
										<TD CLASS="TD5" NOWRAP>수불번호</TD>
										<TD CLASS="TD6" NOWRAP><input NAME="txtItemdocumentNo" TYPE="Text" MAXLENGTH="18" tag="14XXXU" ALT = "수불번호" size=18></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>수불발생일</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/i1711ra1_I953062736_txtDocumentDt.js'></script></TD>
										<TD CLASS="TD5" NOWRAP>회계전표발생일</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/i1711ra1_I686070662_txtPosDt.js'></script></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>수불구분</TD>
										<TD CLASS="TD6" NOWRAP><input NAME="txtTrnsType" TYPE="Text" MAXLENGTH="14" tag="14XXXU" ALT = "수불구분" size=14></TD>
										<TD CLASS="TD5" NOWRAP>수불유형</TD>
										<TD CLASS="TD6" NOWRAP><input NAME="txtMovType" TYPE="Text" MAXLENGTH="30" tag="14XXXU" ALT = "수불유형" size=30></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>총수불수량</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/i1711ra1_fpDoubleSingle1_txtTotQty.js'></script></TD>
										<TD CLASS="TD5" NOWRAP>총수불금액</TD>
										<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/i1711ra1_fpDoubleSingle1_txtTotAmount.js'></script></TD>       
									</TR>
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100%> 
						</TD>
					</TR>
				</TABLE>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD>
							<script language =javascript src='./js/i1711ra1_I666752696_vspdData.js'></script>
							</TD>
						</TR>  
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_01%> >
				</TD>
			</TR>
			<TR HEIGHT=20>
				<TD WIDTH=100%>
					<TABLE <%=LR_SPACE_TYPE_30%>>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD WIDTH=70% NOWRAP></TD>
							<TD WIDTH=30% ALIGN=RIGHT>
								<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
							<TD WIDTH=10>&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
				</TD>
			</TR>
		</TABLE>
		<DIV ID="MousePT" NAME="MousePT">
			<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
		</DIV>
	</BODY>
</HTML>

