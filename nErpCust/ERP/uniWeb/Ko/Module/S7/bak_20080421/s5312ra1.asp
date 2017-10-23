<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5312RA1
'*  4. Program Name         : 세금계산서내역 참조 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/05/03
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kwak Eunkyoung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
'*                            2002/05/03 : ADO변환 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>세금계산서내역</TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s5312rb1.asp"                              '☆: Biz Logic ASP Name

' Constant variables 
'========================================
Const C_MaxKey          = 7                                           '☆: key count of SpreadSheet

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->

' User-defind Variables
'========================================
Dim IsOpenPop  

Dim arrReturn										<% '--- Return Parameter Group %>
Dim arrParam	

Dim ArrPopupParent
Dim PopupParent

ArrPopupParent = window.dialogArguments
Set PopupParent  = ArrPopupParent(0)

'2002-12-28 Dynamic Popup
top.document.title = PopupParent.gActivePRAspName

'========================================
Function InitVariables()
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1
			
	ReDim arrReturn(0)
	Self.Returnvalue = arrReturn

End Function

'========================================
Sub SetDefaultVal()

	Dim arrParam
	
	arrParam = ArrPopupParent(1)
	
	With frm1
		.txtTaxBillNo.value = arrParam(0)
		.txtTaxBillDocNo.value = arrParam(1)
		.txtBilltoParty.value = arrParam(2)
		.txtBilltoPartyNm.value = arrParam(3)
			
		If arrParam(4) = "Y" Then
			.rdoPostYes.checked = True
		ElseIf arrParam(4) = "N" Then
			.rdoPostNo.checked = True
		End If
	
		.txtSupplyAmt.text = arrParam(5)
		.txtCurrency.value = arrParam(6)
		.txtVatAmt.text = arrParam(7)
 	End With

End Sub

'========================================
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "*", "NOCOOKIE", "PA") %>
<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "PA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	    
    Call SetZAdoSpreadSheet("s5312ra1","S","A","V20030301", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	
	frm1.vspdData.OperationMode = 5

	Call SetSpreadLock 
	      
End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
'	ggoSpread.SpreadLock 1 , -1
End Sub

'========================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next 
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================
Sub Form_Load()
	Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>

	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

'========================================
	Sub Form_QueryUnload(Cancel, UnloadMode)
	End Sub

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then Call DbQuery
	End If
End Sub

'==========================================
Function FncQuery() 
    
	FncQuery = False                                                 
    
	Err.Clear                                                        

	Call ggoOper.ClearField(Document, "2")							
	Call InitVariables												

	If DbQuery = False Then Exit Function							

	FncQuery = True									
        
End Function

'==========================================
Function DbQuery()
	Err.Clear															<%'☜: Protect system from crashing%>

	DbQuery = False														<%'⊙: Processing is NG%>

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

	With frm1
		
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtTaxBillNo=" & Trim(.txtTaxBillNo.value)				<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)				<%'☆: 조회 조건 데이타 %>

	End With

    strVal =     strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
	strVal =     strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")         
    strVal =     strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal =     strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

	DbQuery = True														<%'⊙: Processing is NG%>
End Function

'========================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtTaxBillNo.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
%>
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
						<TD CLASS=TD5 NOWRAP>세금계산서관리번호</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtTaxBillNo" ALT="세금계산서관리번호" TYPE="Text" MAXLENGTH="18" tag="14XXXU">&nbsp;</TD>
						<TD CLASS=TD5 NOWRAP>세금계산서번호</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT NAME="txtTaxBillDocNo" ALT="세금계산서관리번호" TYPE="Text" MAXLENGTH="18" tag="14XXXU">&nbsp;</TD>
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
					<TD CLASS=TD5 NOWRAP>발행처</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBilltoParty" SIZE=10 MAXLENGTH=10 TAG="14XXXU" ALT="발행처">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBilltoPartyNm" SIZE=20 TAG="14"></TD>
					<TD CLASS=TD5 NOWRAP>공급가액</TD>
					<TD CLASS=TD6 NOWRAP>
						<TABLE CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>
									<script language =javascript src='./js/s5312ra1_fpDoubleSingle2_txtSupplyAmt.js'></script>
								</TD>
								<TD>
									&nbsp;<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="14">
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>발행여부</TD>
					<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE=radio CLASS="RADIO" NAME="rdoPost" id="rdoPostNo" VALUE="N" tag = "14" CHECKED>
							<LABEL FOR="rdoPostNo">미발행</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;
						<INPUT TYPE=radio CLASS = "RADIO" NAME="rdoPost" id="rdoPostYes" VALUE="Y" tag = "14">
							<LABEL FOR="rdoPostYes">발행</LABEL>
					</TD>
					<TD CLASS=TD5 NOWRAP>VAT금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s5312ra1_fpDoubleSingle4_txtVatAmt.js'></script>							
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/s5312ra1_vaSpread_vspdData.js'></script>
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
				<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
											  <IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG></TD>
			        <TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>&nbsp;&nbsp;</TD>
			        <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>