<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ra3.asp																*
'*  4. Program Name         : Local L/C 상세정보(Local L/C현황조회에서)									*
'*  5. Program Desc         : Local L/C 상세정보(Local L/C현황조회에서)									*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/07/12																*
'*  8. Modified date(Last)  : 2002/04/12																*
'*  9. Modifier (First)     : An ChangHwan 																*
'* 10. Modifier (Last)      : Seo Jinkyung															    *
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/07/12 : Coding ReStart											*
'*							  3. 2002/04/12 : ADO 변환													*
'*																										*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>LOCAL L/C 상세정보</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              

<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_ID = "s3211rb3.asp"				
    
Const TAB1 = 1
Const TAB2 = 2

Dim arrReturn
Dim gSelframeFlg
Dim gblnWinEvent
Dim arrParam

Dim arrParent
Dim PopupParent

ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0										
	lgStrPrevKey = ""	
	gblnWinEvent = False
End Function
	
'========================================================================================================
Sub SetDefaultVal()
	With frm1
		
		arrParam = arrParent(1)

		.txtLCNo.value = arrParam(0)
		.txtSONo.value = arrParam(1)
		
		gblnWinEvent = False
		Self.Returnvalue = ""
	End With
End sub	
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "RA") %>
	
End Sub

'========================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
		
	Call changeTabs(TAB1)
		
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
		
	Call changeTabs(TAB2)
		
	gSelframeFlg = TAB2
End Function
	
	
'=====================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
Sub Form_Load()
		
	Call LoadInfTB19029																
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											
	Call InitVariables
	Call SetDefaultVal
	Call DbQuery()
	Call changeTabs(TAB1)

End Sub
	

'========================================================================================================
Function DbQuery()
	Err.Clear															

	DbQuery = False														
			
	Call ggoOper.ClearField(Document, "2")								
	Call InitVariables													

	frm1.txtLocCurrency.value = PopupParent.gCurrency
		
	Dim strVal
        
	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001						
	strVal = strVal & "&txtLCNo=" & Trim(frm1.txtLCNo.value)			
	strVal = strVal & "&txtSONo=" & Trim(frm1.txtSONo.value)		
	strVal = strVal & "&txtLcKind=" & "L"
				
	Call RunMyBizASP(MyBizASP, strVal)										
   DbQuery = True														
End Function
	
'========================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtDocAmt, .txtCurrency.value, PopupParent.ggamtofmoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000,PopupParent.gComNumDec
	End With
End Sub
	
	
'========================================================================================================
Function DbQueryOk()													

	Call ggoOper.LockField(Document, "Q")								

	If gSelframeFlg <> TAB1 Then
		Call ClickTab1()
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
	<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
		<TABLE CLASS="BasicTB" CELLSPACING=0>
			<TR>
				<TD HEIGHT=5>&nbsp;</TD>
			</TR>
			<TR HEIGHT=23>
				<TD WIDTH=100%>
					<TABLE CLASS="BasicTB" CELLSPACING=0>
						<TR>
							<TD WIDTH=10>&nbsp;</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
									<TR>
										<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>LOCAL L/C 정보</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구비서류 및 기타</font></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>					
							<TD WIDTH=500>&nbsp;</TD>
							<TD WIDTH=*>&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR HEIGHT=*>
				<TD WIDTH=100% CLASS="Tab11">
					<TABLE CLASS="BasicTB">
						<TR>
							<TD HEIGHT=5 WIDTH=100%></TD>
						</TR>
						<TR>
							<TD HEIGHT=20 WIDTH=100%>
								<FIELDSET CLASS="CLSFLD">
									<TABLE CLASS="BasicTB" CELLSPACING=0>
										<TR>
											<TD CLASS=TD5 NOWRAP>LOCAL L/C관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCNo" ALT="LOCAL LC관리번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="14XXXU"></TD>
											<TD CLASS=TD5 NOWRAP>수주번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSONo" ALT="SO번호" TYPE=TEXT MAXLENGTH=18 SIZE=20 TAG="14XXXU"></TD>
										</TR>
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						
						<TR>
							<TD WIDTH=100% HEIGHT=100%>
								<!-- 첫번째 Tab 내용  -->
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
									<TABLE CLASS="BasicTB">
										<TR>
											<TD HEIGHT=2 WIDTH=100%></TD>
										</TR>  
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>LOCAL L/C 번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="LOCAL LC번호" TYPE=TEXT MAXLENGTH=35 SIZE=20 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>통지번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAdvNo" ALT="통지번호" TYPE=TEXT MAXLENGTH=35 SIZE=20 TAG="24XXXU"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>LOCAL L/C유형</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCType" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="LOCAL LC유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCType" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtLCTypeNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>통지일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra3_fpDateTime1_txtAdvDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>추심의뢰은행</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtFromBank" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="추심의뢰은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromBank" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtFromBankNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>유효일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra3_fpDateTime2_txtExpiryDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>개설은행</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOpenBank" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="개설은행"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenBank" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtOpenBankNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>개설일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra3_fpDateTime2_txtOpenDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>개설금액</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s3211ra3_fpDoubleSingle1_txtDocAmt.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>개설자국금액</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=11 MAXLENGTH=3 TAG="24XXXU" ALT="자국화폐">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s3211ra3_fpDoubleSingle2_txtLocAmt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>환율</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra3_fpDoubleSingle3_txtXchRate.js'></script>
														<TD CLASS=TD5 NOWRAP>선통지참조사항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRef" ALT="선통지참조사항" TYPE=TEXT MAXLENGTH=35 SIZE=34 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>물품인도기일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra3_fpDateTime1_txtMoveDt.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>분할인도여부</TD>
														<TD CLASS=TD6 NOWRAP><TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0><TR><TD WIDTH=30%><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="24" VALUE="Y" CHECKED ID="rdoPartailShip1"><LABEL FOR="rdoPartailShip1">Y</LABEL>&nbsp;&nbsp;&nbsp;<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoPartailShip" TAG="24X" VALUE="N" ID="rdoPartailShip2"><LABEL FOR="rdoPartailShip2">N</LABEL></TD></TR></TABLE></TD>									
													</TR>
																										<TR>
														<TD CLASS=TD5 NOWRAP>서류제시기간</TD>
														<TD CLASS=TD6 NOWRAP>
															<script language =javascript src='./js/s3211ra3_fpDoubleSingle5_txtFileDt.js'></script>&nbsp;DAYS
														</TD>
														<TD CLASS=TD5 NOWRAP>개설신청인</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="개설신청인">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>									
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>결제방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="결제방법">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>수혜자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수혜자">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>AMEND일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3211ra3_fpDateTime_txtAmendDt.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>영업그룹</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="영업그룹">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>	
								</DIV>
							
								<!-- 두번째 탭 내용 -->
								<DIV ID="TabDiv" STYLE="DISPLAY: none" SCROLL=no>
									<TABLE CLASS="BasicTB">  
										<TR>
											<TD HEIGHT=5 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>서류제시기간 참조</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFileDtTxt" ALT="서류제시기간 참조" TYPE=TEXT MAXLENGTH=35 SIZE=70 TAG="24"></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=2></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>구비서류</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc1" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="24"></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=2></TD>									
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc2" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="24"></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=2></TD>									
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc3" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="24"></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=2></TD>									
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc4" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="24"></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=2></TD>									
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP></TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDoc5" TYPE=TEXT MAXLENGTH=65 SIZE=70 TAG="24"></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=2></TD>									
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>개설은행앞 정보</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBankTxt" ALT="개설은행앞 정보" TYPE=TEXT MAXLENGTH=35 SIZE=70 TAG="24"></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=2></TD>									
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>기타참조사항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEtcRef" ALT="기타참조사항" TYPE=TEXT MAXLENGTH=35 SIZE=70 TAG="24"></TD>
														<TD CLASS=TD6 NOWRAP COLSPAN=2></TD>									
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</DIV>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD HEIGHT=30>
					<TABLE CLASS="basicTB" CELLSPACING=0>
						<TR>
							<TD WIDTH=* ALIGN=RIGHT>
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" 
							     onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%> ><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
			</TR>
		</TABLE>
	<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtFlgMode" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtHLCNo" TAG="24">
	<INPUT TYPE=HIDDEN NAME="txtHSoNo" TAG="24">
	</FORM>
	<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
		<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>           

