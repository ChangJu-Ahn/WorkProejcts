<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4211ra9.asp																*
'*  4. Program Name         : 통관상세정보(통관현황조회에서)											*
'*  5. Program Desc         : 통관상세정보(통관현황조회에서)											*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/04/11																*
'*  8. Modified date(Last)  : 2001/12/17																*
'*  9. Modifier (First)     : KIm Hyungsuk																*
'* 10. Modifier (Last)      : Park insik																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"	SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<Script Language="VBSCRIPT">

Option Explicit					

Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)


Const BIZ_PGM_ID = "s4211rb9.asp"			

Const TAB1 = 1
Const TAB2 = 2
Const TAB3 = 3

Dim lgBlnFlgChgValue				'☜: Variable is for Dirty flag 
Dim lgIntGrpCount					'☜: Group View Size를 조사할 변수 
Dim lgIntFlgMode					'☜: Variable is for Operation Status
Dim lgBlnSoQueryFlg					'☜: 수주참조가 되었다는것을 나타내는 변수 	
	
Dim gSelframeFlg					'현재 TAB의 위치를 나타내는 Flag
Dim gblnWinEvent					

'========================================================================================================
Function InitVariables()
		
	lgIntFlgMode = PopupParent.OPMD_CMODE				
	lgBlnFlgChgValue = False							
	lgIntGrpCount = 0									
	gblnWinEvent = False
		
End Function

'========================================================================================================
Sub SetDefaultVal()
	Dim arrParam
	arrParam = arrParent(1)	
		
	With frm1
		.txtCCNo.value = arrParam(0)
		.txtIvNo.value = arrParam(1)
	End With

	gblnWinEvent = False
	Self.Returnvalue = ""
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
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
	
Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function
		
	Call changeTabs(TAB3)
		
	gSelframeFlg = TAB3
End Function
'==========================================================================================================
Function CancelClick()
	Self.Close()
End Function
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029								
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")			
 
	Call InitVariables
	Call SetDefaultVal
	Call DbQuery()	
	Call changeTabs(TAB1)
End Sub
	
'********************************************************************************************************
Function DbQuery()
	Err.Clear															

	DbQuery = False														

	Dim strVal
				
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
	strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)		

	Call RunMyBizASP(MyBizASP, strVal)									
	
	DbQuery = True														
End Function

'==========================================================================================================
Function DbQueryOk()	
	
	If gSelframeFlg <> TAB1 Then
		Call ClickTab1()
	End If
		
	frm1.txtLocCCCurrency.value = PopupParent.gCurrency
	frm1.txtLocFobCurrency.value = PopupParent.gCurrency
		
End Function
	

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수출신고</font></td>
										<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								    </TR>
								</TABLE>
							</TD>
							<TD CLASS="CLSMTABP">
								<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
									<TR>
										<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
										<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수출면허</font></td>
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
											<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
											<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=20 MAXLENGTH=18 TAG="14XXXU" ALT="통관관리번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCCNo" ALIGN=top TYPE="BUTTON"></TD>
											<TD CLASS=TD6 NOWRAP></TD>
											<TD CLASS=TD6 NOWRAP></TD>
										</TR>
									</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD WIDTH=100% HEIGHT=100%>
							<!-- 첫번째 탭 내용 -->
								<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
									<TABLE CLASS="BasicTB">
										<TR>
											<TD HEIGHT=2 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD WIDTH=100% HEIGHT=100%>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo1" SIZE=20 MAXLENGTH=18 TAG="24XXXU" ALT="통관관리번호"></TD>
														<TD CLASS=TD5 NOWRAP>수주번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSONo" SIZE=20  MAXLENGTH=18 TAG="24XXXU" ALT="수주번호"></TD>
													</TR>				
													<TR>	
														<TD CLASS=TD5 NOWRAP>송장번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIVNo" ALT="송장번호" MAXLENGTH=35 TYPE=TEXT SIZE=35 TAG="24XXXU">
														<TD CLASS=TD5 NOWRAP>작성일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDateTime3_txtIVDt.js'></script></TD>														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>제출번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEDNo" ALT="제출번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="24XXXU">
														<TD CLASS=TD5 NOWRAP>신고일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDateTime2_txtEDDt.js'></script></TD>
													</TR>												
													<TR>
													    <TD CLASS=TD5 NOWRAP>신고번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEPNo" ALT="신고번호" TYPE=TEXT MAXLENGTH=35 SIZE=35 TAG="24X"></TD>
														<TD CLASS=TD5 NOWRAP>면허일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDateTime4_txtEPDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>VESSEL명</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtVesselNm" ALT="VESSEL명" TYPE=TEXT MAXLENGTH=50 SIZE=35 TAG="24X"></TD>
														<TD CLASS=TD5 NOWRAP>선적완료일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDateTime1_txtShipFinDt.js'></script></TD>																												
													</TR>
																										
													<TR>														
														<TD CLASS=TD5 NOWRAP>중량단위</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtWeightUnit" ALT="중량단위" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnWeightUnit" align=top TYPE="BUTTON"></TD>														
														<TD CLASS=TD5 NOWRAP>L/C번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="L/C번호" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="24XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT STYLE="TEXT-ALIGN: center" MAXLENGTH=1 SIZE=1 TAG="24"></TD>														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>총중량</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDoubleSingle1_txtGrossWeight.js'></script>
														<TD CLASS=TD5 NOWRAP>총순중량</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDoubleSingle1_txtNetWeight.js'></script>														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>화폐</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ALT="화폐"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" align=top TYPE="BUTTON"></TD>																												
														<TD CLASS=TD5 NOWRAP>총포장개수</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDoubleSingle1_txtTotPackingCnt.js'></script>
													</TR>
													
													<TR>
														<TD CLASS=TD5 NOWRAP>환율</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDoubleSingle3_txtXchRate.js'></script>
														<TD CLASS=TD5 NOWRAP>USD환율</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDoubleSingle3_txtUsdXchRate.js'></script>
														
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>통관금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><INPUT TYPE=TEXT NAME="txtCCCurrency" ALT="통관금액" SIZE=10 MAXLENGTH=3 TAG="24XXXU" ></TD>
																	<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s4211ra9_fpDoubleSingle1_txtDocAmt.js'></script>
																	</TD>
																</TR>
															</TABLE>
														</TD>
														<TD CLASS=TD5 NOWRAP>통관자국금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><INPUT TYPE=TEXT NAME="txtLocCCCurrency" ALT="통관자국금액" SIZE=10 MAXLENGTH=3 TAG="24XXXU"></TD>
																	<TD>&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s4211ra9_fpDoubleSingle2_txtLocAmt.js'></script></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>FOB금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><INPUT TYPE=TEXT NAME="txtFobCurrency"  ALT="FOB금액" SIZE=10 MAXLENGTH=3 TAG="24XXXU"></TD>
																	<TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s4211ra9_fpDoubleSingle1_txtFobDocAmt.js'></script>
																	</TD>
																</TR>
															</TABLE>
														</TD>
														<TD CLASS=TD5 NOWRAP>FOB자국금액</TD>
														<TD CLASS=TD6 NOWRAP>
															<TABLE CELLSPACING=0 CELLPADDING=0>
																<TR>
																	<TD><INPUT TYPE=TEXT NAME="txtLocFobCurrency" ALT="FOB자국금액" SIZE=10 MAXLENGTH=2 TAG="24XXXU"></TD>
																	<TD>&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s4211ra9_fpDoubleSingle2_txtFobLocAmt.js'></script></TD>
																</TR>
															</TABLE>
														</TD>
													</TR>
													<TR>	
														<TD CLASS=TD5 NOWRAP>가격조건</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoTerms" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="가격조건"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncoTerms" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtIncoTermsNm" SIZE=20 TAG="24"></TD>										
														<TD CLASS=TD5 NOWRAP>영업그룹</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>결제방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>결제기간</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDoubleSingle5_txtPayDur.js'></script>&nbsp;일</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>수입자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수입자">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>수출자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="수출자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>대행자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAgent" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="대행자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAgent" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtAgentNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>제조자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtManufacturer" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="제조자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnManufacturer" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtManufacturerNm" SIZE=20 TAG="24"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</DIV>
								<!-- 두번째 탭 내용 -->
								<DIV ID="TabDiv" SCROLL=no>
									<TABLE CLASS="BasicTB">
										<TR>
											<TD HEIGHT=5 WIDTH=100%></TD>
										</TR>
										<TR>
											<TD>
												<TABLE CLASS="TB3" CELLSPACING=0 CELLPADDING=0>													
													<TR>
														<TD CLASS=TD5 NOWRAP>선적항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingPort" ALT="선적항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingPort" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingPortNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>선적항국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLoadingCntry" ALT="선적항국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLoadingCntry" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtLoadingCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>도착항</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgePort" ALT="도착항" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgePort" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgePortNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>도착항국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDischgeCntry" ALT="도착항국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDischgeCntry" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtDischgeCntryNm" SIZE=20 TAG="24"></TD>
													</TR>
													
													
													<TR>
														<TD CLASS=TD5 NOWRAP>원산지</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOrigin" ALT="원산지" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOrigin" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>원산지국가</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOriginCntry" ALT="원산지국가" TYPE=TEXT MAXLENGTH=3 SIZE=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOriginCntry" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtOriginCntryNm" SIZE=20 TAG="24"></TD>
													</TR>														
													<TR>
														<TD CLASS=TD5 NOWRAP>최종목적지</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFinalDest" ALT="최종목적지" TYPE=TEXT MAXLENGTH=120 SIZE=35 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>신고자</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReporter" SIZE=10 MAXLENGTH=10 TAG="24XXXU" ALT="신고자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReporter" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtReporterNm" SIZE=20 TAG="24"></TD>
													</TR>	
													<TR>
														<TD CLASS=TD5 NOWRAP>환급신청인</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReturnAppl" SIZE=10 MAXLENGTH=10 TAG="24XXU" ALT="환급신청인"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReturnAppl" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtReturnApplNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>환급기관</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReturnOffice" SIZE=10 MAXLENGTH=30 TAG="24XXXU" ALT="환급기관"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReturnOffice" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtReturnOfficeNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>신고구분</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtEDType" SIZE=10 MAXLENGTH=5 TAG="24XXXU" ALT="신고구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnEDType" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtEDTypeNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>세관</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCustoms" ALT="세관" SIZE=10 MAXLENGTH=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCustoms" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtCustomsNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>컨테이너 운송방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransForm" ALT="컨테이너 운송방법" SIZE=10 MAXLENGTH=5 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransForm" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtTransFormNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>포장조건</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPackingType" ALT="포장조건" SIZE=10 MAXLENGTH=5 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPackingType" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtPackingTypeNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>운송신고인</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransRepCd" SIZE=10 MAXLENGTH=10 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransRepCd" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtTransRepNm" SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>보세운송방법</TD>
														<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTransMeth" SIZE=10 MAXLENGTH=5 TAG="24XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransMeth" align=top TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtTransMethNm" SIZE=20 TAG="24"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>운송시작일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDateTime5_txtTransFromDt.js'></script></TD>
														<TD CLASS=TD5 NOWRAP>운송종료일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDateTime6_txtTransToDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>검사증번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInspCertNo" ALT="검사증번호" TYPE=TEXT MAXLENGTH=20 SIZE=20 TAG="24"></TD>
														<TD CLASS=TD5 NOWRAP>검사증발급일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDateTime7_txtInspCertDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>검역증번호</TD>
														<TD CLASS=TD6 NOWRAP><INPUT NAME="txtQuarCertNo" ALT="검역증번호" TYPE=TEXT MAXLENGTH=20 SIZE=20 TAG="24X"></TD>
														<TD CLASS=TD5 NOWRAP>검역증발급일</TD>
														<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4211ra9_fpDateTime8_txtQuarCertDt.js'></script></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>장치장소</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtDevicePlce" ALT="장치장소" TYPE=TEXT MAXLENGTH=120 SIZE=80 TAG="24X"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>참고사항 1</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark1" ALT="참고사항1" TYPE=TEXT MAXLENGTH=120 SIZE=80 TAG="24X"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>2</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark2" ALT="참고사항2" TYPE=TEXT MAXLENGTH=120 SIZE=80 TAG="24X"></TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>3</TD>
														<TD CLASS=TD6 COLSPAN=3><INPUT NAME="txtRemark3" ALT="참고사항3" TYPE=TEXT MAXLENGTH=120 SIZE=80 TAG="24X"></TD>
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
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=100 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD>
			</TR>
		</TABLE>
		<INPUT TYPE=HIDDEN NAME="txtMode" TAG="24">
		<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHPayTerms" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHIncoterms" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHPayDur" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtHCCNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtLCNo" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtRefFlg" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtSONoFlg" tag="24">
		<INPUT TYPE=HIDDEN NAME="txtPrevNext" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>