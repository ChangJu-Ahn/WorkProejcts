<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'**********************************************************************************************
'*  1. Module Name		    : Production
'*  2. Function Name	    : Entry BOM
'*  3. Program ID		    : p1711ma1.asp
'*  4. Program Name		    : 설계BOM 등록 
'*  5. Program Desc		    :
'*  6. Component List		:
'*  7. Modified date(First)	: 2000/04/21
'*  8. Modified date(Last)	: 2003/03/20
'*  9. Modifier (First)		: Kim, Gyoung-Don
'* 10. Modifier (Last)		: Hong Chang Ho
'* 11. Comment		:
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "p1711ma1.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

'=======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA")%>
End Sub

Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))

End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
	Call AppendNumberPlace("6", "5" ,"0")
	Call AppendNumberPlace("7", "2", "2")
	Call AppendNumberPlace("8", "11" ,"6")	
	
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
 	
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
   
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("11101000000011")
    Call InitCombobox()
    Call SetDefaultVal
	
	Call InitVariables		
	Call SetModChange(0)
	Call InitTreeImage
	Call SetFieldProp(51)   
	
	If parent.gPlant <> "" and CheckPlant(parent.gPlant) = True Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If
	
	Call txtPlantCd_OnChange()

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설계BOM등록</font></td>
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
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=28 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>기준일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtBaseDt CLASSID=<%=gCLSIDFPDT%> tag="12X1" ALT="기준일"></OBJECT>');</SCRIPT>
										</OBJECT>
									</TD>
									<TD CLASS=TD5 ROWSPAN=2 NOWRAP>단계</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoStepType" ID="rdoSrchType1" CLASS="RADIO" tag="1X" Value="1"><LABEL FOR="rdoStepType1">단단계</LABEL></TD>													     
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="12XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP><!--BOM Type--></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoStepType" ID="rdoSrchType2" CLASS="RADIO" tag="1X" Value="2" CHECKED><LABEL FOR="rdoStepType2">다단계</LABEL></TD>
								</TR>	
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VAlign=Top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT=* WIDTH=50%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=uniTree1 width=100% height=100% <%=UNI2KTV_IDVER%>> <PARAM NAME="ImageWidth" VALUE="16"> <PARAM NAME="ImageHeight" VALUE="16"> <PARAM NAME="LineStyle" VALUE="1"> <PARAM NAME="Style" VALUE="7"> <PARAM NAME="LabelEdit" VALUE="1"> </OBJECT>');</SCRIPT>
								</TD>
								<TD WIDTH="50%" HEIGHT=* VAlign=Top>
									<FIELDSET>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>자품목</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=25 MAXLENGTH=18  tag="22XXXU" ALT="자품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd 1" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자품목명</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=40 tag="24" ALT="자품목명"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목계정</TD>
												<TD CLASS=TD6 NOWRAP><SELECT NAME="cboItemAcct" CLASS=required STYLE="text-transform:uppercase; Width: 168px;" ALT="품목계정" tag="24"><INPUT TYPE=hidden NAME="txtItemAcctGrp" tag="24" TABINDEX="-1"></SELECT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>품목규격</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSpec" SIZE=40 tag="24" ALT="품목규격">
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>유효기간</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemFromDt CLASSID=<%=gCLSIDFPDT%> tag="24" ALT="시작일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
													&nbsp;~&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtPlantItemToDt CLASSID=<%=gCLSIDFPDT%> tag="24" ALT="종료일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
												</TD>	
											</TR>
										</TABLE>
									</FIELDSET>
									<FIELDSET>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>BOM 설명</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBOMDesc" SIZE=40 MAXLENGTH=40  tag="21" ALT="BOM 설명"></TD>													
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>도면경로</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDrawPath" SIZE=40 MAXLENGTH=100 tag=21 ALT="도면경로"></TD>
											</TR>
										</TABLE>	
									</FIELDSET>		
									<FIELDSET>
										<TABLE CLASS="BasicTB" CELLSPACING=0>			
											<TR>
												<TD CLASS=TD5 NOWRAP>순서</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtItemSeq CLASS=FPDS140 title=FPDOUBLESINGLE SIZE="15" MAXLENGTH="3" ALT="순서" tag="24X6Z"> </OBJECT>');</SCRIPT>
												</TD>													
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>자품목기준수</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtChildItemQty CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X8Z" ALT="자품목기준수" MAXLENGTH="15" SIZE="15"> </OBJECT>');</SCRIPT>
															</TD>
															<TD>
															&nbsp;<INPUT TYPE=TEXT NAME="txtChildItemUnit" SIZE=4 MAXLENGTH=3  tag="24" STYLE="Text-Transform: uppercase" ALT="자품목단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChildUnit" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit 0">
															</TD>
														</TR>
													</TABLE>
												</TD>														
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>모품목기준수</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLPADDING=0 CELLSPACING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtPrntItemQty CLASS=FPDS140 title=FPDOUBLESINGLE SIZE=15 MAXLENGTH=15 ALT="모품목기준수" tag="24X8Z"> </OBJECT>');</SCRIPT>
															</TD>
															<TD>
																&nbsp;<INPUT TYPE=TEXT NAME="txtPrntItemUnit" align=top SIZE=4 MAXLENGTH=3  tag="24" STYLE="Text-Transform: uppercase" ALT="모품목단위"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrntUnit" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUnit 1">
															</TD>	
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>안전L/T</TD>
												<TD CLASS=TD6 NOWRAP>
													<TABLE CELLSPACING=0 CELLPADDING=0>
														<TR>
															<TD>
																<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSafetyLt CLASS=FPDS140 title=FPDOUBLESINGLE SIZE="15" MAXLENGTH="3" ALT="안전L/T" tag="24X6Z"> </OBJECT>');</SCRIPT>
															</TD>
															<TD valign=bottom>
																&nbsp;일
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>Loss율(%)</TD>
												<TD CLASS=TD6 NOWRAP>
														<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtLossRate CLASS=FPDS140 title=FPDOUBLESINGLE tag="24X7Z" ALT="Loss율" MAXLENGTH="15" SIZE="15"></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>유무상구분</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoSupplyFlg" ID="rdoSupplyFlg1" CLASS="RADIO" tag="24X" Value="F" CHECKED><LABEL FOR="rdoSupplyFlg1">무상</LABEL>
												     				 <INPUT TYPE="RADIO" NAME="rdoSupplyFlg" ID="rdoSupplyFlg2" CLASS="RADIO" tag="24X" Value="C"><LABEL FOR="rdoSupplyFlg2">유상</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>비고</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRemark" SIZE=50 MAXLENGTH=1000 tag="21" ALT="비고"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>유효기간</TD>
												<TD CLASS=TD6 NOWRAP>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidFromDt1 CLASSID=<%=gCLSIDFPDT%> tag="24X1" ALT="시작일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
																										&nbsp;~&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtValidToDt1 CLASSID=<%=gCLSIDFPDT%> tag="24X1" ALT="종료일" MAXLENGTH="10" SIZE="10"></OBJECT>');</SCRIPT>
												</TD>	
											</TR>
										</TABLE>
									</FIELDSET>
									<FIELDSET>
										<TABLE CLASS="BasicTB" CELLSPACING=0>			
											<TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>설계변경번호</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNNo1" SIZE=25 MAXLENGTH=18  tag="24XXXU" ALT="설계변경번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnECNNo1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenECNNo frm1.txtECNNo1.value"  OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>설계변경내용</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtECNDesc1" SIZE=50 MAXLENGTH=100 tag="24" ALT="설계변경내용"></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>설계변경근거</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtReasonCd1" SIZE=5 MAXLENGTH=2  tag="24XXXU" ALT="설계변경근거"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnECNNo1" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenReasonCd frm1.txtReasonCd1.value"  OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT TYPE=TEXT NAME="txtReasonNm1" SIZE=25 tag="24"></TD>
											</TR>
										</TABLE>	
									</FIELDSET>		
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
<!--	
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadEBomHistory()">BOM이력조회</A>&nbsp;|&nbsp;<A href="vbscript:LoadPBomCreate()">제조BOM등록</A>&nbsp;|&nbsp;<A href="vbscript:LoadEBomToPBom()">제조BOM이관</A></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TABLE>
		</TD>
	</TR>	
-->
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdrMode" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtDtlMode" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrntItemCd" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBaseItemCd" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtPrntBomNo" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBaseBomNo" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBasicUnit" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtProcType" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBomType" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtValidFromDt" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtValidToDt" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPrntProcType" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBomNo"  tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtBomNo1" tag="14" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
