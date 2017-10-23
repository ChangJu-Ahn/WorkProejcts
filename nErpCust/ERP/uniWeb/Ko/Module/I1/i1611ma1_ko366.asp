<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i1611ma1.asp
'*  4. Program Name         : 수불현황 조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2004/06/01
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript" SRC="i1611ma1_ko366.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" -->
                                            
Dim StartDate
Dim FromDate

Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgStrPrevKey4

StartDate = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)  
FromDate = UNIDateAdd("m", -1, StartDate, Parent.gDateFormat)        

'========================================  2.3 LoadInfTB19029()  =========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I", "NOCOOKIE", "QA") %>
End Sub



'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029
    Call LockObjectField(frm1.txtTrnsFrDt,"R")
    Call LockObjectField(frm1.txtTrnsToDt,"R")
    Call FormatDATEField(frm1.txtTrnsFrDt)
    Call FormatDATEField(frm1.txtTrnsToDt)
  Call GetValue_ko441()
	Call SetDefaultVal 
	Call InitComboBox
	Call InitVariables
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")    
	
End Sub

 '==========================================   InitComboBox()  ========================================
Sub InitComboBox()
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("I0002", "''", "S") & " ORDER BY MINOR_CD ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboTrnsType, lgF0, lgF1, Chr(11))
	
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ORDER BY MINOR_CD ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboItemAcct, lgF0, lgF1, Chr(11))
	
End Sub
'========================================================================================================= 
'	Name : openGlpopup
'	Description :전표  POP-UP
'========================================================================================================= 
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	Dim IntRetCD, cYear, TempGlNo, GlNo, ConfFg, GlRefNo
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	
	If lgIsOpenPop  = True Then Exit Function
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		msgbox "조회를 먼저 선행하십시오", vbinformation, "uniERP II-[Warning]"
		Exit Function
	End If
	
	frm1.vspddata.Row  = frm1.vspddata.ActiveRow
	
	'frm1.vspddata.col  = C_DocumentDt 'GetKeyPos("A",8)
	'cYear = Left(UniConvDate(frm1.vspddata.Text),4)
	
	frm1.vspddata.col  = C_DocumentNo 'GetKeyPos("A",9)
	
    Err.Clear

	Call CommonQueryRs(" Temp_Gl_No, Gl_No, Ref_No, Conf_Fg ", " A_Temp_Gl ", " Ref_No Like '" & Trim(frm1.vspddata.text) & "%'", _
	                   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	                   

	If lgF0 > "" Then
		TempGlNo = Replace(lgF0, Chr(11),"")
		GlNo	 = Replace(lgF1, Chr(11),"")
		GlRefNo	 = Replace(lgF2, Chr(11),"")
		ConfFg	 = Replace(lgF3, Chr(11),"")
	Else
		Call CommonQueryRs(" Gl_No, Ref_No ", " A_Gl ", " Ref_No Like '" & Trim(frm1.vspddata.text) & "%'", _
	                   lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	                   
		ConfFg	 = "C" '회계전표 자동기표시
		GlNo	 = Replace(lgF0, Chr(11),"")
		GlRefNo	 = Replace(lgF1, Chr(11),"")
	
	End If

	If Err.number <> 0 Then
	 MsgBox Err.description 
	 Err.Clear 
	 Exit Function
	End If
 
	If ConfFg = "U" Then
		iCalledAspName = AskPRAspName("A5130RA1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
			lgIsOpenPop  = False
			Exit Function
		End If
		
		arrParam(0) = TempGlNo	'결의전표번호 
		arrParam(1) = GlRefNo 'Reference번호 
	Else
		iCalledAspName = AskPRAspName("A5120RA1")

		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "A5120RA1", "X")
			lgIsOpenPop  = False
			Exit Function
		End If
		
		arrParam(0) = GlNo	'회계전표번호 
		arrParam(1) = GlRefNo'Reference번호
	
	End If
	
	
	lgIsOpenPop  = True
    arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		     	
	lgIsOpenPop  = False
End Function

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
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
		<TR>
			<TD <%=HEIGHT_TYPE_00%> >
			</TD>
		</TR>
		<TR HEIGHT=23>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_10%> WIDTH=100% border=0>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD CLASS="CLSMTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수불현황조회(S)</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
								</TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A HREF="VBSCRIPT:OpenPopupGL()">전표조회</A></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS="TD5" NOWRAP HEIGHT=5></TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=5></TD>
									<TD CLASS="TD5" NOWRAP HEIGHT=5></TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=5></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT SIZE="6" NAME="txtPlantCd" MAXLENGTH="7" CLASS=required STYLE="Text-Transform: uppercase" tag="12XXXU" ALT = "공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=29 MAXLENGTH=40 tag="14"></TD>    
									<TD CLASS="TD5" NOWRAP>수불기간</TD>
									<TD CLASS="TD6" NOWRAP>
									<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtTrnsFrDt classid=<%=gCLSIDFPDT%> ALT="검색시작날짜" tag="12X1"></OBJECT>&nbsp;~&nbsp;
									<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtTrnsToDt classid=<%=gCLSIDFPDT%> ALT="검색끝날짜" tag="12x1"></OBJECT>
									</TD>      
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>창고</TD>
									<TD CLASS="TD6" NOWRAP>
									<input TYPE=TEXT NAME="txtFrSlCd" SIZE="15" MAXLENGTH="18" STYLE="Text-Transform: uppercase" tag="11XXXU" ALT = "창고"><IMG align=top height=20 name="btnFrSlCd" onclick="vbscript:OpenSl1()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtFrSlNm" CLASS=protected readonly=true TABINDEX="-1" SIZE="20" tag="14" >&nbsp;~&nbsp;
									</TD>
									<TD CLASS="TD5" NOWRAP HEIGHT=5>수불구분</TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=5>
									<SELECT Name="cboTrnsType" ALT="수불구분" STYLE="WIDTH: 133px" tag="11"><OPTION Value=""></OPTION></SELECT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP>
									<input TYPE=TEXT NAME="txtToSlCd" SIZE="15" MAXLENGTH="18" STYLE="Text-Transform: uppercase" tag="11XXXU" ALT = "창고"><IMG align=top height=20 name="btnToSlCd" onclick="vbscript:OpenSl2()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtToSlNm" CLASS=protected readonly=true TABINDEX="-1" SIZE="20" tag="14" >
									</TD>
									<TD CLASS="TD5" NOWRAP HEIGHT=5>수불유형</TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=5>
										<TABLE CELLSPACING=0 CELLPADDING= 0>
											<TR>
												<TD>
													<INPUT TYPE=TEXT Name="txtMovType" SIZE="5" MAXLENGTH="3"  ALT="수불유형" tag="11XXXU"><IMG align=top height=20 name=btnMovType onclick="vbscript:OpenMovType()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtMovTypeNm" CLASS=protected readonly=true TABINDEX="-1" size="20" tag="14">
												</TD>
												<TD WIDTH="*">
													&nbsp;
												</TD>
												<TD  WIDTH="20" STYLE="TEXT-ALIGN: RIGHT" ><IMG SRC="../../../CShared/image/BigPlus.gif" Style="CURSOR: hand" ALT="DetailCondition" ALIGN= "TOP" ID = "IMG_DetailCondition" NAME="pop1" ONCLICK= 'vbscript:viewHidden "DetailCondition" ,2, 3' ></IMG></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR ID="DetailCondition1" style="display: none">
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP>
									<input TYPE=TEXT NAME="txtItemCd" SIZE="15" MAXLENGTH="18" STYLE="Text-Transform: uppercase" ALT="품목" tag="11XXXU" ><IMG align=top height=20 name="btnItemCd" onclick="vbscript:OpenItem()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtItemNm" CLASS=protected readonly=true TABINDEX="-1" SIZE="20" tag="14" >
									</TD>
									<TD CLASS="TD5" NOWRAP HEIGHT=5>품목계정</TD>
									<TD CLASS="TD6" NOWRAP HEIGHT=5>
									<SELECT Name="cboItemAcct" ALT="품목계정" STYLE="WIDTH: 133px" tag="11"><OPTION Value=""></OPTION></SELECT>
									</TD>
								</TR>
								<TR ID="DetailCondition2" style="display: none">
									<TD CLASS="TD5" NOWRAP>작업장</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtWcCd" SIZE="10" MAXLENGTH="7" STYLE="Text-Transform: uppercase" ALT="작업장" tag="11XXXU" ><IMG align=top height=20 name="btnWcCd" onclick="vbscript:OpenWcCd()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<INPUT TYPE=TEXT NAME="txtWcNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=20 MAXLENGTH=40 tag="14">
									</TD>
									<TD CLASS="TD5" NOWRAP HEIGHT=5>Tracking No.</TD>      
									<TD CLASS="TD6" NOWRAP HEIGHT=5>
									<INPUT TYPE=TEXT SIZE=20 NAME="txtTrackingNo" MAXLENGTH="25"  tag="11XXXU" ALT = "Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()">
									</TD>
								</TR>
								<TR ID="DetailCondition3" >
									<TD CLASS="TD5" NOWRAP>참조번호(전표)</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtRefNo" ALT="참조번호(전표)" STYLE="Text-Transform: uppercase" TYPE="Text" MAXLENGTH="20" SIZE=20 STYLE="" tag="11XXXU"></TD>

									<TD CLASS="TD5" NOWRAP HEIGHT=5></TD>      
									<TD CLASS="TD6" NOWRAP HEIGHT=5></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE WIDTH="100%" HEIGHT="100%" >
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD">
								<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
								</OBJECT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR HEIGHT="20">
					<TD HEIGHT=* WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT Name="BpCd" Size="15" MAXLENGTH="10" CLASS=protected readonly=true TABINDEX="-1" ALT="거래처" Tag="24">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" CLASS=protected readonly=true TABINDEX="-1" SIZE=30 MAXLENGTH=40 tag="24">
									</TD>
									<TD CLASS="TD5" NOWRAP>이동창고</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT Name="TrnsSlCd" Size="15" MAXLENGTH="10" CLASS=protected readonly=true TABINDEX="-1" ALT="이동창고" Tag="24">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>작업장</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT Name="WcCd" Size="15" MAXLENGTH="10" CLASS=protected readonly=true TABINDEX="-1" ALT="작업장" Tag="24">&nbsp;<INPUT TYPE=TEXT NAME="WcNm" SIZE=30 MAXLENGTH=40 CLASS=protected readonly=true TABINDEX="-1" tag="24">
									</TD>
									<TD CLASS="TD5" NOWRAP>비고</TD>
									<TD CLASS="TD6" MOWPAP>
									<INPUT TYPE=TEXT Name="Remark" SIZE=45 MAXLENGTH=40 CLASS=protected readonly=true TABINDEX="-1" tag="24">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</table>
		</TD>
		<TD>
			<TR>
				<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
				</TD>
			</TR>
		</TD>
	</TR>

	<TR>
	    <TD <%=HEIGHT_TYPE_01%> >
	    </TD>
	</TR>
		<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기(재고이동)</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄(재고이동)</BUTTON>&nbsp;
					</TD>					
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
	<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
	<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
	<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
	<INPUT TYPE=HIDDEN NAME="hTrnsFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hTrnsToDt" tag="24">
	<INPUT TYPE=HIDDEN NAME="hRefNo" tag="24"><INPUT TYPE=HIDDEN NAME="hOrderNo" tag="24">
	<INPUT TYPE=HIDDEN NAME="hFrSlCd" tag="24"><INPUT TYPE=HIDDEN NAME="hToSlCd" tag="24">
	<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24"><INPUT TYPE=HIDDEN NAME="hWcCd" tag="24">
	<INPUT TYPE=HIDDEN NAME="hMovType" tag="24"><INPUT TYPE=HIDDEN NAME="hTrnsType" tag="24">
	<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</B
