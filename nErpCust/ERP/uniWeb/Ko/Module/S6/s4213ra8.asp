<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s4213ra8.asp																*
'*  4. Program Name         : 통관란정보(통관현황조회에서)												*
'*  5. Program Desc         : 통관란정보(통관현황조회에서)												*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2002/04/25																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : RYU KYUNGRAE																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'*                            2. 2002/04/25 : ADO변환                                                   *
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

Const BIZ_PGM_ID 		= "S4213RB8.asp"                                 
Const C_MaxKey          = 3                                           

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent

Dim arrReturn										
Dim arrParent
Dim lgIsOpenPop

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
'========================================================================================================
Function InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          
    lgSortKey        = 1
			
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
'========================================================================================================
Sub SetDefaultVal()

	Dim arrParam
	
	arrParam = arrParent(1)
			
	frm1.txtCCNo.value = arrParam(0)
	frm1.txtIvNo.value = arrParam(1)
	frm1.txtCurrency.value = arrParam(2)
	frm1.txtFOBAmt.text = arrParam(3)
	frm1.txtLocCurrency.value = arrParam(4)
	frm1.txtFOBLocAmt.Text = arrParam(5)

	If Len(Trim(frm1.txtCcNo.value)) Then Call ggoOper.SetReqAttr(frm1.txtCcNo, "Q")	

End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "RA") %>
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S4213RA801","S","A","V20030322",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
								C_MaxKey, "X","X")		
	Call SetSpreadLock 								
	      
End Sub
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'========================================================================================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount, frm1.vspdData.MaxCols)

		For intRowCnt = 1 To frm1.vspdData.MaxRows

			frm1.vspdData.Row = intRowCnt

			If frm1.vspdData.SelModeSelected Then
				For intColCnt = 0 To frm1.vspdData.MaxCols - 1
					frm1.vspdData.Col = intColCnt + 1
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
				Next

				intInsRow = intInsRow + 1

			End IF
		Next
	End if			
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function	
'========================================================================================================
Function CancelClick()

	Redim arrReturn(1,1)

	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
	
End Function
'========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029				                                           

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)	
	Call ggoOper.LockField(Document, "N")					

	Call InitVariables														    
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If    

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If lgPageNo <> "" Then		                                                    
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtFOBAmt, .txtCurrency.value, PopupParent.ggamtofmoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000,PopupParent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtFOBLocAmt, .txtLocCurrency.value, PopupParent.ggamtofmoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000,PopupParent.gComNumDec
	End With
End Sub
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function
'========================================================================================
Function FncQuery() 
    
    FncQuery = False

    Err.Clear                                                        

	Call InitVariables

    If DbQuery = False Then Exit Function

    FncQuery = True			
        
End Function
'********************************************************************************************************
Function DbQuery()
	Err.Clear															

	DbQuery = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)					
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
		strVal = strVal & "&txtLocCurrency=" & Trim(frm1.txtLocCurrency.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
		strVal = strVal & "&txtCCNo=" & Trim(frm1.txtCCNo.value)					
		strVal = strVal & "&txtCurrency=" & Trim(frm1.txtCurrency.value)
		strVal = strVal & "&txtLocCurrency=" & Trim(frm1.txtLocCurrency.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If
'--------- Developer Coding Part (End) ------------------------------------------------------------
    strVal = strVal & "&lgPageNo="       & lgPageNo                              
	strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
	strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))	

	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True													
End Function
'========================================================================================
Function DbQueryOk()														

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtCcNo.focus
	End If

End Function
'===========================================================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()  
       Call FncQuery()     
   End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=20 MAXLENGTH=18 TAG="14XXXU"></TD>
						<TD CLASS=TD5 NOWRAP>송장번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIvNo" TYPE=TEXT MAXLENGTH=35 SIZE=20 TAG="14XXXU">&nbsp;</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>통관FOB금액</TD>
						<TD CLASS=TD6 NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=5 MAXLENGTH=3 TAG="24XXXU">&nbsp;</TD>
									<TD>
										&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s4213ra8_fpDoubleSingle1_txtFOBAmt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>	
						<TD CLASS=TD5 NOWRAP>통관FOB원화금액</TD>
						<TD CLASS=TD6 NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<INPUT TYPE=TEXT NAME="txtLocCurrency" SIZE=5 MAXLENGTH=3 TAG="24XXXU">&nbsp;
									</TD>
									<TD>
										&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s4213ra8_fpDoubleSingle1_txtFOBLocAmt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>	
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=* valign=top>
				<TABLE WIDTH="100%" HEIGHT="100%">
					<TR>
						<TD HEIGHT="100%">
							<script language =javascript src='./js/s4213ra8_vaSpread1_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR HEIGHT="20">
			<TD WIDTH="100%">
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>					
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>						
							<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>						
						<TD WIDTH=30% ALIGN=RIGHT>							
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG></TD>
						<TD WIDTH=10>&nbsp;</TD>					
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="HCCNO" tag="24">
<INPUT TYPE=HIDDEN NAME="HIvNo" tag="24">
<INPUT TYPE=HIDDEN NAME="HCurrency" tag="24">
<INPUT TYPE=HIDDEN NAME="HFOBAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="HLocCurrency" tag="24">
<INPUT TYPE=HIDDEN NAME="HFOBLocAmt" tag="24">
</FORM>		
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
  </DIV>
</BODY>
</HTML>