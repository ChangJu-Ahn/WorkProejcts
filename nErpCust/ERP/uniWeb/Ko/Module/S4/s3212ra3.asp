<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        : 																			*
'*  3. Program ID           : s3212ra3.asp																*
'*  4. Program Name         : Local L/C 내역정보(Local L/C현황조회에서)									*
'*  5. Program Desc         : Local L/C 내역정보(Local L/C현황조회에서)									*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2002/04/11																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kwak Eun Kyoung															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'*                            2. 2001/12/18 : Son bum Yeol												*
'*                            3. 2002/04/11 : ADO변환													*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>LOCAL L/C 내역정보</TITLE>

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
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              

<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_ID 		= "s3212rb3.asp"                              
Const C_MaxKey          = 16                                          

Dim gblnWinEvent
Dim arrReturn										
Dim arrParam	
Dim lgIsOpenPop

Dim arrParent
Dim PopupParent

ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

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

	arrParam = arrParent(1)
	
	With frm1
		.txtLCNo.value = arrParam(0)
		.txtLCDocNo.value = arrParam(1)
		.txtLCAmendSeq.value = arrParam(2)		
		.txtCurrency.value = arrParam(3)
		.txtLocAmt.value = arrParam(4)
 	End With
		
	Call CurFormatNumericOCX()
		
End Sub

'========================================================================================================
 Sub LoadInfTB19029()
 	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
 	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "RA") %>
 	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "RA") %>
 End Sub

'========================================================================================================
 Sub InitSpreadSheet()
	    
    Call SetZAdoSpreadSheet("S3212RA3","S","A","V20030318", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
 	With frm1.vspdData
 		ggoSpread.Source = frm1.vspdData
 		.OperationMode = 5
 		Call SetSpreadLock 
 	End With
	      
 End Sub

'========================================================================================================
 Sub SetSpreadLock()
     With frm1
     .vspdData.ReDraw = False
 	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 	ggoSpread.SpreadLockWithOddEvenRowColor()
 	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
     .vspdData.ReDraw = True

     End With
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
 	Redim arrReturn(0)
 	arrReturn(0) = ""
 	Self.Returnvalue = arrReturn
 	Self.Close()
 End Function

'========================================================================================================
Sub TotalSum()

	on error resume next
	
	Dim SumTotal, lRow
		
	SumTotal = 0
	ggoSpread.source = frm1.vspdData
	For lRow = 1 To frm1.vspdData.MaxRows 		
		frm1.vspdData.Row = lRow
		frm1.vspdData.Col = GetKeyPos("A",8)
		If frm1.vspdData.Text <> "" Then
			SumTotal = SumTotal + UNICDbl(frm1.vspdData.Text)
		End If
	Next	
	
	frm1.txtTotItemAmt.text = UNIFormatNumberByCurrecny(SumTotal,frm1.txtCurrency.value,PopupParent.ggAmtOfMoneyNo)
End Sub

'========================================================================================================
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
	End If
End Function		

'========================================================================================================
 Sub Form_Load()
		
 	Call LoadInfTB19029				                                            
 	Call ggoOper.LockField(Document, "N")						
 	Call InitVariables														    
 	Call SetDefaultVal	
 	Call InitSpreadSheet()
 	Call FncQuery()
	
 End Sub

'========================================================================================================
 Function vspdData_KeyPress(KeyAscii)
      On Error Resume Next
      If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
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

 	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
 		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
 			If DbQuery = False Then
 				Exit Sub
 			End if
 		End If
 	End If
 End Sub

'========================================================================================================
 Function FncQuery() 
    
 	FncQuery = False                                                 
    
 	Err.Clear                                                        

 	Call ggoOper.ClearField(Document, "2")							
 	Call InitVariables
 	
 	If DbQuery = False Then Exit Function							

 	FncQuery = True									
        
 End Function

'========================================================================================================
 Function DbQuery()
 	Err.Clear															

 	DbQuery = False														

 	If LayerShowHide(1) = False Then
 		Exit Function
 	End If

 	Dim strVal

 	With frm1
		
 		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
 		strVal = strVal & "&txtLCNo=" & Trim(.txtLCNo.value)					
 		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey

 	End With

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

	Call TotalSum()	

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtLCNo.focus
	End If

End Function

'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		'개설금액 
		ggoOper.FormatFieldByObjectOfCur .txtLocAmt, .txtCurrency.value, PopupParent.ggamtofmoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000,PopupParent.gComNumDec
		'총품목금액 
		ggoOper.FormatFieldByObjectOfCur .txtTotItemAmt, .txtCurrency.value, PopupParent.ggamtofmoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000,PopupParent.gComNumDec
	End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5 NOWRAP>LOCAL L/C관리번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLCNo" SIZE=20 MAXLENGTH=18 TAG="14XXXU"></TD>
							<TD CLASS=TD5 NOWRAP>개설금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<TABLE CELLSPACING=0 CELLPADDING=0>
									<TR>
										<TD>
											<INPUT TYPE=TEXT NAME="txtCurrency" SIZE=5 MAXLENGTH=3 TAG="14XXXU">&nbsp;
										</TD>
										<TD>
											&nbsp;&nbsp;&nbsp;&nbsp;<script language =javascript src='./js/s3212ra3_fpDoubleSingle1_txtLocAmt.js'></script>
										</TD>
									</TR>
								</TABLE>
							</TD>	
						</TR>	
						<TR>	
							<TD CLASS=TD5 NOWRAP>LOCAL L/C번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" TYPE=TEXT MAXLENGTH=35 SIZE=20 TAG="14XXXU">&nbsp;-&nbsp;<INPUT NAME="txtLCAmendSeq" TYPE=TEXT MAXLENGTH=4 SIZE=4 TAG="14XXXU"></TD>
							<TD CLASS=TD5 NOWRAP>총품목금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/s3212ra3_fpDoubleSingle2_txtTotItemAmt.js'></script>
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
							<script language =javascript src='./js/s3212ra3_vaSpread1_vspdData.js'></script>
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
						<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;						
						<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
						<TD WIDTH=* ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" 
						     onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
