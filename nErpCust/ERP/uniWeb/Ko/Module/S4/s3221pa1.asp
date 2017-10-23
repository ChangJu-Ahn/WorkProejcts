<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : ��������																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3221pa1.asp																*
'*  4. Program Name         : L/C Amend��ȣ(L/C Amend��Ͽ���)	 										*
'*  5. Program Desc         : L/C Amend��ȣ(L/C Amend��Ͽ���)											*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/03																*
'*  8. Modified date(Last)  : 2002/07/09																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/04/03 : ȭ�� design												*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>L/C AMEND ��ȣ</TITLE>

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


Dim gblnWinEvent				
Dim strReturn																   
Dim lgIsOpenPop

Dim arrParent
Dim PopupParent

Dim prDBSYSDate
Dim EndDate ,StartDate

ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

prDBSYSDate = "<%=GetSvrDate%>"

EndDate = PopupParent.UNIConvDateAToB(prDBSYSDate ,PopupParent.gServerDateFormat,PopupParent.gDateFormat)               'Convert DB date type to Company
StartDate = PopupParent.UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

Const BIZ_PGM_ID        = "s3221pb1.asp"
Const C_MaxKey          = 3                                  
Const gstrLCTypeMajor	= "S9000"					

'===========================================================================================================
Sub InitVariables()
         
    lgBlnFlgChgValue = False                               
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1
    lgPageNo         = ""
	lgIntFlgMode = PopupParent.OPMD_CMODE	
	
	ReDim strReturn(0)
	strReturn = ""
	gblnWinEvent = False
	Self.Returnvalue = strReturn

End Sub

'===========================================================================================================
Sub SetDefaultVal()

	frm1.txtFromDt.text = StartDate
	frm1.txtToDt.text = EndDate

End Sub

'===========================================================================================================
Sub LoadInfTB19029()
	
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "PA") %>
End Sub

'===========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("s3221pa1","S","A","V20030320", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		.OperationMode = 3
		Call SetSpreadLock 
	End With
      
End Sub

'===========================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub

'===========================================================================================================
 Function OKClick()
		
 	dim arrReturn
 	If frm1.vspdData.ActiveRow > 0 Then				
		
 		frm1.vspdData.Row = frm1.vspdData.ActiveRow
 		frm1.vspdData.Col = GetKeyPos("A",1)
 		arrReturn = Trim(frm1.vspdData.Text)

 		Self.Returnvalue = arrReturn
 	End If

 	Self.Close()
 End Function
 
'===========================================================================================================
 Function CancelClick()
 	Self.Close()
 End Function

'===========================================================================================================
 Function OpenBizPartner()
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function

 	gblnWinEvent = True

 	arrParam(0) = "������"							
 	arrParam(1) = "B_BIZ_PARTNER"						
 	arrParam(2) = Trim(frm1.txtApplicant.value)				
 	arrParam(3) = ""									
 	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
 	arrParam(5) = "������"							

 	arrField(0) = "BP_CD"								
 	arrField(1) = "BP_NM"								

 	arrHeader(0) = "������"							
 	arrHeader(1) = "�����ڸ�"						

 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetBizPartner(arrRet)
 	End If
 End Function
	
'===========================================================================================================
 Function OpenSalesGroup()
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function

 	gblnWinEvent = True

 	arrParam(0) = "�����׷�"						
 	arrParam(1) = "B_SALES_GRP"							
 	arrParam(2) = Trim(frm1.txtSalesGroup.value)				
 	arrParam(3) = ""									
 	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						
 	arrParam(5) = "�����׷�"						

 	arrField(0) = "SALES_GRP"							
 	arrField(1) = "SALES_GRP_NM"						

 	arrHeader(0) = "�����׷�"						
 	arrHeader(1) = "�����׷��"						

 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetSalesGroup(arrRet)
 	End If
 End Function
	
'===========================================================================================================
 Function OpenLCDocNo()
 	Dim arrRet
 	Dim arrParam(5), arrField(6), arrHeader(6)

 	If gblnWinEvent = True Then Exit Function

 	gblnWinEvent = True
	
 	arrParam(0) = "L/C��ȣ"									
 	arrParam(1) = "S_LC_AMEND_HDR"									
 	arrParam(2) = Trim(frm1.txtLCDocNo.value)						
 	arrParam(3) = ""											
 	arrParam(4) = "LC_KIND = " & FilterVar("M", "''", "S") & " "								
 	arrParam(5) = "L/C��ȣ"									

 	arrField(0) = "ED33" & PopupParent.gColSep & "LC_DOC_NO"						
 	arrField(1) = "ED12" & PopupParent.gColSep & "LC_AMEND_SEQ"	                    '����--CONVERT(char(3),LC_AMEND_SEQ) ǥ�عݿ������� convert ���� 2002-06-17

 	arrHeader(0) = "L/C��ȣ"								
 	arrHeader(1) = "AMEND����"											

 	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
 			"dialogWidth=422px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 	gblnWinEvent = False

 	If arrRet(0) = "" Then
 		Exit Function
 	Else
 		Call SetLCDocNo(arrRet)
 	End If
 End Function
	
'===========================================================================================================
 Function SetBizPartner(arrRet)
 	frm1.txtApplicant.Value = arrRet(0)
 	frm1.txtApplicantNm.Value = arrRet(1)
 	frm1.txtApplicant.focus
 End Function

'===========================================================================================================
 Function SetSalesGroup(arrRet)
 	frm1.txtSalesGroup.value = arrRet(0)
 	frm1.txtSalesGroupNm.value = arrRet(1)
 	frm1.txtSalesGroup.focus
 End Function
 
'===========================================================================================================
 Function SetLCDocNo(arrRet)
 	frm1.txtLCDocNo.Value = arrRet(0)
 	frm1.txtLCDocNo.focus
 End Function	
 
'===========================================================================================================
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
'===========================================================================================================
Sub Form_Load()
	
 Call LoadInfTB19029			
 Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
 Call ggoOper.LockField(Document, "N")                                   
 Call InitVariables			    
 Call SetDefaultVal	
 Call InitSpreadSheet()
 Call FncQuery()

End Sub
'===========================================================================================================
 Sub Form_QueryUnload(Cancel, UnloadMode)
 End Sub

'===========================================================================================================
 Sub btnApplicantOnClick()
 	Call OpenBizPartner()
 End Sub
'===========================================================================================================
 Sub btnSalesGroupOnClick()
 	Call OpenSalesGroup()
 End Sub
	

'===========================================================================================================
 Sub btnLCDocNoOnClick()
 	Call OpenLCDocNo()
 End Sub
	
'===========================================================================================================
 Sub txtFromDt_DblClick(Button)
     If Button = 1 Then
         frm1.txtFromDt.Action = 7 
         Call SetFocusToDocument("P")
		 frm1.txtFromDt.Focus
     End If
 End Sub

'===========================================================================================================
 Sub txtToDt_DblClick(Button)
     If Button = 1 Then
         frm1.txtToDt.Action = 7
         Call SetFocusToDocument("P")
         frm1.txtToDt.Focus
     End If
 End Sub

'===========================================================================================================
 Sub txtFromDt_Keypress(KeyAscii)
 	On Error Resume Next
 	If KeyAscii = 27 Then
 		Call CancelClick()
 	Elseif KeyAscii = 13 Then
 		Call FncQuery()
 	End if
 End Sub

 Sub txtToDt_Keypress(KeyAscii)
 	On Error Resume Next
 	If KeyAscii = 27 Then
 		Call CancelClick()
 	Elseif KeyAscii = 13 Then
 		Call FncQuery()
 	End if
 End Sub

'===========================================================================================================
 Function vspdData_KeyPress(KeyAscii)
      On Error Resume Next
      If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1������ frm1���� 
         Call OKClick()
      ElseIf KeyAscii = 27 Then
         Call CancelClick()
      End If
 End Function
	
'===========================================================================================================
 Function vspdData_DblClick(ByVal Col, ByVal Row)
 	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
 		Exit Function
 	End If
 	If frm1.vspdData.MaxRows > 0 Then
 		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
 			Call OKClick
 		End If
 	End If
 End Function
	
'===========================================================================================================
 Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
 	With frm1.vspdData
 		If Row >= NewRow Then
 			Exit Sub
 		End If
 		If NewRow = .MaxRows Then
 			If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
 				DbQuery
 			End If
 		End If
 	End With
 End Sub

'===========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
	If OldLeft <> NewLeft Then
	  	Exit Sub
	End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	  '��: ������ üũ	
		If lgPageNo <> "" Then                   '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
    
End Sub

'===========================================================================================================
Function FncQuery() 
	FncQuery = False                                                        
	    
	Err.Clear                                                               

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", PopupParent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
	End If

	Call ggoOper.ClearField(Document, "2")									
	Call InitVariables 														
		
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
				
	Call DbQuery															'��: Query db data

	FncQuery = True		
End Function

'==========================================================================================================
Function DbQuery() 
 Dim strVal

 DbQuery = False
    
 Err.Clear                                                               
	
 If LayerShowHide(1) = False Then
 	Exit Function
 End If
    
 With frm1
    
 	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
 		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
 		strVal = strVal & "&txtApplicant=" & Trim(frm1.txtHApplicant.value)		
 		strVal = strVal & "&txtSalesGroup=" & Trim(frm1.txtHSalesGroup.value)	
 		strVal = strVal & "&txtLCDocNo=" & Trim(frm1.txtHLCDocNo.value)			
 		strVal = strVal & "&txtFromDt=" & Trim(frm1.txtHFromDt.value)		
 		strVal = strVal & "&txtToDt=" & Trim(frm1.txtHToDt.value)		
			
 	Else
			
 		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
 		strVal = strVal & "&txtApplicant=" & Trim(frm1.txtApplicant.value)		
 		strVal = strVal & "&txtSalesGroup=" & Trim(frm1.txtSalesGroup.value)		
 		strVal = strVal & "&txtLCDocNo=" & Trim(frm1.txtLCDocNo.value)			
 		strVal = strVal & "&txtFromDt=" & Trim(frm1.txtFromDt.text)		
 		strVal = strVal & "&txtToDt=" & Trim(frm1.txtToDt.text)		
			
 	End If       
 			
 		strVal = strVal & "&lgPageNo="       & lgPageNo                
 		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
 		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
 		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
			
    Call RunMyBizASP(MyBizASP, strVal)										
			
 End With

 DbQuery = True

End Function

'===========================================================================================================
Function DbQueryOk()														
 lgIntFlgMode = PopupParent.OPMD_UMODE

If frm1.vspdData.MaxRows > 0 Then
    frm1.vspdData.Focus		
Else
    frm1.txtApplicant.focus
End If     

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5>������</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnApplicantOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>�����׷�</TD>
							<TD CLASS=TD6 NOWRAP>
								<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="�����׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSalesGroupOnClick()">
								<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
							</TD>
						<TR>
							<TD CLASS=TD5 NOWRAP>L/C��ȣ</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtLCDocNo" ALT="L/C��ȣ" TYPE=TEXT SIZE=35 MAXLENGTH=35 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLCDocNo" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnLCDocNoOnClick()"></TD>
							<TD CLASS=TD5 NOWRAP>AMEND ��</TD>						
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/s3221pa1_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/s3221pa1_fpDateTime2_txtToDt.js'></script>
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
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/s3221pa1_vaSpread_vspdData.js'></script>
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
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>&nbsp;
						<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
						<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHLCDocNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


