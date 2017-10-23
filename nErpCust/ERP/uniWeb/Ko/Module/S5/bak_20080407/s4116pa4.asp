<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S4116PA4
'*  4. Program Name         : 출고상세현황 
'*  5. Program Desc         : 출고상세현황 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/29
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit				


Dim lgIsOpenPop                                

Dim lgMark                                     
Dim lgblnWinEvent					'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
									'PopUp Window가 사용중인지 여부를 나타내는 variable
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim arrParent
arrParent = window.dialogArguments

Set PopupParent = arrParent(0)

top.document.title = PopupParent.gActivePRAspName

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID = "s4116pb4.asp"			              
Const C_MaxKey          = 20                 
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------	
	
'=============================================
Sub InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False                               
    lgSortKey        = 1
End Sub

'=============================================
Sub SetDefaultVal()
	
	'--------------- 개발자 coding part(실행로직,Start)---------------------------------------------------
	Dim arrParam	
	
	arrParam = arrParent(1)
	
	With frm1
	
		.txtConDNNo.value = arrParam(0)		
		.txtConShipToParty.value = arrParam(1)
		.txtConShipToPartyNm.value = arrParam(2)
		.txtConDnType.value = arrParam(3)
		.txtConDnTypeNm.value = arrParam(4)
		.txtConFromDt.text = arrParam(5)
		.txtConToDt.text = arrParam(6)
		
		If arrParam(7) = "Y" Then
			.rdoConf.checked = True
			.txtConConfFlag.value = .rdoConf.value
		Else
			.rdoNonConf.checked = True
			.txtConConfFlag.value = .rdoNonConf.value
		End If
		
		Call GetAmtLoc()
		
	End With

	lgblnWinEvent = False
	Self.Returnvalue = ""
	'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------

End Sub

'============================================
Sub GetAmtLoc()
	
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs
	
	iStrSelectList = " SUM(ISNULL(DL.GI_AMT_LOC,0) + ISNULL(DL.VAT_AMT_LOC,0)), " & _
					 " SUM(ISNULL(DL.GI_AMT_LOC,0)), SUM(ISNULL(DL.VAT_AMT_LOC,0)), " & _
					 " SUM(ISNULL(DL.DEPOSIT_AMT,0)) "
	iStrFromList  = " S_DN_HDR DH, S_DN_DTL DL, B_ITEM IT "
	iStrWhereList = " DH.DN_NO = DL.DN_NO AND DL.ITEM_CD = IT.ITEM_CD "
	
	If frm1.txtConFromDt.text <> "" Then
		If frm1.txtConConfFlag.value = "Y" Then
			iStrWhereList = iStrWhereList & " AND DH.ACTUAL_GI_DT >=  " & FilterVar(UNIConvDate(Trim(frm1.txtConFromDt.text)), "''", "S") & ""
		Else
			iStrWhereList = iStrWhereList & " AND DH.PROMISE_DT >=  " & FilterVar(UNIConvDate(Trim(frm1.txtConFromDt.text)), "''", "S") & ""
		End If
	End If
	
	If frm1.txtConToDt.text <> "1900-01-01" Then
		If frm1.txtConConfFlag.value = "Y" Then
			iStrWhereList = iStrWhereList & " AND DH.ACTUAL_GI_DT <=  " & FilterVar(UNIConvDate(Trim(frm1.txtConToDt.text)), "''", "S") & ""
		Else
			iStrWhereList = iStrWhereList & " AND DH.PROMISE_DT <=  " & FilterVar(UNIConvDate(Trim(frm1.txtConToDt.text)), "''", "S") & ""
		End If		
	End If
	
	If frm1.txtConDNNo.value <> "" Then
		iStrWhereList = iStrWhereList & " AND DH.DN_NO =  " & FilterVar(frm1.txtConDNNo.value, "''", "S") & ""
	End If
	
	If frm1.txtConShipToParty.value <> "" Then
		iStrWhereList = iStrWhereList & " AND DH.SHIP_TO_PARTY =  " & FilterVar(frm1.txtConShipToParty.value, "''", "S") & ""
	End If
	
	If frm1.txtConDnType.value <> "" Then
		iStrWhereList = iStrWhereList & " AND DH.MOV_TYPE =  " & FilterVar(frm1.txtConDnType.value, "''", "S") & ""
	End If
	
	If frm1.txtConConfFlag.value <> "" Then
		iStrWhereList = iStrWhereList & " AND DH.POST_FLAG =  " & FilterVar(frm1.txtConConfFlag.value , "''", "S") & ""
	End If	
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, PopupParent.gColSep)	
		frm1.txtDNTotAmtLoc.text = UNIFormatNumber(iArrRs(1), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
		frm1.txtDNAmtLoc.text = UNIFormatNumber(iArrRs(2), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
		frm1.txtVATAmtLoc.text = UNIFormatNumber(iArrRs(3), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
		frm1.txtDepoAmtLoc.text = UNIFormatNumber(iArrRs(4), ggAmtOfMoney.DecPoint, -2, 0, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit)
	Else
		If Err.number <> 0 Then
			MsgBox Err.description 
			Err.Clear
			Exit Sub
		End If
	End If
		
End Sub

'============================================
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	
		<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "RA") %>
		<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "RA") %>
	
End Sub
	
'============================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S4116PA4","S","A","V20030529", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
    Call SetSpreadLock        
End Sub

'=============================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspdData.OperationMode = 3
End Sub

'=============================================
Function CancelClick()
	Self.Close()
End Function

'=============================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format		
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field  

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
		 
End Sub

'=============================================
Function vspdData_KeyPress(KeyAscii)
   On Error Resume Next
   If KeyAscii = 27 Then
	  Call CancelClick()
   End If
End Function

'=============================================
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

'=============================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub

'=============================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DbQuery() 
    	End If
    End If
    
End Sub

'=============================================
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=============================================
Function DbQuery() 
	Dim strVal
    
    DbQuery = False
    
    Err.Clear                                                               

	If LayerShowHide(1) = False Then
      	Exit Function
    End If
    
    With frm1
    
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
		.txtMode.value = PopupParent.UID_M0001	
		.txtHlgPageNo.value	= lgPageNo			
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
        .lgSelectListDT.value = GetSQLSelectListDataType("A")
        .lgTailList.value = MakeSQLGroupOrderByList("A")
		.lgSelectList.value = EnCoding(GetSQLSelectList("A"))		
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    End With
    
    DbQuery = True
End Function

'=============================================
Function DbQueryOk()
	frm1.vspdData.Focus     													
	frm1.vspdData.SelModeSelected = True		
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
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
						<TD CLASS=TD5>출하번호</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConDNNo" SIZE=20 MAXLENGTH=18 TAG="14XXXU" ALT="수주번호"></TD>
						<TD CLASS="TD5" NOWRAP>출하형태</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConDnType" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="출하형태">
															<INPUT TYPE=TEXT NAME="txtConDnTypeNm" SIZE=20 tag="14" ALT="출하형태명"></TD>												
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>납품처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConShipToParty" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="납품처">
															<INPUT TYPE=TEXT NAME="txtConShipToPartyNm" SIZE=20 tag="14" ALT="납품처명"></TD>
						<TD CLASS="TD5" NOWRAP>출고일</TD>									
						<TD CLASS="TD6" NOWRAP>							        
						<TABLE CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD>
								<script language =javascript src='./js/s4116pa4_OBJECT1_txtConFromDt.js'></script>
								</TD>
								<TD>
								&nbsp;~&nbsp;
								</TD>
								<TD>
								<script language =javascript src='./js/s4116pa4_OBJECT2_txtConToDt.js'></script>
								</TD>
							</TR>										
						</TABLE>							        
						</TD>	
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>출고여부</TD>
						<TD CLASS=TD6 NOWRAP>										
							<input type=radio CLASS="RADIO" name="rdoConfFlag" id="rdoConf" value="Y" tag = "14X" checked>
								<label for="rdoConf">출고</label>&nbsp;&nbsp;
							<input type=radio CLASS = "RADIO" name="rdoConfFlag" id="rdoNonConf" value="N" tag = "14X">
								<label for="rdoNonConf">미출고</label>
						</TD>	
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP>	
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
					<TD CLASS=TD5 NOWRAP>출고총자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s4116pa4_fpDoubleSingle1_txtDNTotAmtLoc.js'></script>							
					</TD>
					<TD CLASS=TD5 NOWRAP>출고자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s4116pa4_fpDoubleSingle2_txtDNAmtLoc.js'></script>							
					</TD>
				</TR>	
				<TR>
					<TD CLASS=TD5 NOWRAP>VAT자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s4116pa4_fpDoubleSingle3_txtVATAmtLoc.js'></script>							
					</TD>
					<TD CLASS=TD5 NOWRAP>적립금자국금액</TD>
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s4116pa4_fpDoubleSingle4_txtDepoAmtLoc.js'></script>							
					</TD>
				</TR>
				<TR>
					<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
						<script language =javascript src='./js/s4116pa4_vaSpread1_vspdData.js'></script>
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
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
					                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgPageNo"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="lgSelectListDT" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="lgTailList" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="lgSelectList" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtConConfFlag"    tag="14" TABINDEX="-1">

</FORM>	
		
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
