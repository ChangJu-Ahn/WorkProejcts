<%@ LANGUAGE="VBSCRIPT" %>
<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S4211PA1
'*  4. Program Name         : 통관관리번호 팝업 
'*  5. Program Desc         : 통관관리번호 팝업 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : kim hyung suk
'* 10. Modifier (Last)      : Seo jin kyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**************************************************************************************
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
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" -->	
'==========================================================================================================

Dim lgIsOpenPop                                              

Dim arrParam	
Dim gblnWinEvent			

Dim strReturn										   

Dim arrParent					
arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s4211pb1.asp"
Const C_MaxKey          = 2                                    

Const gstrLCTypeMajor = "S9000"					
'==========================================================================================================
Sub InitVariables()
         
    lgBlnFlgChgValue = False                               
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1
    lgPageNo         = ""
	lgIntFlgMode = PopupParent.OPMD_CMODE	
	
	ReDim strReturn(0)
	strReturn(0) = ""
	gblnWinEvent = False
	Self.Returnvalue = strReturn
End Sub
'==========================================================================================================
Sub SetDefaultVal()
	frm1.txtFromDt.text = StartDate
	frm1.txtToDt.text = EndDate
End Sub
'==========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "PA") %>
End Sub
'==========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("s4211pa1","S","A","V20021203",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
    Call SetSpreadLock       
End Sub
'==========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================================
Function OKClick()
	ReDim strReturn(2)
		
		If frm1.vspdData.ActiveRow > 0 Then
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = GetKeyPos("A",1)		
			strReturn(0) = Trim(frm1.vspdData.Text)
				
			frm1.vspdData.Col = GetKeyPos("A",2)		
			strReturn(1) = Trim(frm1.vspdData.Text)

			Self.Returnvalue = strReturn
		End If
					
	Self.Close()
End Function
'==========================================================================================================
Function CancelClick()
	Self.Close()
End Function
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "수입자"							
	arrParam(1) = "B_BIZ_PARTNER"						
	arrParam(2) = Trim(frm1.txtApplicant.value)				
	arrParam(3) = ""									
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
	arrParam(5) = "수입자"							

	arrField(0) = "BP_CD"								
	arrField(1) = "BP_NM"								

	arrHeader(0) = "수입자"							
	arrHeader(1) = "수입자명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenSalesGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "영업그룹"						
	arrParam(1) = "B_SALES_GRP"							
	arrParam(2) = Trim(frm1.txtSalesGroup.value)				
	arrParam(3) = ""									
	arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "						
	arrParam(5) = "영업그룹"						

	arrField(0) = "SALES_GRP"							
	arrField(1) = "SALES_GRP_NM"						

	arrHeader(0) = "영업그룹"						
	arrHeader(1) = "영업그룹명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetSalesGroup(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenIvNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "송장번호"						
	arrParam(1) = "S_CC_HDR"							
	arrParam(2) = Trim(frm1.txtIvNo.value)				
	arrParam(3) = ""									
	arrParam(4) = ""									
	arrParam(5) = "송장번호"						

	arrField(0) = "ED22" & PopupParent.gColSep & "IV_NO"				
	arrField(1) = "DD22" & PopupParent.gColSep & "IV_DT"			

	arrHeader(0) = "송장번호"						
	arrHeader(1) = "송장작성일"						
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetIvNo(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenEpNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "면허번호"						
	arrParam(1) = "S_CC_HDR"							
	arrParam(2) = Trim(frm1.txtEpNo.value)					
	arrParam(3) = ""									
	arrParam(4) = ""									
	arrParam(5) = "면허번호"						

	arrField(0) = "CONVERT(char(35),EP_NO)"				
	arrField(1) = "CONVERT(char(11),EP_DT)"				

	arrHeader(0) = "면허번호"						
	arrHeader(1) = "면허일"							

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetEpNo(arrRet)
	End If
End Function	
'========================================================================================================
Function SetBizPartner(arrRet)
	frm1.txtApplicant.Value = arrRet(0)
	frm1.txtApplicantNm.Value = arrRet(1)
	frm1.txtApplicant.focus
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetSalesGroup(arrRet)
	frm1.txtSalesGroup.value = arrRet(0)
	frm1.txtSalesGroupNm.value = arrRet(1)
	frm1.txtSalesGroup.focus
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetIvNo(arrRet)
	frm1.txtIvNo.Value = arrRet(0)
	frm1.txtIvNo.focus
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetEpNo(arrRet)
	frm1.txtEpNo.Value = arrRet(0)
	frm1.txtEpNo.focus
End Function	
'=========================================================================================================

Sub Form_Load()
	Call LoadInfTB19029			
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")       	

	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    
	Call InitVariables			    
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'********************************************************************************************************
Sub btnApplicantOnClick()
	Call OpenBizPartner()
End Sub
'========================================================================================================
Sub btnSalesGroupOnClick()
	Call OpenSalesGroup()
End Sub
'========================================================================================================
Sub btnCurrencyOnClick()
	Call OpenIvNo()
End Sub
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7 
        Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
    End If
End Sub
'======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
    End If
End Sub
'=======================================================================================================
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
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)

	If Row = 0 Then Exit Function

	If frm1.vspdData.MaxRows = 0 Then Exit Function

	If Row > 0 Then Call OKClick()

End Function
'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If
		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
				DbQuery
			End If
		End If
	End With
End Sub
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
      	Exit Sub
    End If
    If CheckRunningBizProcess = True Then
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
'*********************************************************************************************************

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
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

    Call InitVariables 														

	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
			
    Call DbQuery									

    FncQuery = True		
End Function
'*********************************************************************************************************
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
			strVal = strVal & "&txtIvNo=" & Trim(frm1.txtHIVNo.value)
			strVal = strVal & "&txtFromDt=" & Trim(frm1.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(frm1.txtHToDt.value)
	    Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							
		    strVal = strVal & "&txtApplicant=" & Trim(frm1.txtApplicant.value)	
			strVal = strVal & "&txtSalesGroup=" & Trim(frm1.txtSalesGroup.value)
			strVal = strVal & "&txtIvNo=" & Trim(frm1.txtIVNo.value)
			strVal = strVal & "&txtFromDt=" & Trim(frm1.txtFromDt.text)
			strVal = strVal & "&txtToDt=" & Trim(frm1.txtToDt.text)
       End if   
			strVal = strVal & "&lgPageNo="       & lgPageNo                			
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

       Call RunMyBizASP(MyBizASP, strVal)										
			
	End With

	DbQuery = True

End Function
'========================================================================================
Function DbQueryOk()														

	If frm1.vspdData.MaxRows > 0 Then
		lgIntFlgMode = PopupParent.OPMD_UMODE
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtApplicant.focus
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
   End If
End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>수입자</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="수입자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnApplicantOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD5 NOWRAP>영업그룹</TD>
						<TD CLASS=TD6 NOWRAP>
						<INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnSalesGroupOnClick()">&nbsp;
					    <INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14">
					    </TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>송장번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=30 MAXLENGTH=35 TAG="11XXXU" ALT="송장번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvNo" align=top TYPE="BUTTON" OnClick="vbscript:btnCurrencyOnClick()"></TD>
						<TD CLASS=TD5 NOWRAP>작성일</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/s4211pa1_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
							<script language =javascript src='./js/s4211pa1_fpDateTime2_txtToDt.js'></script>
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
			<TABLE WIDTH="100%" HEIGHT="100%">
				<TR>
					<TD HEIGHT="100%" NOWRAP>
						<script language =javascript src='./js/s4211pa1_vaSpread_vspdData.js'></script>
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
						<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>
						<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>						
					<TD WIDTH=30% ALIGN=RIGHT>
						<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>
						<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHIVNo" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
