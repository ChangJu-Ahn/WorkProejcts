<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5211ra1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : B/L Reference ASP															*
'*  6. Comproxy List        : + B19029LookupNumericFormat												*
'*  7. Modified date(First) : 2000/04/30																*
'*  8. Modified date(Last)  : 2000/05/02																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : PARK NO YEOL																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<!--TITLE>B/L 참조</TITLE-->
<TITLE></TITLE>
<!--
'******************************************  1.1 Inc 선언   **********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'==========================================  1.1.1 Style Sheet  ======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
'==========================================  1.1.2 공통 Include   ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID 		= "m5211rb1_KO441.asp"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 9                                           '☆: key count of SpreadSheet
Const gstrPayTermsMajor = "B9004"
Const gstrIncotermsMajor = "B9006"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'========================================== 2.1.1 InitVariables()  ======================================
Function InitVariables()
		
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False
    Redim arrReturn(0)        
    Self.Returnvalue = arrReturn     

End Function
'==========================================  2.2.1 SetDefaultVal()  ====================================
Sub SetDefaultVal()

	frm1.txtIssueFromDt.text = StartDate
	frm1.txtIssueToDt.text = EndDate					

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrp, "Q") 
		frm1.txtPurGrp.Tag = left(frm1.txtPurGrp.Tag,1) & "4" & mid(frm1.txtPurGrp.Tag,3,len(frm1.txtPurGrp.Tag))
        frm1.txtPurGrp.value = lgPGCd
	End If
						
End Sub
'==========================================  2.2.2 LoadInfTB19029() =====================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("M5211RA1","S","A","V20030321",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 3 
End Sub

'============================================ 2.2.4 SetSpreadLock()  ====================================
Sub SetSpreadLock(ByVal pOpt)
    IF pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End IF
End Sub

'==========================================  2.3.1 OkClick()  ===========================================
Function OKClick()

	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		
		Redim arrReturn(frm1.vspdData.MaxCols - 2)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
		For intColCnt = 0 To frm1.vspdData.MaxCols - 2
			'frm1.vspdData.Col = intColCnt + 1
			frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
			arrReturn(intColCnt) = frm1.vspdData.Text
		Next	
					
	End If
		
	Self.Returnvalue = arrReturn
	Self.Close()
	
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================  OpenSortPopup()  ===========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
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
'=========================================  OpenConSItemDC()  ===========================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(6), arrField(5), arrHeader(5)


	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	
	Select Case iWhere
	
		Case 0						
			arrParam(0) = "결제방법"							    <%' 팝업 명칭 %>
			arrParam(1) = "b_minor,b_configuration"					<%' TABLE 명칭 %>
			arrParam(2) = Trim(frm1.txtPayTerms.Value)				<%' Code Condition%>
	'		arrParam(3) = Trim(txtBeneficiaryNm.value)				<%' Name Cindition%>
			arrParam(4) = "b_minor.Major_Cd= " & FilterVar(gstrPayTermsMajor, "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd AND b_configuration.REFERENCE = " & FilterVar("M", "''", "S") & " "
			arrParam(5) = "결제방법"								<%' TextBox 명칭 %>
			
		    arrField(0) = "b_minor.Minor_CD"									<%' Field명(0)%>
			arrField(1) = "b_minor.Minor_NM"									<%' Field명(1)%>

		    
		    arrHeader(0) = "결제방법"								<%' Header명(0)%>
		    arrHeader(1) = "결제방법명"							<%' Header명(1)%>
		    
		Case 1
			arrParam(0) = "가격조건"								<%' 팝업 명칭 %>
			arrParam(1) = "B_Minor"										<%' TABLE 명칭 %>
			arrParam(2) = Trim(frm1.txtIncoterms.Value)						<%' Code Condition%>
	'		arrParam(3) = Trim(txtIncotermsNm.Value)					<%' Name Cindition%>
			arrParam(4) = "MAJOR_CD= " & FilterVar(gstrIncotermsMajor, "''", "S") & ""		<%' Where Condition%>
			arrParam(5) = "가격조건"								<%' TextBox 명칭 %>

			arrField(0) = "Minor_CD"									<%' Field명(0)%>
			arrField(1) = "Minor_NM"									<%' Field명(1)%>

			arrHeader(0) = "가격조건"								<%' Header명(0)%>
			arrHeader(1) = "가격조건명"	

		Case 2
		    If frm1.txtPurGrp.className = "protected" Then Exit Function
		    
			arrParam(0) = "구매그룹"								<%' 팝업 명칭 %>
			arrParam(1) = "B_PUR_GRP"									<%' TABLE 명칭 %>
			arrParam(2) = Trim(frm1.txtPurGrp.Value)							<%' Code Condition%>
	'		arrParam(3) = Trim(txtPurGrpNm.Value)						<%' Name Cindition%>
			arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "								<%' Where Condition%>
			arrParam(5) = "구매그룹"								<%' TextBox 명칭 %>

			arrField(0) = "PUR_GRP"										<%' Field명(0)%>
			arrField(1) = "PUR_GRP_NM"									<%' Field명(1)%>
	 
			arrHeader(0) = "구매그룹"								<%' Header명(0)%>
			arrHeader(1) = "구매그룹명"		
	    
		Case 3
		    arrParam(0) = "수출자팝업"							<%' 팝업 명칭 %>
			arrParam(1) = "B_BIZ_PARTNER"			<%' TABLE 명칭 %>
			arrParam(2) = Trim(frm1.txtBeneficiary.value)				<%' Code Condition%>
	'		arrParam(3) = Trim(txtBeneficiaryNm.value)				<%' Name Cindition%>
			arrParam(4) = ""					<%' Where Condition%>
			arrParam(5) = "수출자"								<%' TextBox 명칭 %>
		
			arrField(0) = "BP_Cd"							<%' Field명(0)%>
			arrField(1) = "BP_NM"								<%' Field명(1)%>
	    
			arrHeader(0) = "수출자"								<%' Header명(0)%>
			arrHeader(1) = "수출자명"	
		
	End Select

	arrParam(0) = arrParam(5)												' 팝업 명칭	
	
	Select Case iWhere
	
	Case 0,1,2,3
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'-------------------------------------------------------------------------------------------------------
'	Name : SetConSItemDC()
'	Description : OpenConSItemDC Popup에서 Return되는 값 setting
'-------------------------------------------------------------------------------------------------------
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtPayTerms.Value = arrRet(0)
			.txtPayTermsNm.Value = arrRet(1)
			.txtPayTerms.focus	
		case 1
			.txtIncoterms.Value = arrRet(0)
			.txtIncotermsNm.Value = arrRet(1)
			.txtIncoterms.focus	
		Case 2						
			.txtPurGrp.Value = arrRet(0)
			.txtPurGrpNm.Value = arrRet(1)
			.txtPurGrp.focus	
		CASE 3							
			.txtBeneficiary.value = arrRet(0)
			.txtBeneficiaryNm.value = arrRet(1)
			.txtBeneficiary.focus	
		End Select
	End With
	Set gActiveElement = document.activeElement
End Function

'==========================================  3.1.1 Form_Load()  =========================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    Call InitVariables											  '⊙: Initializes local global variables
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    Call FncQuery()
    
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================  vspdData_Click()  ============================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	'Call SetPopupMenuItemInf("0001111111")		
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
End Sub
'=========================================  3.3.1 vspdData_DblClick()  ==================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
	End If
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'========================================  3.3.2 vspdData_KeyPress()  ===================================
    Function vspdData_KeyPress(KeyAscii)
         On Error Resume Next
         If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
            Call OKClick()
         ElseIf KeyAscii = 27 Then
            Call CancelClick()
         End If
    End Function

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
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
'   Event Name : OCX_DbClick()
'========================================================================================================
Sub txtIssueFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIssueFromDt.Action = 7	
		Call SetFocusToDocument("P")
		frm1.txtIssueFromDt.focus	
	End If
End Sub

Sub txtIssueToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtIssueToDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtIssueToDt.focus
	End If
End Sub

'=======================================================================================================
'   Event Name : OCX_KeyDown()
'=======================================================================================================
Sub txtIssueFromDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtIssueToDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	With frm1
		if (UniConvDateToYYYYMMDD(.txtIssueFromDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtIssueToDt.text,PopupParent.gDateFormat,"")) and Trim(.txtIssueFromDt.text)<>"" and Trim(.txtIssueToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","B/L접수일", "X")			
			Exit Function
		End if   
    End with

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'⊙: Initializes local global variables
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    Set gActiveElement = document.activeElement
End Function

'========================================================================================================
' Function Name : DbQuery
'========================================================================================================
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
		   
		    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&txtBeneficiary = " & Trim(.hdnBeneficiary.value)	<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtPurGrp=" & Trim(.hdnPurGrp.value)
			strVal = strVal & "&txtBLDocNo=" & Trim(.hdnBLDocNo.value)
			strVal = strVal & "&txtIssueFromDt=" & Trim(.hdnIssueFromDt.value)
			strVal = strVal & "&txtIssueToDt=" & Trim(.hdnIssueToDt.value)
			strVal = strVal & "&txtPayTerms=" & Trim(.hdnPayTerms.value)
			strVal = strVal & "&txtIncoterms=" & Trim(.hdnIncoterms.value)			
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001		
			strVal = strVal & "&txtBeneficiary=" & Trim(.txtBeneficiary.value)	<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtPurGrp=" & Trim(.txtPurGrp.value)
			strVal = strVal & "&txtBLDocNo=" & Trim(.txtBLDocNo.value)
			strVal = strVal & "&txtIssueFromDt=" & Trim(.txtIssueFromDt.Text)
			strVal = strVal & "&txtIssueToDt=" & Trim(.txtIssueToDt.Text)
			strVal = strVal & "&txtPayTerms=" & Trim(.txtPayTerms.value)
			strVal = strVal & "&txtIncoterms=" & Trim(.txtIncoterms.value)
		End If				
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

        Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    

End Function
'=========================================================================================================
' Function Name : DbQueryOk
'=========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtBLDocNo.focus
	End If

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
							<TD CLASS=TD5 NOWRAP>수출자</TD>
						    <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="수출자" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 3">
											    <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>구매그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10  MAXLENGTH=4 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON" ONCLICK="VBSCRIPT:OpenConSItemDC 2">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>B/L번호</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBLDocNo" TYPE=TEXT MAXLENGTH=35 SIZE=20  TAG="11XXXU"></TD><!--<IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBLDocNo" align=top TYPE="BUTTON"></TD>-->
							<TD CLASS=TD5 NOWRAP>B/L접수일</TD>
							<TD CLASS=TD6 NOWRAP>
								<TABLE CELLSPACING=0 CELLPADDING=0>
									<TR>
										<TD>
											<script language =javascript src='./js/m5211ra1_fpDateTime1_txtIssueFromDt.js'></script>
										</TD>
										<TD>
											~
										</TD>
										<TD>
											<script language =javascript src='./js/m5211ra1_fpDateTime2_txtIssueToDt.js'></script>
										</TD>
									</TR>
								</TABLE>
							</TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>결제방법</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10  MAXLENGTH=5 TAG="11XXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK="VBSCRIPT:OpenConSItemDC 0">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>가격조건</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIncoterms" TYPE="Text" MAXLENGTH="5" SIZE=10 STYLE=" Text-Transform: uppercase" tag="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncoterms" align=top TYPE="BUTTON" ONCLICK="VBSCRIPT:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 TAG="14"></TD>
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
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
							<script language =javascript src='./js/m5211ra1_vaSpread2_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
						                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

	<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
    <INPUT TYPE=HIDDEN NAME="hdnBeneficiary" tag="24">
	<INPUT TYPE=HIDDEN NAME="hdnPurGrp" tag="24">
	<INPUT TYPE=HIDDEN NAME="hdnBLDocNo" tag="24">
	<INPUT TYPE=HIDDEN NAME="hdnIssueFromDt" tag="24">
	<INPUT TYPE=HIDDEN NAME="hdnIssueToDt" tag="24">
	<INPUT TYPE=HIDDEN NAME="hdnPayTerms" tag="24">
	<INPUT TYPE=HIDDEN NAME="hdnIncoterms" tag="24">
	
	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
</BODY>
</HTML>
