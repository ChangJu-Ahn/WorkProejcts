<%@ LANGUAGE="VBSCRIPT" %>
<!--
<%
'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5212ra1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : B/L내역참조 PopUp ASP														*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/06/03																*
'*  9. Modifier (First)     : Sun-joung Lee																*
'* 10. Modifier (Last)      : Lee Eun Hee																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
%>
-->
<HTML>
<HEAD>
<!--TITLE>B/L 내역참조</TITLE-->
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

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID 		= "m5212rb1.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 22                                           '☆: key count of SpreadSheet
Const C_LC_NO			= 1

<!-- #Include file="../../inc/lgvariables.inc" -->	


Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim IsOpenPop

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

'========================================== 2.1.1 InitVariables()  ======================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
    lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
		
	gblnWinEvent = False
	IsOpenPop = False
        
    Redim arrReturn(0, 0)   
	self.Returnvalue = arrReturn     

End Function

'==========================================  2.2.1 SetDefaultVal()  ====================================
Sub SetDefaultVal()
	Err.Clear
	
	'통관번호(Hdr)를 조회조건으로 넘겨받는다 %>
	frm1.txtCCNo.value = arrParam(0)
End Sub

'==========================================  2.2.2 LoadInfTB19029() =====================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("m5212ra1","S","A","V20030320",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 5 
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

	Dim intColCnt, intRowCnt, intInsRow
		
	if frm1.vspdData.SelModeSelCount > 0 then
			
		intInsRow = 0
			
		Redim arrReturn(frm1.vspdData.SelModeSelCount,frm1.vspdData.MaxCols - 2)
			
		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1
				
			frm1.vspdData.Row = intRowCnt + 1
				
			if frm1.vspdData.SelModeSelected Then
				For intColCnt = 0 To frm1.vspdData.MaxCols - 2
					'frm1.vspdData.Col = intColCnt + 1
					frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
					arrReturn(intInsRow,intColCnt) = frm1.vspdData.Text
				Next
					
				intInsRow = intInsRow + 1
			End If
		Next
			
	End IF
		 
	Self.Returnvalue = arrReturn

	Self.Close()
	
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Self.Close()
End Function
'==========================================  OpenSortPopup()  ===========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	'If lgIsOpenPop = True Then Exit Function
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
'==========================================  OpenItemPop()  ===========================================
Function OpenItemPop()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "품목"							
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtItemCode.value)						
'		arrParam(3) = Trim(txtItemName.value)						
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "			
	arrParam(5) = "품목"									
	
	arrField(0) = "B_Item.Item_Cd"				
	arrField(1) = "B_Item.Item_NM"		
    
	arrHeader(0) = "품목"								
	arrHeader(1) = "품목명"									

	iCalledAspName = AskPRAspName("M1111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M1111PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtItemCode.focus
		Exit Function
	Else
		frm1.txtItemCode.value = arrRet(0)
		frm1.txtItemName.value = arrRet(1)
		frm1.txtItemCode.focus
		Set gActiveElement = document.activeElement
	End If	
End Function
'===============================  OpenTrackingNo()  ============================
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
    arrParam(2) = ""	'공장 
    arrParam(3) = ""	'모품목 
    arrParam(4) = ""	'수주번호 
    arrParam(5) = ""	'추가 Where절 
    
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	IsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		frm1.txtTrackingNo.focus
		lgBlnFlgChgValue = True
		Set gActiveElement = document.activeElement
	End If	

End Function
'******************************************  2.5.1 CcHdrQuery()  ****************************************
Function CcHdrQuery()
	Dim strVal
		
	Err.Clear															<%'☜: Protect system from crashing%>

	txtBlDocNo.value = ""

	CcHdrQuery = False													<%'⊙: Processing is NG%>

	Call LayerShowHide(1)

	strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0002					<%'☜: 비지니스 처리 ASP의 상태 %>
	strVal = strVal & "&txtCcNo=" & Trim(frm1.txtCcNo.value)					<%'☆: 조회 조건 데이타 %>

	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

	CcHdrQuery = True													<%'⊙: Processing is NG%>
End Function

'==========================================  3.1.1 Form_Load()  =========================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
	
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================   vspdData_Click()  =======================================
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

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'⊙: Initializes local global variables

	If DbQuery = False Then Exit Function									

    FncQuery = True		
    Set gActiveElement = document.activeElement
End Function

'==================================  DbQuery()  ======================================================
Function DbQuery() 
	
	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
	     
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtCcNo=" & Trim(.txtCcNo.value)					<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtItemCode=" & Trim(.txtHItemCode.value)			<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtPurGrp=" & Trim(.txtHPurGrp.value)
			strVal = strVal & "&txtBeneficiary=" & Trim(.txtHBeneficiary.value)
			strVal = strVal & "&txtCurrency=" & Trim(.txtHCurrency.value)
			strVal = strVal & "&txtPayTerms=" & Trim(.txtHPayTerms.value)
			strVal = strVal & "&txtIncoterms=" & Trim(.txtHIncoterms.value)
			strVal = strVal & "&txtBlDocNo=" & Trim(.txtBlDocNo.value)
		Else
		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
			strVal = strVal & "&txtCcNo=" & Trim(.txtCcNo.value)					<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtItemCode=" & Trim(.txtItemCode.value)			<%'☆: 조회 조건 데이타 %>
			strVal = strVal & "&txtPurGrp=" & Trim(.txtPurGrp.value)
			strVal = strVal & "&txtBeneficiary=" & Trim(.txtBeneficiary.value)
			strVal = strVal & "&txtCurrency=" & Trim(.txtCurrency.value)
			strVal = strVal & "&txtPayTerms=" & Trim(.txtPayTerms.value)
			strVal = strVal & "&txtIncoterms=" & Trim(.txtIncoterms.value)
			strVal = strVal & "&txtBlDocNo=" & Trim(.txtBlDocNo.value)
		End If
			strVal = strVal & "&txtTrackingNo=" &Trim(.txtTrackingNo.value)	
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    
	
End Function

'===================================  DbQueryOk()  =======================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.vspdData.focus
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
							<TD CLASS=TD5 NOWRAP>품목코드</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemCode" ALT="품목코드" TYPE="Text" MAXLENGTH="18" SIZE=10 STYLE=" Text-Transform: uppercase" tag="11"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemPop()">&nbsp;<INPUT NAME="txtItemName" TYPE="Text" MAXLENGTH="40" SIZE=25 STYLE=" Text-Transform: uppercase" tag="14"></TD>
							<TD CLASS=TD5 NOWRAP>구매그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrp" SIZE=10 MAXLENGTH=4 TAG="14XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 TAG="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>수출자</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10 MAXLENGTH=10 TAG="14XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
							<TD CLASS="TD5" NOWRAP>화폐단위</TD>
							<TD CLASS="TD6"><INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>결제방법</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10  MAXLENGTH=5 TAG="14XXXU" ALT="결제방법">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>가격조건</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIncoterms" TYPE="Text" MAXLENGTH="5" SIZE=10 STYLE=" Text-Transform: uppercase" tag="14">
												 <INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 TAG="14"></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
							<TD CLASS="TD6" NOWRAP><INPUT NAME="txtTrackingNo" ALT="Tracking번호" TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP></TD>						
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
							<script language =javascript src='./js/m5212ra1_vaSpread_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="txtMode" tag="14">
<INPUT TYPE=HIDDEN NAME="txtBlDocNo" tag="14">
<INPUT TYPE=HIDDEN NAME="txtCCNo" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHItemCode" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHPurGrp" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHBeneficiary" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHIncoterms" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHBlDocNo" tag="14">
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
