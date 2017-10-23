<%@ LANGUAGE="VBSCRIPT" %>
<!--
<%
'************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 발주참조 Popup
'*  3. Program ID           : M3111RA1
'*  4. Program Name         : P/O Reference ASP
'*  5. Program Desc         : ADO Query
'*  6. Component List       :																			*
'*  7. Modified date(First) : 2000/04/29
'*  8. Modified date(Last)  : 2002/05/11
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/04 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/17 : ADO변환 
'**************************************************************************************
%>
-->
<HTML>
<HEAD>
<TITLE>발주참조</TITLE>
<!--
'********************************************  1.1 Inc 선언  ********************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
<!--
'============================================  1.1.2 공통 Include  ======================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					<% '☜: indicates that All variables must be declared in advance %>
	
Const BIZ_PGM_ID 		= "m3111rb1.asp"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 9                                           '☆: key count of SpreadSheet
Const C_PoNo			= 1											  '☆: Spread Sheet 의 Columns 인덱스 
Const C_RateOp			= 7
'------ Minor Code PopUp을 위한 Major Code정의 ------ 
Const gstrPayTermsMajor = "B9004"					'결제방법 
Const gstrPOTypeMajor	= "M3101"					'발주형태 


<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgSelectList                                            '☜: SpreadSheet의 초기  위치정보관련 변수 
Dim lgSelectListDT                                          '☜: SpreadSheet의 초기  위치정보관련 변수 

Dim lgSortFieldNm                                           '☜: Orderby popup용 데이타(필드설명)      
Dim lgSortFieldCD                                           '☜: Orderby popup용 데이타(필드코드)      

Dim lgPopUpR                                                '☜: Orderby default 값                    

Dim lgKeyPos                                                '☜: Key위치                               
Dim lgKeyPosVal                                             '☜: Key위치 Value                         
Dim IscookieSplit 

Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam
Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)

top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'========================================== 2.1.1 InitVariables()  ======================================
Function InitVariables()
	Dim arrParent
		
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
	gblnWinEvent = False        
End Function
'==========================================  2.2.1 SetDefaultVal()  ====================================
Sub SetDefaultVal()
	Dim arrTemp		
	Dim strReturn
	Redim strReturn(1) 
						
    Self.Returnvalue = strReturn     
	frm1.txtBeneficiary.focus	

	frm1.txtFrDt.Text = StartDate
	frm1.txtToDt.Text = EndDate
	frm1.vspdData.OperationMode = 3
	
End Sub

'==========================================  2.2.2 LoadInfTB19029() =====================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "RA") %>                                '☆: 

End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M3111RA1","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
	Call SetSpreadLock 	    
End Sub


'============================================ 2.2.4 SetSpreadLock()  ====================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	
'==========================================  2.3.1 OkClick()  ===========================================
Function OKClick()
	Dim strReturn
	
	With frm1.vspdData 
		If .ActiveRow > 0 Then	
			Redim strReturn(1)
		
			.Row = .ActiveRow
			.Col = GetKeyPos("A",C_PoNo)
			strReturn(0) = Trim(.Text)

			.Col = GetKeyPos("A",C_RateOp)
			strReturn(1) = Trim(.Text)

			Self.Returnvalue = strReturn
		End If
	End With
		
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
	Redim arrReturn(1)
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================================================================================
' Function Name : OpenConSItemDC
' Function Desc : OpenConSItemDC Reference Popup
'========================================================================================================
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	
	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"							' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtBeneficiary.value)				' Code Condition
		arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "' Where Condition
		arrParam(5) = "수출자"								' TextBox 명칭 
	
		arrField(0) = "BP_CD"									' Field명(0)
		arrField(1) = "BP_NM"									' Field명(1)
    
		arrHeader(0) = "수출자"								' Header명(0)
		arrHeader(1) = "수출자명"							' Header명(1)
    
	Case 1
		arrParam(1) = "B_Pur_Grp"
		arrParam(2) = Trim(frm1.txtGroupCd.Value)
		arrParam(4) = ""
		arrParam(5) = "구매그룹"			
	
		arrField(0) = "PUR_GRP"	
		arrField(1) = "PUR_GRP_NM"	
    
		arrHeader(0) = "구매그룹"		
		arrHeader(1) = "구매그룹명"
    
	Case 2
		arrParam(1) = "m_config_process"							' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtPOType.Value)							' Code Condition
		arrParam(4) = ""											' Where Condition
		arrParam(5) = "발주형태"								' TextBox 명칭 

		arrField(0) = "PO_TYPE_CD"									' Field명(0)
		arrField(1) = "PO_TYPE_NM"									' Field명(1)

		arrHeader(0) = "발주형태"								' Header명(0)
		arrHeader(1) = "발주형태명"								' Header명(1)

	Case 3
		
		arrParam(1) = "b_minor,b_configuration"						' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtPayTerms.Value)						' Code Condition
		arrParam(4) = "b_minor.Major_Cd= " & FilterVar(gstrPayTermsMajor, "''", "S") & " and b_minor.minor_cd=b_configuration.minor_cd AND b_configuration.REFERENCE = " & FilterVar("M", "''", "S") & " "
		arrParam(5) = "결제방법"								' TextBox 명칭 

		arrField(0) = "b_minor.Minor_CD"							' Field명(0)
		arrField(1) = "b_minor.Minor_NM"							' Field명(1)

		arrHeader(0) = "결제방법"								' Header명(0)
		arrHeader(1) = "결제방법명"								' Header명(1)
				
	 Case 4			
		arrParam(1) = "B_Minor"										' TABLE 명칭 
		arrParam(2) = Trim(frm1.txtIncoterms.Value)						' Code Condition
		arrParam(3) = ""											' Name Cindition
		arrParam(4) = "MAJOR_CD=" & FilterVar("B9006", "''", "S") & ""							' Where Condition
		arrParam(5) = "가격조건"								' TextBox 명칭 

		arrField(0) = "Minor_CD"									' Field명(0)
		arrField(1) = "Minor_NM"									' Field명(1)

		arrHeader(0) = "가격조건"								' Header명(0)
		arrHeader(1) = "가격조건명"								' Header명(1)

		
	End Select
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
		arrParam(0) = arrParam(5)												' 팝업 명칭	

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
			.txtBeneficiary.value = arrRet(0)
			.txtBeneficiaryNm.value = arrRet(1)	 
			.txtBeneficiary.focus  
		Case 1
			.txtGroupCd.Value = arrRet(0)
			.txtGroupNm.Value = arrRet(1)
			.txtGroupCd.focus
		Case 2
			.txtPOType.Value = arrRet(0)
			.txtPOTypeNm.Value = arrRet(1) 
			.txtPOType.focus
		Case 3
			.txtPayTerms.Value = arrRet(0)
			.txtPayTermsNm.Value = arrRet(1)
			.txtPayTerms.focus		 
		Case 4
			.txtIncoterms.Value = arrRet(0)
			.txtIncotermsNm.Value = arrRet(1)
			.txtIncoterms.focus 		 	
	End Select
	End With
	Set gActiveElement = document.activeElement
End Function
'==========================================  3.1.1 Form_Load()  =========================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call fncQuery
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'================================== vspdData_KeyPress() ==========================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function
	
'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtFrDt_Keypress(KeyAscii)
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
'   Event Name : OCX_DbClick()
'   Event Desc : OCX_DbClick() 시 Calendar Popup
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
       frm1.txtFrDt.Action = 7
       Call SetFocusToDocument("P")
       frm1.txtFrDt.Focus
    End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
       frm1.txtToDt.Action = 7                                    ' 7 : Popup Calendar ocx
       Call SetFocusToDocument("P")
       frm1.txtToDt.Focus
    End If
End Sub

'=========================================  3.3.1 vspdData_DblClick()  ==================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
          Exit Function
	End If
	With frm1.vspdData
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
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
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* %>

Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	if (UniConvDateToYYYYMMDD(frm1.txtFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToDt.text,gDateFormat,"")) and Trim(frm1.txtFrDt.text)<>"" and Trim(frm1.txtToDt.text)<>"" then	
		Call DisplayMsgBox("17a003", "X","발주일", "X")			
		frm1.txtToDt.Focus
		Exit Function
	End if   

	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
 
    Call InitVariables 														'⊙: Initializes local global variables

	If DbQuery = False Then Exit Function									

    FncQuery = True	
    Set gActiveElement = document.activeElement	
End Function	

'********************************************  5.1 DbQuery()  *******************************************
Function DbQuery() 
	Dim strVal

	Err.Clear															<%'☜: Protect system from crashing%>
	DbQuery = False														<%'⊙: Processing is NG%>

	If LayerShowHide(1) =false then
	    Exit Function
	End if
	
	With frm1
			
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtBeneficiary=" & Trim(.txtHBeneficiary.value)	<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtPOType=" & Trim(.txtHPOType.value)
		strVal = strVal & "&txtPayTerms=" & Trim(.txtHPayTerms.value)
		strVal = strVal & "&txtFrDt=" & Trim(.txtHFrDt.Value)
		strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.Value)
		strVal = strVal & "&txtGroup=" & Trim(.txtHGrp.Value)
		strVal = strVal & "&txtIncoterms=" & Trim(.txtHIncoterms.Value)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		strVal = strVal & "&txtBeneficiary=" & Trim(.txtBeneficiary.value)	<%'☆: 조회 조건 데이타 %>
		strVal = strVal & "&txtPOType=" & Trim(.txtPOType.value)
		strVal = strVal & "&txtPayTerms=" & Trim(.txtPayTerms.value)
		strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
		strVal = strVal & "&txtGroup=" & Trim(.txtGroupCd.Value)
		strVal = strVal & "&txtIncoterms=" & Trim(.txtIncoterms.Value)
	End if

	End With
			
		strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
		DbQuery = True		
End Function	

<%
'=============================================  5.2.2 DbQueryOk()  ======================================
%>
Function DbQueryOk()
	lgIntFlgMode = PopupParent.OPMD_UMODE
	With frm1.vspdData
		If .MaxRows > 0 Then
			.Focus
			.Row = 1	
			.SelModeSelected = True		
		Else
			.Focus
		End If
	End With		
End Function


'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderByPopup()
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
				<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
					<TR>
						<TD CLASS=TD5 NOWRAP>수출자</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBeneficiary" SIZE=10  MAXLENGTH=10 TAG="11XXXU" ALT="수출자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBeneficiary" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 0" >
											 <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 TAG="14"></TD>
						<TD CLASS="TD5" NOWRAP>구매그룹</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4  tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSItemDC 1">
											   <INPUT TYPE=TEXT Alt="구매그룹" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>발주형태</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPOType" SIZE=10  MAXLENGTH=5 TAG="1XXXXU" ALT="발주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPOType" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 2">
											 <INPUT TYPE=TEXT NAME="txtPOTypeNm" SIZE=20 TAG="14"></TD>
						<TD CLASS="TD5" NOWRAP>발주일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td>
										<script language =javascript src='./js/m3111ra1_fpDateTime1_txtFrDt.js'></script>
									</td>
									<td>~</td>
									<td>
										<script language =javascript src='./js/m3111ra1_fpDateTime2_txtToDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>결제방법</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPayTerms" SIZE=10  MAXLENGTH=5 TAG="1XNXXU" ALT="결제방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" onclick="vbscript:OpenConSItemDC 3">
											 <INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD5 NOWRAP>가격조건</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIncoterms" SIZE=10  MAXLENGTH=5 TAG="1XNXXU" ALT="가격조건"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIncoterms" align=top TYPE="BUTTON" ONCLICK="VBScript:OpenConSItemDC 4">
											 <INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 TAG="14"></TD>
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
						<script language =javascript src='./js/m3111ra1_vaSpread1_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP><IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
										<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderByPopup()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
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
<INPUT TYPE=HIDDEN NAME="txtHBeneficiary" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHPOType" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHFrDt" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHIncoterms" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHGrp" TAG="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
