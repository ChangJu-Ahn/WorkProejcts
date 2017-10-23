<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m2111ra1
'*  4. Program Name         : 구매요청참조 
'*  5. Program Desc         : 구매요청참조 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/03/21	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Shin jin hyun		
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!--'☆: 해당 위치에 따라 달라짐, 상대 경로 -->
<!--
'============================================  1.1.2 공통 Include  ======================================
'========================================================================================================
-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>

<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<Script Language="VBS">
Option Explicit					 '☜: indicates that All variables must be declared in advance 
	

'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================

Const C_ReqNo 			= 1
Const C_PlantCd 		= 2															'☆: Spread Sheet의 Column별 상수 
Const C_PlantNm 		= 3
Const C_ItemCd 			= 4
Const C_ItemNm			= 5
Const C_Spec			= 6
Const C_Qty 			= 7
Const C_Unit 			= 8
Const C_DlvyDt 			= 9
Const C_PlantDt 		= 10	
Const C_ReqType			= 11
Const C_ReqTypeNm		= 12
Const C_SoNo			= 13	
Const C_SoSeqNo			= 14
Const C_TrackingNo		= 15	
Const C_SLCd 			= 16
Const C_SLNm 			= 17
Const C_HSCd			= 18
Const C_HSNm 			= 19	
	
Const BIZ_PGM_ID 		= "m2111rb2.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 19                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop  
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

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()
		lgStrPrevKey     = ""								   'initializes Previous Key
		lgPageNo         = ""
        lgBlnFlgChgValue = False	                           'Indicates that no value changed
        lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
        lgSortKey        = 1   
        
        lgIntGrpCount = 0										'⊙: Initializes Group View Size

        Redim arrReturn(0,0)        
        Self.Returnvalue = arrReturn     
End Function

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>                                '☆: 
End Sub

'==========================================  2.2.3 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("M2111RA1_2","S","A","V20030303",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock 
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub	

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt, intRowCnt, intInsRow
	with frm1
	If .vspdData.SelModeSelCount > 0 Then 
			
		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols-2)
			
		For intRowCnt = 1 To frm1.vspdData.MaxRows

			frm1.vspdData.Row = intRowCnt
				
			If frm1.vspdData.SelModeSelected Then
				For intColCnt = 0 To frm1.vspdData.MaxCols - 2
					frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
				Next
				intInsRow = intInsRow + 1
			End IF								
		Next
	End if		

	end with
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'===========================================================================
' Function Name : OpenSoNo
' Function Desc : OpenSoNo Reference Popup
'===========================================================================
Function OpenSoNo()
   Dim strRet
   Dim iCalledAspName
   Dim IntRetCD
	
   If IsOpenPop = True Then Exit Function
			
   IsOpenPop = True
		
   iCalledAspName = AskPRAspName("S3111PA1")
	
   If Trim(iCalledAspName) = "" Then
   	IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3111PA1", "X")
   	IsOpenPop = False
   	Exit Function
   End If
	
   strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,""), _
   	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

   IsOpenPop = False

   If strRet = "" Then
		frm1.txtSoNo.focus
   		Exit Function
   Else
   		frm1.txtSoNo.value = strRet
		frm1.txtSoNo.focus
   End If	
End Function

'===========================================================================
' Function Name : OpenTrackingNo
' Function Desc : OpenTrackingNo Reference Popup
'===========================================================================
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
	End If	
End Function
 
'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderBy()
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

'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

Sub SetDefaultVal()
	Dim arrParam
		
	arrParam = arrParent(1)
		
	frm1.vspdData.OperationMode = 5
	frm1.hdnSupplierCd.value 	= arrParam(6)
	frm1.hdnGroupCd.value 		= arrParam(2)
	frm1.hdnSubcontraflg.value 	= arrParam(4)
		
	If ubound(arrParam) >= 5 then		'2002-12-04(LJT)
		frm1.hdnSTOflg.value = arrParam(5)
	Else 
		frm1.hdnSTOflg.value = "N"
	End If

	If ubound(arrParam) >= 6 then		
		frm1.hdnPlantCd.value = arrParam(6)
	End If 
		
	frm1.txtFrDlvyDt.text 	= EndDate
	frm1.txtToDlvyDt.text 	= UnIDateAdd("m", +1, EndDate, PopupParent.gDateFormat)
End Sub


Sub txtFrDlvyDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToDlvyDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub


Sub txtFrDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDlvyDt.Action = 7
	    Call SetFocusToDocument("P")  
        frm1.txtFrDlvyDt.Focus
	End if
End Sub

Sub txtToDlvyDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDlvyDt.Action = 7
	    Call SetFocusToDocument("P")  
        frm1.txtToDlvyDt.Focus
	End if
End Sub

Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
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


Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
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

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
		
	If ValidDateCheck(frm1.txtFrDlvyDt, frm1.txtToDlvyDt) = False Then Exit Function
   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
End Function
	

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'☜: 비지니스 처리 ASP의 상태							
			strVal = strVal & "&txtFrDlvyDt=" & .hdnFrDt2.value
			strVal = strVal & "&txtToDlvyDt=" & .hdnToDt2.value		
			strVal = strVal & "&txtSoNo=" & .hdnSoNo.value
			strVal = strVal & "&txtTrackingNo=" & .hdnTrackingNo.value		
			strVal = strVal & "&txtSupplier=" & .hdnSupplierCd.value
			strVal = strVal & "&txtGroup=" & .hdnGroupCd.value
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey     
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001						
			strVal = strVal & "&txtFrDlvyDt=" & Trim(.txtFrDlvyDt.text)
			strVal = strVal & "&txtToDlvyDt=" & Trim(.txtToDlvyDt.text)
			strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)
			strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
			strVal = strVal & "&txtSupplier=" & .hdnSupplierCd.value
			strVal = strVal & "&txtGroup=" & .hdnGroupCd.value
			strVal = strVal & "&txtSubconfraflg=" & .hdnSubcontraflg.value
			strVal = strVal & "&txtSTOflg=" & .hdnSTOflg.value				'2002-12-04(LJT)
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If				
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    
End Function

'=========================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'=========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
	    Call SetFocusToDocument("P")  		
	End If
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'########################################################################################################
'#						6. TAG 부																		#
'########################################################################################################
-->
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
						<TD CLASS="TD5" NOWRAP>필요일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m2111ra2_fpDateTime2_txtFrDlvyDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m2111ra2_fpDateTime2_txtToDlvyDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>						
						<TD CLASS="TD5" NOWRAP>Tracking번호</TD>
						<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking번호" TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>수주번호</TD>
						<TD CLASS="TD6"><INPUT NAME="txtSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=26 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSoNo"></TD>
						<TD CLASS="TD5" NOWRAP>
						<TD CLASS="TD6" NOWRAP>
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
						<script language =javascript src='./js/m2111ra2_vaSpread1_vspdData.js'></script>
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
					<IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  ONCLICK="OpenOrderBy()"></IMG></TD>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% SRC="../../blank.htm" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnFrDt2" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt2" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnTrackingNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSTOflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     