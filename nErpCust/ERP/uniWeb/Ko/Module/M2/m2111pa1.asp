<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : M2111PA1
'*  4. Program Name         : 구매요청번호 
'*  5. Program Desc         : 구매요청번호 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/04/28
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)     : Kim Jae Soon
'* 10. Modifier (Last)      : KANG SU HWAN
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
<!--
'********************************************  1.1 Inc 선언  ********************************************
'*	Description : Inc. Include																			*
'********************************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================
-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>
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
<Script Language="VBScript">

Option Explicit					<% '☜: indicates that All variables must be declared in advance %>

'****************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************%>
Const BIZ_PGM_ID        = "M2111pb1.asp"                       ' 비지니스 로직 페이지 지정 
'========================================================================================================
'=									1.2 Constant variables 
'========================================================================================================
Const C_MaxKey  = 12                                           '☆: key count of SpreadSheet
Const C_No 		= 1		
'========================================================================================================
'=									1.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=									1.4 User-defind Variables
'========================================================================================================
Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim IsOpenPop  

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName

'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------
'==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= %>
'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- %>
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>
Dim arrValue(3)                    ' Popup되는 창으로 넘길때 인수를 배열로 넘김 
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Sub InitVariables()
	Dim arrParent
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                         'Indicates that current mode is Create mode
    IsOpenPop = False
    gblnWinEvent = False
    
    <% '------ Coding part ------ %>
	Self.Returnvalue = Array("")
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== %>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "PA") %>
End Sub
'******************************************* 2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ====================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)	
'=======================================================================================================
Sub SetDefaultVal()
	frm1.txtPlantCd.Value = PopupParent.gPlant
End Sub
'==========================================  2.2.3 InitSpreadSheet()  ===================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("M2111PA1","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 3 
End Sub
'============================================ 2.2.4 SetSpreadLock()  ====================================
' Name : SetSpreadLock
' Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    IF pOpt = "A" Then
	  ggoSpread.Source = frm1.vspdData
	  ggoSpread.SpreadLockWithOddEvenRowColor()
	Else
	
	End IF
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intColCnt

	With frm1.vspdData	
		Redim arrReturn(.MaxCols - 1)
		If .MaxRows > 0 Then 
			.Row = .ActiveRow
			.Col = GetKeyPos("A",C_No)	
			arrReturn(0) = .Text
		end if
	End With
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================================================================================
' Function Name : OpenSortPopup
' Function Desc : OpenSortPopup Reference Popup
'========================================================================================================
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

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_Plant"				
	
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)
	
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "Plant_CD"	
    arrField(1) = "Plant_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.Value= arrRet(1)		
		frm1.txtPlantCd.focus
	End If	
	
End Function

'------------------------------------------  Openitem()  -------------------------------------------------
'	Name : Openitem()
'	Description : item PopUp
'--------------------------------------------------------------------------------------------------------- %>
Function Openitem()  
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd.ClassName)=UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	if Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	end if 
	
	IsOpenPop = True

	arrParam(0) = "품목"						<%' 팝업 명칭 %>
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	<%' TABLE 명칭 %>
	
	arrParam(2) = Trim(frm1.txtItemCd.Value)	<%' Code Condition%>
	
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	    
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(UCase(frm1.txtPlantCd.Value), "''", "S") & " "    <%' Where Condition%>
	End if 
	
	arrParam(5) = "품목"							<%' TextBox 명칭 %>
    arrField(0) = "B_Item.Item_Cd"					<%' Field명(0)%>
    arrField(1) = "B_Item.Item_NM"	
    arrField(2) = "B_Plant.Plant_Cd"					<%' Field명(0)%>
    arrField(3) = "B_Plant.Plant_NM"					<%' Field명(1)%>				<%' Field명(1)%>
    
    arrHeader(0) = "품목"						<%' Header명(0)%>
    arrHeader(1) = "품목명"						<%' Header명(1)%>
    arrHeader(2) = "공장"						<%' Header명(0)%>
    arrHeader(3) = "공장명"						<%' Header명(1)%>
    
	iCalledAspName = AskPRAspName("M1111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M1111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent, arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.value =arrRet(0)
		frm1.txtitemNm.value =arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
	frm1.txtPlantCd.focus
End Sub

'==========================================================================================
'   Event Name : txtFrDt  	 
'   Event Desc :
'==========================================================================================
 Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrDt.Focus
	End If
End Sub

'==========================================================================================
'   Event Name : txtToDt  	 
'   Event Desc :
'==========================================================================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtFrDt2 	 
'   Event Desc :
'==========================================================================================
Sub txtFrDt2_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt2.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrDt2.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtToDt2 	 
'   Event Desc :
'==========================================================================================
Sub txtToDt2_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt2.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToDt2.Focus
	End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
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
'=	Event Name : vspdData_DblClick																		=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
        Exit Sub
    End If

	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
End Sub

'========================================  3.3.2 vspdData_KeyPress()  ===================================
'=	Event Name : vspdData_KeyPress																		=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Sub

'======================================  3.3.3 vspdData_TopLeftChange()  ================================
'=	Event Name : vspdData_TopLeftChange																	=
'=	Event Desc :																						=
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If frm1.vspdData.MaxRows < NewTop +  + VisibleRowCnt(frm1.vspdData,NewTop) And lgPageNo <> "" Then	    <%'☜: 재쿼리 체크 %>
		If CheckRunningBizProcess = True Then
			Exit Sub
		End If
		If DBQuery = False Then		
			Exit Sub
		End If
    End If
End Sub	

'==========================================================================================
'   Event Name : OCX_Keypress()
'   Event Desc : 조회조건부의 OCX_Keypress시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtFrDt2_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToDt2_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

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

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
'       Exit Function
'    End If

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
			strVal = strVal & "&txtPlant=" & .hdnPlant.value
			strVal = strVal & "&txtitem=" & .hdnItem.value
			strVal = strVal & "&txtFrDt=" & .hdnFrDt.Value
			strVal = strVal & "&txtToDt=" & .hdnToDt.Value
			strVal = strVal & "&txtFrDt2=" & .hdnFrDt2.Value
			strVal = strVal & "&txtToDt2=" & .hdnToDt2.Value
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey     
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtPlant=" & .txtPlantCd.value
			strVal = strVal & "&txtitem=" & .txtItemCd.value
			strVal = strVal & "&txtFrDt=" & .txtFrDt.Text
			strVal = strVal & "&txtToDt=" & .txtToDt.Text
			strVal = strVal & "&txtFrDt2=" & .txtFrDt2.Text
			strVal = strVal & "&txtToDt2=" & .txtToDt2.Text
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If				
			
			strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    
End Function	

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	End If
	
	frm1.vspdData.focus
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
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
											   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 tag="14x"></TD>
						<TD CLASS="TD5" NOWRAP>품목</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="품목" NAME="txtItemCd" MAXLENGTH=18 SIZE=10 tag="11NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItem()">
											   <INPUT TYPE=TEXT ALT="품목" NAME="txtItemNm" MAXLENGTH=20 SIZE=20 tag="14X" ></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>요청일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellpadding=0 cellspacing=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m2111pa1_fpDateTime1_txtFrDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
									   <script language =javascript src='./js/m2111pa1_fpDateTime2_txtToDt.js'></script>
									</td>
								</tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>필요일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellpadding=0 cellspacing=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m2111pa1_fpDateTime3_txtFrDt2.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
									   <script language =javascript src='./js/m2111pa1_fpDateTime4_txtToDt2.js'></script>
									</td>
								</tr>
							</table>
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
						<script language =javascript src='./js/m2111pa1_vaSpread1_vspdData.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hdnItem" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlant" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt2" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt2" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
