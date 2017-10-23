
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: Production																*
'*  2. Function Name		: 																			*
'*  3. Program ID			: Reference Popup 확정결과조회												*
'*  4. Program Name			: p4110ra1.asp																*
'*  5. Program Desc			: Reference Popup															*
'*  7. Modified date(First)	: 2000/12/13																*
'*  8. Modified date(Last)	: 2002/12/12																*
'*  9. Modifier (First)		:	Park , Bumsoo															*
'* 10. Modifier (Last)		: Ryu Sung Won																*
'* 11. Comment 				:																			*
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03) 
'*                          : Order Number관련 자리수 조정(2003.04.14) Park Kye Jin                          *
'******************************************************************************************************%>

<HTML>
<HEAD>
<!--'####################################################################################################
'#						1. 선 언 부																		#
'#####################################################################################################-->
<!--'********************************************  1.1 Inc 선언  ****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'============================================  1.1.1 Style Sheet  ===================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 공통 Include  ==================================
'=====================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

'********************************************  1.2 Global 변수/상수 선언  *******************************
'*	Description : 1. Constant는 반드시 대문자 표기														*
'********************************************************************************************************
'============================================  1.2.1 Global 상수 선언  ==================================
'========================================================================================================
Const BIZ_PGM_ID = "p4110rb1.asp"					<% '☆: 비지니스 로직 ASP명 %>

' Grid 1(vspdData1) - Operation
Dim C_OrderNo1
Dim C_ItemCd1
Dim C_ItemNm1
Dim C_StartDt1
Dim C_DueDt1
Dim C_PlanQty1
Dim C_TrackingNo1

' Grid 2(vspdData2) - Operation
Dim C_OrderNo2
Dim C_ItemCd2
Dim C_ItemNm2
Dim C_StartDt2
Dim C_DueDt2
Dim C_PlanQty2
Dim C_TrackingNo2

'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim lgPlantCD
Dim lgPlantNm
Dim lgItemCd
Dim lgProdOrdNo
Dim lgPlanOrdNo
Dim lgOrderQty
Dim lgStartDt
Dim lgEnddt
Dim lgInvStock
Dim lgSFStock
Dim lgForward
Dim lgItemNm
Dim lgStrPrevKey2,lgStrPrevKey3,lgStrPrevKey4
'*********************************************  1.3 변 수 선 언  ****************************************
'*	설명: Constant는 반드시 대문자 표기.																*
'********************************************************************************************************
Dim arrParent
Dim arrParam					
		
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
lgPlantCD = arrParent(1)		
lgPlantNm = arrParent(2)
lgItemCd = arrParent(3)
lgProdOrdNo = arrParent(4)
lgPlanOrdNo = arrParent(5)
lgOrderQty = arrParent(6)
lgStartDt = arrParent(7)
lgEndDt = arrParent(8)
lgInvStock = arrParent(9)
lgSFStock = arrparent(10)
lgForward = arrparent(11)
lgItemNm = arrparent(12)

top.document.title = PopupParent.gActivePRAspName
'########################################################################################################
'#						2. Function 부																	#
'#																										#
'#	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기술					#
'#	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.							#
'#						 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함)			#
'########################################################################################################
'*******************************************  2.1 변수 초기화 함수  *************************************
'*	기능: 변수초기화																					*
'*	Description : Global변수 처리, 변수초기화 등의 작업을 한다.											*
'********************************************************************************************************
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(pvSpdNo)
		Case "A"
			C_OrderNo1		= 1
			C_ItemCd1		= 2
			C_ItemNm1		= 3
			C_StartDt1		= 4
			C_DueDt1		= 5
			C_PlanQty1		= 6
			C_TrackingNo1	= 7
		Case "B"
			C_OrderNo2		= 1
			C_ItemCd2		= 2
			C_ItemNm2		= 3
			C_StartDt2		= 4
			C_DueDt2		= 5
			C_PlanQty2		= 6
			C_TrackingNo2	= 7
	End Select			
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================
Function InitVariables()
	lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
	lgStrPrevKey = ""                           'initializes Previous Key		
	Self.Returnvalue = Array("")
End Function
	
'*******************************************  2.2 화면 초기화 함수  *************************************
'*	기능: 화면초기화																					*
'*	Description : 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다.						*
'********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	'lblTitle.innerHTML = "확정결과조회"
	frm1.txtPlantCd.value  = lgPlantCd
	frm1.txtPlantNm.value = lgPlantNm
	frm1.txtItemCd.value  = lgItemCd
	frm1.txtProdOrderNo.value = lgProdOrdNo
	frm1.txtPlanOrderNo.value = lgPlanOrdNO
	frm1.txtOrderQty.value = lgOrderQty
	frm1.txtStartDt.text = lgStartDt
	frm1.txtEndDt.text = lgEndDt
	frm1.chkInvStock.checked = lgInvStock
	frm1.chkSFStock.checked = lgSFStock
	frm1.chkForward.checked = lgForward
	frm1.txtItemNm.value = lgItemNm
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "RA")%>
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
'=	Name : InitSpreadSheet()																			=
'=	Description : This method initializes spread sheet column property									=
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Select Case UCase(pvSpdNo)
		Case "A"
			'------------------------------------------
			' Grid 1 - Operation Spread Setting
			'------------------------------------------
			Call InitSpreadPosVariables(pvSpdNo)
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread

			With frm1.vspdData1 
			.ReDraw = false
			.MaxCols = C_TrackingNo1 + 1   
			.MaxRows = 0
	
			Call GetSpreadColumnPos(pvSpdNo)
	
			ggoSpread.SSSetEdit 	C_OrderNo1,      "작업지시"		,18 
			ggoSpread.SSSetEdit 	C_ItemCd1,       "품목"			,18 
			ggoSpread.SSSetEdit 	C_ItemNm1,       "품목명"		,25           
			ggoSpread.SSSetDate 	C_StartDt1,		 "착수예정일"	,10, 2, PopupParent.gDateFormat
			ggoSpread.SSSetDate 	C_DueDt1,		 "완료예정일"	,10, 2, PopupParent.gDateFormat        
			ggoSpread.SSSetFloat	C_PlanQty1,		 "오더수량"		,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_TrackingNo1,	 "Tracking No."	,25
  
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)							'frozen 기능추가 
			.ReDraw = true
			End With
		
		Case "B"
			'------------------------------------------
			' Grid 2 - Component Spread Setting
			'------------------------------------------
			Call InitSpreadPosVariables(pvSpdNo)
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021125",, PopupParent.gAllowDragDropSpread
	
			With frm1.vspdData2
			.ReDraw = false
			.MaxCols = C_TrackingNo2 + 1									'☜: 최대 Columns의 항상 1개 증가시킴    
			.MaxRows = 0

			Call GetSpreadColumnPos(pvSpdNo)
	
			ggoSpread.SSSetEdit 	C_OrderNo2,      "구매요청"		, 18 
			ggoSpread.SSSetEdit 	C_ItemCd2,       "품목"			, 18 
			ggoSpread.SSSetEdit 	C_ItemNm2,       "품목명"		, 25           
			ggoSpread.SSSetDate 	C_StartDt2,		 "발주예정일"	, 10, 2, PopupParent.gDateFormat
			ggoSpread.SSSetDate 	C_DueDt2,		 "납기예정일"	, 10, 2, PopupParent.gDateFormat      
			ggoSpread.SSSetFloat	C_PlanQty2,		 "오더수량"		, 15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"    
			ggoSpread.SSSetEdit 	C_TrackingNo2,	 "Tracking No."	, 25
    
			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)							'frozen 기능추가 
			.ReDraw = true
			End With
    End Select
    
	Call SetSpreadLock 
End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    '--------------------------------
    'Grid 1
    '--------------------------------
    ggoSpread.Source = frm1.vspdData1
	ggoSpread.SpreadLockWithOddEvenRowColor()
    
    '--------------------------------
    'Grid 2
    '--------------------------------
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SpreadLockWithOddEvenRowColor()
	
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_OrderNo1		= iCurColumnPos(1)
			C_ItemCd1		= iCurColumnPos(2)
			C_ItemNm1		= iCurColumnPos(3)
			C_StartDt1		= iCurColumnPos(4)
			C_DueDt1		= iCurColumnPos(5)
			C_PlanQty1		= iCurColumnPos(6)
			C_TrackingNo1	= iCurColumnPos(7)
		Case "B"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_OrderNo2		= iCurColumnPos(1)
			C_ItemCd2		= iCurColumnPos(2)
			C_ItemNm2		= iCurColumnPos(3)
			C_StartDt2		= iCurColumnPos(4)
			C_DueDt2		= iCurColumnPos(5)
			C_PlanQty2		= iCurColumnPos(6)
			C_TrackingNo2	= iCurColumnPos(7)
    End Select
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet(gActiveSpdSheet.Id)
	
    If gActiveSpdSheet.Id = "A" Then
		ggoSpread.Source = frm1.vspdData1
	Else
		ggoSpread.Source = frm1.vspdData2
	End If
	
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

Sub vspdData1_KeyPress(KeyAscii)
	If KeyAscii=27 Then
		Call CancelClick()
	End If
End Sub	

Sub vspdData2_KeyPress(KeyAscii)
	If KeyAscii=27 Then
		Call CancelClick()
	End If
End Sub	

'########################################################################################################
'#						3. Event 부																		#
'#	기능: Event 함수에 관한 처리																		#
'#	설명: Window처리, Single처리, Grid처리 작업.														#
'#		  여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.								#
'#		  각 Object단위로 Grouping한다.																	#
'########################################################################################################
'********************************************  3.1 Window처리  ******************************************
'*	Window에 발생 하는 모든 Even 처리																	*
'********************************************************************************************************
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
Sub Form_Load()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
		
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)    
	Call InitVariables
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
		
	Call SetDefaultVal()
	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")
	If DbQuery = False Then	
			Exit Sub
	End If
    frm1.vspdData1.focus
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData1
	Call SetPopupMenuItemInf("0000111111")
	
	If frm1.vspdData1.MaxRows <= 0 Then Exit Sub
	
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	Set gActiveSpdSheet = frm1.vspdData2
	gMouseClickStatus = "SP2C"
	Call SetPopupMenuItemInf("0000111111")
	
    If frm1.vspdData2.MaxRows <= 0 Then Exit Sub
   	  
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Or lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Exit Sub
			End If
		End If
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey3 <> "" Or lgStrPrevKey4 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbDtlQuery = False Then	
				Exit Sub
			End If
		End If
    End if
    
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'*********************************************  3.2 Tag 처리  *******************************************
'*	Document의 TAG에서 발생 하는 Event 처리																*
'*	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나							*
'*	Event간 충돌을 고려하여 작성한다.																	*
'********************************************************************************************************
'==========================================  3.2.1 Search_OnClick =======================================
'========================================================================================================
Function FncQuery()
	On Error Resume Next
End Function

'*********************************************  3.3 Object Tag 처리  ************************************
'*	Object에서 발생 하는 Event 처리																		*
'********************************************************************************************************
'########################################################################################################
'#					     4. Common Function부															#
'########################################################################################################
'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
'********************************************  5.1 DbQuery()  *******************************************
' Function Name : DbQuery																				*
' Function Desc : This function is data query and display												*
'********************************************************************************************************
Function DbQuery()
	Dim strVal    
        
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear                                                               '☜: Protect system from crashing
   
    With frm1
   		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001						'☜: 
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtPlanOrderNo=" & Trim(.txtPlanOrderNo.value)				'☆: 조회 조건 데이타		
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
    
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
    
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    If lgIntFlgMode = PopupParent.OPMD_CMODE Then
		Call SetActiveCell(frm1.vspdData1,1,1,"P","X","X")
		Set gActiveElement = document.activeElement
    End If
    
    lgIntFlgMode = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	frm1.vspdData1.focus
End Function

Function FncExit()
	FncExit = True
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELLSPACING=0>
	<TR>
		<TD HEIGHT=5 WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
 					<TR>
 						<TD CLASS=TD5 NOWRAP>공장</TD>
						<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14" ALT="공장">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>			 						
						<TD CLASS=TD5 NOWRAP>제조오더번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProdOrderNo" SIZE=18 MAXLENGTH=18 tag="14" ALT="제조오더번호"></TD>
						<TD CLASS=TD5 NOWRAP>가용재고반영</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME=chkInvStock ALT="가용재고반영" tag="14" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid"></INPUT>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="14" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
						<TD CLASS=TD5 NOWRAP>계획오더번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlanOrderNo" SIZE=18 tag="14" ALT="계획오더번호"></TD>
						<TD CLASS=TD5 NOWRAP>안전재고반영</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME=chkSFStock ALT="안전재고반영" tag="14" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid"></INPUT></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>작업일자</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/p4110ra1_I375940459_txtStartDt.js'></script>
							&nbsp;~&nbsp;
							<script language =javascript src='./js/p4110ra1_I129479730_txtEndDt.js'></script>									
						</TD>
						<TD CLASS=TD5 NOWRAP>오더수량</TD>
						<TD CLASS=TD6 NOWRAP>
							<script language =javascript src='./js/p4110ra1_I931648784_txtOrderQty.js'></script>
						</TD>
						<TD CLASS=TD5 NOWRAP>Forward</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=CHECKBOX NAME=chkForward ALT="Forward" tag="14" STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid"></INPUT></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE WIDTH="100%" HEIGHT="100%">
				<TR HEIGHT="100%">
					<TD WIDTH="50%">
						<script language =javascript src='./js/p4110ra1_A_vspdData1.js'></script>
					</TD>							
					<TD WIDTH="50%">
						<script language =javascript src='./js/p4110ra1_B_vspdData2.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR><TD HEIGHT=30>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
                                                                                                                                                                                                                                                                                                              