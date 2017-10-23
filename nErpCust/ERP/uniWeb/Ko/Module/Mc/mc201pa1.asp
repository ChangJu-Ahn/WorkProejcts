<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC201PA1
'*  4. Program Name         : 납입지시대상 팝업 
'*  5. Program Desc         : 납입지시대상 팝업 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/24
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Lee Seung Wook
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
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->
'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

Const BIZ_PGM_ID = "MC201pb1.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_PoNo    
Dim C_PoSeqNo
Dim C_BpCd
Dim C_BpNm
Dim C_PoQty
Dim C_PoUnit
Dim C_BaseQty
Dim C_BaseUnit
Dim C_PurGroup
Dim C_PurNo


Dim arrReturn
Dim arrParent
Dim arrParam
Dim arrPlantCd
Dim arrPlantNm
Dim arrItemCd
Dim arrItemNm
Dim arrTrackingNo
Dim arrReqQty
Dim arrBpCd
		
arrParent		= window.dialogArguments

set PopupParent = arrParent(0)
Dim arrTemp

arrTemp		= arrParent(1)

arrPlantCd		= arrTemp(0)		
arrPlantNm		= arrTemp(1)
arrItemCd		= arrTemp(2)
arrItemNm		= arrTemp(3)
arrTrackingNo	= arrTemp(4)
arrReqQty		= arrTemp(5)
arrBpCd			= arrTemp(6)

top.document.title = PopupParent.gActivePRAspName

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
Dim lgOldRow
Dim gblnWinEvent
Dim strReturn
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
	lgStrPrevKeyIndex	= ""
	lgLngCurRows		= 0
	lgSortKey			= 1
	Redim arrReturn(0)
    Self.Returnvalue	= arrReturn  
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","PA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtPlantCd.value		= arrPlantCd
	frm1.txtPlantNm.value		= arrPlantNm
	frm1.txtItemCd.value		= arrItemCd
	frm1.txtPlantNm.value		= arrItemNm
	frm1.txtTrackingNo.value	= arrTrackingNo
	
	If frm1.txtPlantCd.value <> "" Then
		If CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			Call DisplayMsgBox("125000","X","X","X")
			frm1.txtPlantNm.Value = ""
			Exit Sub
		End If
	End if
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	
	If frm1.txtItemCd.value <> "" Then
		If  CommonQueryRs(" B.ITEM_NM "," B_ITEM_BY_PLANT A, B_ITEM B ", " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				Call DisplayMsgBox("122600","X","X","X")
				frm1.txtItemNm.Value = ""
				Exit Sub
			Else
				lgF0 = Split(lgF0, Chr(11))
				frm1.txtItemNm.Value = lgF0(0)
				Call DisplayMsgBox("122700","X","X","X")
				Exit Sub
			End If
		End If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtItemNm.Value = lgF0(0)
	Else
		frm1.txtItemNm.Value = ""
	End if 
	
    Self.Returnvalue = Array("")
End Sub 

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030108", , PopupParent.gAllowDragDropSpread

	With  frm1.vspdData
			.ReDraw = false
	        .OperationMode = 3
	    	.MaxCols = C_PurNo+1												'☜: 최대 Columns의 항상 1개 증가시킴 
	    	.MaxRows = 0

		Call GetSpreadColumnPos("A")
		
	    ggoSpread.SSSetEdit  C_PoNo,		"발주번호",			18, 0, -1, 18
	    ggoSpread.SSSetEdit  C_PoSeqNo,		"발주순번",			8, 0, -1, 4
	    ggoSpread.SSSetEdit  C_BpCd,		"공급처",			18, 0, -1,10
	    ggoSpread.SSSetEdit	 C_BpNm,		"공급처명",			25, 0, -1, 50
	    ggoSpread.SSSetFloat C_PoQty,		"발주량",			10,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	    ggoSpread.SSSetEdit  C_PoUnit,		"발주단위",			8, 0, -1, 3
	    ggoSpread.SSSetFloat C_BaseQty,		"재고단위발주수량",	10,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	    ggoSpread.SSSetEdit  C_BaseUnit,	"재고단위",			8, 0, -1, 3
	    ggoSpread.SSSetEdit  C_PurGroup,	"구매그룹",			8, 0, -1, 4
	    ggoSpread.SSSetEdit  C_PurNo,		"구매요청번호",		12, 0, -1, 18
	    
	    Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	    		
		ggoSpread.SSSetSplit(2)
		
		Call SetSpreadLock() 
		.ReDraw = true
    End With
End Sub

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_PoNo				= 1											'☆: Spread Sheet의 Column별 상수 
	C_PoSeqNo			= 2
	C_BpCd				= 3
	C_BpNm				= 4
	C_PoQty				= 5
	C_PoUnit			= 6
	C_BaseQty			= 7
	C_BaseUnit			= 8
	C_PurGroup			= 9
	C_PurNo				= 10
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock()
		ggoSpread.Source = frm1.vspdData
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
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_PoNo			= iCurColumnPos(1)
    	C_PoSeqNo		= iCurColumnPos(2)
	    C_BpCd			= iCurColumnPos(3)
	    C_BpNm			= iCurColumnPos(4)
	    C_PoQty			= iCurColumnPos(5)
	    C_PoUnit		= iCurColumnPos(6)
	    C_BaseQty		= iCurColumnPos(7)
	    C_BaseUnit		= iCurColumnPos(8)
	    C_PurGroup		= iCurColumnPos(9)
	    C_PurNo			= iCurColumnPos(10)
	End Select
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================
Sub SetSpreadColor(ByVal lRow)
    frm1.vspdData.ReDraw = False
		
		ggoSpread.SSSetProtected	C_PoNo, lRow, lRow
		ggoSpread.SSSetProtected	C_PoSeqNo, lRow, lRow
		ggoSpread.SSSetProtected	C_BpCd, lRow, lRow
		ggoSpread.SSSetProtected	C_BpNm, lRow, lRow
		ggoSpread.SSSetProtected	C_PoQty, lRow, lRow
		ggoSpread.SSSetProtected	C_PoUnit, lRow, lRow
		ggoSpread.SSSetProtected	C_BaseQty, lRow, lRow
		ggoSpread.SSSetProtected	C_BaseUnit, lRow, lRow
		ggoSpread.SSSetProtected	C_PurGroup, lRow, lRow
		ggoSpread.SSSetProtected	C_PurNo, lRow, lRow
	
    frm1.vspdData.ReDraw = True
End Sub

'===========================================  2.3.1 ()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
  Dim intColCnt
  
  If frm1.vspdData.ActiveRow > 0 Then 
   Redim arrReturn(frm1.vspdData.MaxCols - 1)
  
   frm1.vspdData.Row = frm1.vspdData.ActiveRow
     
   frm1.vspdData.Col	=	C_PoNo
	arrReturn(0)		=	frm1.vspdData.Text
	frm1.vspdData.Col	=	C_PoSeqNo
	arrReturn(1)		=	frm1.vspdData.Text
	frm1.vspdData.Col	=	C_BpCd
	arrReturn(2)		=	frm1.vspdData.Text
	frm1.vspdData.Col	=	C_BpNm
	arrReturn(3)		=	frm1.vspdData.Text
	frm1.vspdData.Col	=	C_PoQty
	arrReturn(4)		=	frm1.vspdData.Text
	frm1.vspdData.Col	=	C_PoUnit
	arrReturn(5)		=	frm1.vspdData.Text
	frm1.vspdData.Col	=	C_BaseQty
	arrReturn(6)		=	frm1.vspdData.Text
	frm1.vspdData.Col	=	C_BaseUnit
	arrReturn(7)		=	frm1.vspdData.Text 
    frm1.vspdData.Col	=	C_PurGroup
	arrReturn(8)		=	frm1.vspdData.Text 
	frm1.vspdData.Col	=	C_PurNo
	arrReturn(9)		=	frm1.vspdData.Text 
    
   Self.Returnvalue = arrReturn
  End If
  
  Self.Close()
 End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()		
	Self.Close()
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    
	'----------  Coding part  -------------------------------------------------------------
    Call InitSpreadSheet
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal()
    Call InitSpreadSheet()
    
    If DbQuery = False Then
		Exit Sub
	End if
End Sub

'******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If DbQuery = False Then
				Exit Sub
			End if
		End if
	End if
End Sub

'==========================================================================================
'   Event Name : vspdData_KeyPress(KeyAscii)
'   Event Desc : 
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
	On error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	End if
	
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	Elseif KeyAscii = 27 Then
		Call CancelClick()
	End IF
End Function

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
'    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
	frm1.vspdData.Maxrows = 0
    Call InitVariables() 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
'       Exit Function
'    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery() = False Then
		Exit Function
	End if
       
    FncQuery = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Call LayerShowHide(1)
    
    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing
    Dim strVal
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------        
    
    strVal = BIZ_PGM_ID	& "?txtPlantCd="		& Trim(frm1.txtPlantCd.value)
    strVal = strVal		& "&txtItemCd="			& Trim(frm1.txtItemCd.value)
    strVal = strVal		& "&txtTrackingNo="		& Trim(frm1.txtTrackingNo.value)
    strVal = strVal		& "&hReqQty="			& arrReqQty
    strVal = strVal		& "&hBpCd="				& arrBpCd
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------
    strVal = strVal     & "&txtMaxRows="		& frm1.vspdData.MaxRows
	strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
    
    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	frm1.vspdData.Focus
End Function


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag부 
'	기능: Tag부분 설정 
	' 입력 필드의 경우 MaxLength=? 를 기술 
	' CLASS="required" required  : 해당 Element의 Style 과 Default Attribute 
		' Normal Field일때는 기술하지 않음 
		' Required Field일때는 required를 추가하십시오.
		' Protected Field일때는 protected를 추가하십시오.
			' Protected Field일경우 ReadOnly 와 TabIndex=-1 를 표기함 
	' Select Type인 경우에는 className이 ralargeCB인 경우는 width="153", rqmiddleCB인 경우는 width="90"
	' Text-Transform : uppercase  : 표기가 대문자로 된 텍스트 
	' 숫자 필드인 경우 3개의 Attribute ( DDecPoint DPointer DDataFormat ) 를 기술 
'######################################################################################################### %>
-->

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="GET">


<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" NOWRAP>공장</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtPlantCd" SIZE=12 MAXLENGTH=18 tag="14XXXU" ALT="공장">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 MAXLENGTH=40 tag="14XXXU" ALT="공장명"></TD>
				<TD CLASS="TD5" NOWRAP>품목</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtItemCd" SIZE=12 MAXLENGTH=18 tag="14XXXU" ALT="품목">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 MAXLENGTH=40 tag="14XXXU" ALT="품목명"></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="14XXXU" ALT="Tracking No."></TD>
				<TD CLASS="TD5" NOWRAP></TD>
				<TD CLASS="TD6" NOWRAP></TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/mc201pa1_OBJECT1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
	
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
	
				<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="okclick()"    ></IMG>
						                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
	
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>		
		</TD>
	</TR>
</TABLE>
	<INPUT TYPE=HIDDEN NAME="hReqQty" tag="14" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


