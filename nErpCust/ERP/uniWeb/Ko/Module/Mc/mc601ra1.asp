<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC601PA1
'*  4. Program Name         : 납입지시입고대상 팝업 
'*  5. Program Desc         : 납입지시입고대상 팝업 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/26
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

Const BIZ_PGM_ID = "MC601RB1.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================
Dim C_PlantCd
Dim C_PlantNm
Dim C_ProdtOrderNo    
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_BaseUnit
Dim C_DoQty
Dim C_RcptQty
Dim C_SlCd
Dim C_SlNm
Dim C_DoDate
Dim C_DoTime
Dim C_TrackingNo
Dim C_InspFlag
Dim C_PoNo
Dim C_PoSeqNo
Dim C_WcCd
Dim C_OprNo
Dim C_Seq
Dim C_SubSeq
Dim C_PurGrp

Dim IsOpenPop          
Dim arrReturn
Dim arrParent
Dim arrBpCd
Dim arrBpNm

arrParent		= window.dialogArguments
set PopupParent = arrParent(0)

Dim arrTemp
arrTemp = arrParent(1)

arrBpCd				= arrTemp(0)
arrBpNm				= arrTemp(1)

top.document.title = PopupParent.gActivePRAspName

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()
	lgStrPrevKeyIndex	= ""
	lgLngCurRows		= 0
	lgSortKey			= 1
	Redim arrReturn(0, 0)
    Self.Returnvalue	= arrReturn  
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","RA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	Dim StartDate
	Dim EndDate

	StartDate	= UNIDateAdd("D", -3, "<%=GetSvrDate%>", PopupParent.gServerDateFormat)
	EndDate		= UNIDateAdd("D", 3, "<%=GetSvrDate%>", PopupParent.gServerDateFormat)	
	frm1.txtDocumentDt1.Text	= UniConvDateAToB(StartDate, PopupParent.gServerDateFormat,PopupParent.gDateFormat)
	frm1.txtDocumentDt2.Text	= UniConvDateAToB(EndDate, PopupParent.gServerDateFormat,PopupParent.gDateFormat)
	
	frm1.txtBpCd.value	= arrBpCd
	frm1.txtBpNm.value	= arrBpNm
End Sub 

'====================================== 2.2.3 InitSpreadSheet() =========================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030228", , PopupParent.gAllowDragDropSpread

	With  frm1.vspdData
		.ReDraw = false
		.MaxCols = C_PurGrp + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.MaxRows = 0
		.OperationMode = 5

		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit	 C_PlantCd,			"공장",			6
	    ggoSpread.SSSetEdit  C_PlantNm,			"공장명",		30
	    ggoSpread.SSSetEdit  C_ProdtOrderNo,	"제조오더번호",	20
	    ggoSpread.SSSetEdit  C_ItemCd,			"품목",			18
	    ggoSpread.SSSetEdit  C_ItemNm,			"품목명",		30
	    ggoSpread.SSSetEdit  C_Spec,			"규격",			30
	    ggoSpread.SSSetEdit  C_BaseUnit,		"단위",			5
	    ggoSpread.SSSetEdit  C_DoQty,			"납입지시수량",	15
	    ggoSpread.SSSetEdit  C_RcptQty,			"입고수량",		15
	    ggoSpread.SSSetEdit  C_SlCd,			"창고",			7
	    ggoSpread.SSSetEdit  C_SlNm,			"창고명",		20
	    ggoSpread.SSSetEdit  C_DoDate,			"납입지시일",	10
	    ggoSpread.SSSetEdit  C_DoTime,			"납입지시시간",	12
	    ggoSpread.SSSetEdit  C_TrackingNo,		"Tracking No.",	20
	    ggoSpread.SSSetEdit  C_InspFlag,		"검사품여부",	8
	    ggoSpread.SSSetEdit  C_PoNo,			"발주번호",		18
	    ggoSpread.SSSetEdit  C_PoSeqNo,			"발주순번",		8
	    ggoSpread.SSSetEdit  C_WcCd,			"작업장",		7
	    ggoSpread.SSSetEdit  C_OprNo,			"공정",			6
	    ggoSpread.SSSetEdit  C_Seq,				"부품예약일련번호",	20
	    ggoSpread.SSSetEdit  C_SubSeq,			"납입지시 순번",	18
		ggoSpread.SSSetEdit	 C_PurGrp,			"구매그룹",			6
	    
	    Call ggoSpread.SSSetColHidden(C_Seq, .MaxCols, True)

		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	    		
		ggoSpread.SSSetSplit2(2)
		
		.ReDraw = true
    End With
End Sub

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_PlantCd		= 1
	C_PlantNm		= 2
	C_ProdtOrderNo	= 3											'☆: Spread Sheet의 Column별 상수 
	C_ItemCd		= 4
	C_ItemNm		= 5
	C_Spec			= 6
	C_BaseUnit		= 7
	C_DoQty			= 8
	C_RcptQty		= 9
	C_SlCd			= 10
	C_SlNm			= 11
	C_DoDate		= 12
	C_DoTime		= 13
	C_TrackingNo	= 14
	C_InspFlag		= 15
	C_PoNo			= 16
	C_PoSeqNo		= 17
	C_WcCd			= 18
	C_OprNo			= 19
	C_Seq			= 20
	C_SubSeq		= 21
	C_PurGrp		= 22
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
		
		C_PlantCd		= iCurColumnPos(1)
	    C_PlantNm		= iCurColumnPos(2)
		C_ProdtOrderNo	= iCurColumnPos(3)
    	C_ItemCd		= iCurColumnPos(4)
	    C_ItemNm		= iCurColumnPos(5)
	    C_Spec			= iCurColumnPos(6)
	    C_BaseUnit		= iCurColumnPos(7)
	    C_DoQty			= iCurColumnPos(8)
	    C_RcptQty		= iCurColumnPos(9)
	    C_SlCd			= iCurColumnPos(10)
	    C_SlNm			= iCurColumnPos(11)
	    C_DoDate		= iCurColumnPos(12)
	    C_DoTime		= iCurColumnPos(13)
	    C_TrackingNo	= iCurColumnPos(14)
	    C_InspFlag		= iCurColumnPos(15)
	    C_PoNo			= iCurColumnPos(16)
	    C_PoSeqNo		= iCurColumnPos(17)
	    C_WcCd			= iCurColumnPos(18)
	    C_OprNo			= iCurColumnPos(19)
	    C_Seq			= iCurColumnPos(20)
	    C_SubSeq		= iCurColumnPos(21)
	    C_PurGrp		= iCurColumnPos(22)
	End Select
End Sub

'===========================================  2.3.1 ()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()
	Dim intInsRow, intRowCnt, intColCnt
    
    With frm1.vspdData

        If .SelModeSelCount > 0 Then

			intInsRow = 0

			Redim arrReturn(.SelModeSelCount-1, .MaxCols - 1)

            For intRowCnt = 0 To .MaxRows
                .Row = intRowCnt
                If .SelModeSelected Then
					For intColCnt = 0 To .MaxCols - 2
					    .Col = intColCnt + 1
					    arrReturn(intInsRow, intColCnt) = .Text
					Next
					intInsRow = intInsRow + 1
				End If
            Next
        End If
    End With
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
	
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================================================================================
' Name : InitComboBox()	
' Desc : Initialize combo value
'========================================================================================================
Sub InitComboBox()
	On Error Resume Next
    Err.Clear
    
    '------------------------------------------------------------
	' Setting Item Account Combo
	'------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("M2110", "''", "S") & "  ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboDoTime,lgF0 ,lgF1 ,Chr(11))

End Sub

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
    Call InitComboBox()
    
    If DbQuery = False Then	Exit Sub
End Sub

Sub txtDocumentDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentDt1.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtDocumentDt1.Focus
    End If
End Sub

Sub txtDocumentDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentDt2.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtDocumentDt2.Focus
    End If
End Sub

Sub txtDocumentDt1_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

Sub txtDocumentDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

Sub txtDocumentDt1_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtDocumentDt2_Change()
    lgBlnFlgChgValue = True
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

'========================================================================================
' Function Name : vspdData_Click
' Function Desc : 그리드 헤더 클릭시 정렬 
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 

 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
 		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)

End Sub	
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If DbQuery = False Then Exit Sub
		End if
	End if

End Sub

'==========================================================================================
'   Event Name : vspdData_KeyPress(KeyAscii)
'   Event Desc : 
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
	On error Resume Next
	
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	Elseif KeyAscii = 27 Then
		Call CancelClick()
	End IF
	
End Function

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : 그리드 폭조정 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
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
'    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
 	frm1.vspdData.Maxrows = 0
   Call InitVariables() 														'⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
'    If Not chkField(Document, "1") Then	Exit Function						'⊙: This function check indispensable field

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery() = False Then Exit Function								     '☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function


'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'	설명 : 
'********************************************************************************************************* 
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
    
    strVal = BIZ_PGM_ID	& "?txtBpCd="				& Trim(frm1.txtBpCd.value)
    strVal = strVal		& "&txtDocumentDt1="		& Trim(frm1.txtDocumentDt1.text)
    strVal = strVal		& "&txtDocumentDt2="		& Trim(frm1.txtDocumentDt2.text)
    strVal = strVal		& "&cboDoTime="				& Trim(frm1.cboDoTime.value)
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
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	End If
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
'######################################################################################################### 
-->
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="GET">


<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" NOWRAP>공급처</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtBpCd" SIZE=14 MAXLENGTH=18 tag="14XXXU" ALT="공급처">&nbsp;<INPUT TYPE="Text" Name="txtBpNm" SIZE=30 MAXLENGTH=40 tag="14XXXU" ALT="공급처명"></TD>
				<TD CLASS="TD5" NOWRAP></TD>
				<TD CLASS="TD6" NOWRAP></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>납입지시일</TD>
				<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/mc601ra1_OBJECT3_txtDocumentDt1.js'></script>
										&nbsp;~&nbsp;
									   <script language =javascript src='./js/mc601ra1_OBJECT4_txtDocumentDt2.js'></script></TD>
				<TD CLASS="TD5" NOWRAP>납입지시시간</TD>
				<TD CLASS="TD6" NOWRAP>
					<SELECT Name="cboDoTime" ALT="납입지시시간" STYLE="WIDTH: 98px" tag="11"><OPTION Value=""></OPTION></SELECT>
				</TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/mc601ra1_OBJECT1_vspdData.js'></script>
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
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


