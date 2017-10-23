<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 자료 동기화
'*  2. Function Name        : BOM정보 송신형황(S)
'*  3. Program ID           : XI216MA1_KO119
'*  4. Program Name         : BOM정보 전송형황(S)
'*  5. Program Desc         : BOM정보 전송형황(S)
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2006/04/28
'*  8. Modified date(Last)  : 2006/04/28
'*  9. Modifier (First)     : 권순태
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRDSQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit				'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 
Dim lgTypeCD                                                <%'☜: 'G' is for group , 'S' is for Sort    %>
Dim lgFieldCD                                               <%'☜: 필드 코드값                           %>
Dim lgFieldNM                                               <%'☜: 필드 설명값                           %>
Dim lgFieldLen                                              <%'☜: 필드 폭(Spreadsheet관련)              %>
Dim lgFieldType                                             <%'☜: 필드 설명값                           %>
Dim lgDefaultT                                              <%'☜: 필드 기본값                           %>
Dim lgNextSeq                                               <%'☜: 필드 Pair값                           %>
Dim lgKeyTag                                                <%'☜: Key 정보                              %>
Dim lgNextSeq_T                                             <%'☜: 필드 Pair값                           %>
Dim lgKeyTag_T                                              <%'☜: Key 정보                              %>
Dim lgSortTitleNm                                           <%'☜: Orderby popup용 데이타(필드설명)      %>
Dim lgSortFieldCD1                                          <%'☜: Orderby popup용 데이타(필드코드)      %>
Dim lgMark                                                  <%'☜: 마크                                  %>
Dim lgKeyPos                                                <%'☜: Key위치                               %>
Dim lgKeyPosVal                                             <%'☜: Key위치 Value                         %>
Dim IsOpenPop
Dim arrParam

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID = "XI216MB1_KO119.asp"	
Const BIZ_PGM_JUMP_ID = "P1401MA4"				      '☆: Jump ASP명 
Const C_MaxKey = 2																'☆☆☆☆: Max key value

'========================================================================================================
Dim P_ITEM_CD
Dim P_ITEM_NM
Dim C_ITEM_SQ
Dim C_ITEM_CD
Dim C_ITEM_NM
Dim C_ITEM_NO
Dim C_ITEM_MA
Dim LOT_MANAG
Dim AVL_SDATE
Dim AVL_EDATE
Dim COMM_FLAG
Dim COMM_SEND
Dim COMM_UPDT
Dim COMM_ERRD
Dim COMM_RECV

<% '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- %>
Dim lsConcd
Dim lsConNm
Dim StartDate
Dim LastDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
'LastDate =  UNIDateAdd("m", 1, StartDate, Parent.gDateFormat)
LastDate = StartDate

<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>

'========================================================================================================= 
Sub InitVariables()
   lgBlnFlgChgValue = False                               'Indicates that no value changed
   lgStrPrevKey     = ""                                  'initializes Previous Key
   lgPageNo         = ""
   lgSortKey        = 1
   lgIntFlgMode = parent.OPMD_CMODE	
End Sub

'========================================================================================================
Sub SetDefaultVal()
	frm1.txtFrDt.text = StartDate
	frm1.txtToDt.text = LastDate
	frm1.txtPlantCD.focus
End Sub


'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		WriteCookie "txtplantcd", lsConPlantcd
		WriteCookie "txtitemcd", lsConItemCd
	ElseIf flgs = 0 Then
		Call WriteCookie("txtplantcd" , "")
		Call WriteCookie("txtitemcd" , "")

		strTemp =  ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		If Err.number <> 0 Then
			Err.Clear
			 WriteCookie CookieSplit , ""
			Exit Function 
		End If

		 WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If
End Function

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>
End Sub


'========================================================================================================
Sub InitSpreadSheet()

    Call initSpreadPosVariables()  

	With frm1.vspdData  
		ggoSpread.Source = frm1.vspdData	
		ggoSpread.Spreadinit "V20051206",,parent.gAllowDragDropSpread

		.ReDraw = false

		.MaxCols = COMM_RECV + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols															'☆: 사용자 별 Hidden Column
		.ColHidden = True

		.MaxRows = 0
		ggoSpread.ClearSpreadData

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit  P_ITEM_CD, "모품목코드",		15, , ,18
		ggoSpread.SSSetEdit  P_ITEM_NM, "모품목명",			20, , ,50
		ggoSpread.SSSetEdit  C_ITEM_SQ, "순서",				 4, 1, , 3
		ggoSpread.SSSetEdit  C_ITEM_CD, "자품목코드",		15, , ,18
		ggoSpread.SSSetEdit  C_ITEM_NM, "자품목명",			20, , ,50
		ggoSpread.SSSetFloat C_ITEM_NO, "자품목수량",		10,	parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
		ggoSpread.SSSetEdit  C_ITEM_MA, "자품목단위",		10, 2, , 8
		ggoSpread.SSSetEdit  LOT_MANAG, "LOT관리여부",		10, 2, ,1
		ggoSpread.SSSetDate  AVL_SDATE, "유효시작일",		10, 2, parent.gDateFormat
		ggoSpread.SSSetDate  AVL_EDATE, "유효종료일",		10, 2, parent.gDateFormat
		ggoSpread.SSSetEdit  COMM_FLAG, "생성구분",			 8, 2, , 1
		ggoSpread.SSSetEdit  COMM_SEND, "최종송신일시",		20, , ,30
		ggoSpread.SSSetEdit  COMM_UPDT, "MES수신여부",		10, 2, , 1
		ggoSpread.SSSetEdit  COMM_ERRD, "에러내역",			25, , ,50
		ggoSpread.SSSetEdit  COMM_RECV, "MES최종수신일시",	20, , ,30

		.ReDraw = true

		Call SetSpreadLock 

    End With
End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
    P_ITEM_CD = 1
    P_ITEM_NM = 2
    C_ITEM_SQ = 3
    C_ITEM_CD = 4
    C_ITEM_NM = 5
    C_ITEM_NO = 6
    C_ITEM_MA = 7
    LOT_MANAG = 8
    AVL_SDATE = 9
    AVL_EDATE = 10
    COMM_FLAG = 11
    COMM_SEND = 12
    COMM_UPDT = 13
    COMM_ERRD = 14
    COMM_RECV = 15

End Sub

'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			P_ITEM_CD = iCurColumnPos(1)
			P_ITEM_NM = iCurColumnPos(2)
			C_ITEM_SQ = iCurColumnPos(3)
			C_ITEM_CD = iCurColumnPos(4)
			C_ITEM_NM = iCurColumnPos(5)
			C_ITEM_NO = iCurColumnPos(6)
			C_ITEM_MA = iCurColumnPos(7)
			LOT_MANAG = iCurColumnPos(8)
			AVL_SDATE = iCurColumnPos(9)
			AVL_EDATE = iCurColumnPos(10)
			COMM_FLAG = iCurColumnPos(11)
			COMM_SEND = iCurColumnPos(12)
			COMM_UPDT = iCurColumnPos(13)
			COMM_ERRD = iCurColumnPos(14)
			COMM_RECV = iCurColumnPos(15)

    End Select
End Sub

'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal lRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected  1, lRow, lRow
		ggoSpread.SSSetProtected  2, lRow, lRow
		ggoSpread.SSSetProtected  3, lRow, lRow
		ggoSpread.SSSetProtected  4, lRow, lRow
		ggoSpread.SSSetProtected  5, lRow, lRow
		ggoSpread.SSSetProtected  6, lRow, lRow
		ggoSpread.SSSetProtected  7, lRow, lRow
		ggoSpread.SSSetProtected  8, lRow, lRow
		ggoSpread.SSSetProtected  9, lRow, lRow
		ggoSpread.SSSetProtected 10, lRow, lRow
		ggoSpread.SSSetProtected 11, lRow, lRow
		ggoSpread.SSSetProtected 12, lRow, lRow
		ggoSpread.SSSetProtected 13, lRow, lRow
		ggoSpread.SSSetProtected 14, lRow, lRow
		ggoSpread.SSSetProtected 15, lRow, lRow
		ggoSpread.SSSetProtected 16, lRow, lRow
		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	If iWhere = 0 Then
		arrParam(0) = "공  장"								<%' 팝업 명칭 %>
		arrParam(1) = "B_PLANT"								<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtPlantCD.value)			<%' Code Condition%>
		arrParam(3) = ""									<%' Name Cindition%>
		arrParam(4) = ""									<%' Where Condition%>
		arrParam(5) = "공  장"								<%' TextBox 명칭 %>
		
		arrField(0) = "PLANT_CD"							<%' Field명(0)%>
		arrField(1) = "PLANT_NM"							<%' Field명(1)%>
	    
		arrHeader(0) = "공 장"								<%' Header명(0)%>
		arrHeader(1) = "공장명"								<%' Header명(1)%>

		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
										Array(arrParam, arrField, arrHeader), _
										"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		frm1.txtPlantCD.focus 
	ElseIf iWhere = 1 Then
		If frm1.txtPlantCd.value = "" Then
			Call DisplayMsgBox("971012", "X", "공장", "X")
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
		arrParam(1) = Trim(frm1.txtPItemCd.value)	' Item Code
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
		arrParam(3) = ""							' Default Value
		
		arrField(0) = 1 							' Field명(0) : "ITEM_CD"	
		arrField(1) = 2 							' Field명(1) : "ITEM_NM"	
		arrField(2) = 3								' Field명(1) : "ITEM_ACCT"
		arrField(3) = 8								' Field명(1) : "PHANTOM_FLG"	
		arrField(4) = 5								' Field명(1) : "PROCUR_TYPE"
	    
		iCalledAspName = AskPRAspName("B1B11PA4")
		
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
			IsOpenPop = False
			Exit Function
		End If
		
		arrRet = window.showModalDialog(iCalledAspName, _
										Array(Window.parent, arrParam, arrField), _
										"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		frm1.txtPItemCD.focus 
	Else
		If frm1.txtPlantCd.value = "" Then
			Call DisplayMsgBox("971012", "X", "공장", "X")
			frm1.txtPlantCd.focus
			Set gActiveElement = document.activeElement 
			IsOpenPop = False
			Exit Function
		End If

		arrParam(0) = Trim(frm1.txtPlantCd.value)   ' Plant Code
		arrParam(1) = Trim(frm1.txtCItemCd.value)	' Item Code
		arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
		arrParam(3) = ""							' Default Value
		
		arrField(0) = 1 							' Field명(0) : "ITEM_CD"	
		arrField(1) = 2 							' Field명(1) : "ITEM_NM"	
		arrField(2) = 3								' Field명(1) : "ITEM_ACCT"
		arrField(3) = 8								' Field명(1) : "PHANTOM_FLG"	
		arrField(4) = 5								' Field명(1) : "PROCUR_TYPE"
	    
		iCalledAspName = AskPRAspName("B1B11PA4")
		
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA4", "X")
			IsOpenPop = False
			Exit Function
		End If
		
		arrRet = window.showModalDialog(iCalledAspName, _
										Array(Window.parent, arrParam, arrField), _
										"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		frm1.txtCItemCD.focus 
	End If
    

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBpCode(arrRet, iWhere)
	End If	
End Function

'========================================================================================================= 
Function SetBpCode(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtPlantCD.value = arrRet(0) 
			.txtPlantNM.value = arrRet(1)   
		ElseIf iWhere = 1 Then
			.txtPItemCD.value = arrRet(0) 
			.txtPItemNM.value = arrRet(1)   
		Else
			.txtCItemCD.value = arrRet(0) 
			.txtCItemNM.value = arrRet(1)   
		End If
	End With
End Function

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029										'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)	
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gDateFormat,	parent.gComNum1000,	parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field

    '----------  Coding part  -------------------------------------------------------------
	Call InitVariables														    '⊙: Initializes local global variables

	Call SetDefaultVal		

	Call InitSpreadSheet()

    Call SetToolbar("11000000000011")				'⊙: 버튼 툴바 제어 

	Call CookiePage(0)
	
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtPItemCd.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If

End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
End Sub

'==========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

'========================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    
	If CheckRunningBizProcess = True Then  Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgStrPrevKey <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromReqrdDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToReqrdDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
Function FncQuery() 
    FncQuery = False											'⊙: Processing is NG
    
    If Trim(frm1.txtFrDt.Text) > Trim(frm1.txtToDt.Text) Then
		MsgBox "종료일은 시작일 이후이어야 합니다."
        Exit Function
    End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
    Call InitVariables 											'⊙: Initializes local global variables
'    Call SetDefaultVal

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							'⊙: This function check indispensable field
       Exit Function
    End If
	
    '-----------------------
    'Radio Button Check area
    '-----------------------
	If frm1.rdoUsage_flag1.checked = True Then
		frm1.txtRadioFlag.value  = "" 
	ElseIf frm1.rdoUsage_flag2.checked = True Then
		frm1.txtRadioFlag.value = "Y" 
	ElseIf frm1.rdoUsage_flag3.checked = True Then
		frm1.txtRadioFlag.value = "N" 
	End If

    '-----------------------
    'Query function call area
    '------------------------
    Call DbQuery												'☜: Query db data

    FncQuery = True												'⊙: Processing is OK
End Function

'========================================================================================
Function FncPrint()
    ggoSpread.Source = frm1.vspdData
	Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_MULTI, False)
End Function

'========================================================================================
Function FncSplitColumn()
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
	iColumnLimit  = frm1.vspdData.MaxCols

	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   
    
	Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE

	ggoSpread.Source = frm1.vspdData

	ggoSpread.SSSetSplit(ACol)    

	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0    
	frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
End Function

'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"x","x")   'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	FncExit = True
End Function

'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    DbQuery = False

	If LayerShowHide(1) = False Then
		Exit Function 
	End If

	Dim strVal

    With frm1
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 & _
				 "&txtPlantCD=" & Trim(.txtPlantCD.value) & _
				 "&txtPItemCD=" & Trim(.txtPItemCD.value) & _
				 "&txtCItemCD=" & Trim(.txtCItemCD.value) & _
				 "&txtFrDT=" & Trim(.txtFrDt.text) & _
				 "&txtToDT=" & Trim(.txtToDt.text) & _
				 "&txtRadioFlag=" & Trim(.txtRadioFlag.value) & _
				 "&txtMaxRows="   & .vspdData.MaxRows & _
				 "&lgStrPrevKey=" & lgStrPrevKey

		Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
    End With
   
    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	lgIntFlgMode = parent.OPMD_UMODE										'Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
	'-----------------------
	'Reset variables area
	'-----------------------
	Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	'폴더/조회/입력
	'/삭제/저장/한줄In
	'/한줄Out/취소/이전
	'/다음/복사/엑셀
	'/인쇄/찾기
	Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어 

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					<TD WIDTH=* align=right></TD>
					<TD WIDTH=10>&nbsp;</TD>	
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공  장</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT NAME="txtPlantCD" TYPE="Text" MAXLENGTH="10" TAG="12XXXU" SIZE="10"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="ImgPlantCD" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenPopUp frm1.txtPlantCD.value, 0"> <INPUT NAME="txtPlantNM" TYPE="Text" TAG="14" SIZE="25">
									</TD>
									<TD CLASS="TD5" NOWRAP>송신기간</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> 
										ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtFrDt CLASSID=<%=gCLSIDFPDT%> tag="12" ALT="시작일"></OBJECT>');
										</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> 
										ExternalWrite('<OBJECT title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtToDt CLASSID=<%=gCLSIDFPDT%> tag="12" ALT="종료일"></OBJECT>');
										</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>모품목</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT NAME="txtPItemCD" TYPE="Text" MAXLENGTH="20" TAG="11XXXU" SIZE="20"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="ImgPItemCD" ALIGN="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenPopUp frm1.txtPItemCD.value, 1"> <INPUT NAME="txtPItemNM" TYPE="Text" TAG="14" SIZE="25">
									</TD>
									<TD CLASS="TD5" NOWRAP>수신여부</TD>
									<TD CLASS="TD6" NOWRAP>
									<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag1" value="" tag = "11" checked>
										<label for="rdoUsage_flag1">전체</label>
									<input type=radio CLASS = "RADIO" name="rdoUsage_flag" id="rdoUsage_flag2" value="Y" tag = "11">
										<label for="rdoUsage_flag2">성공</label>
									<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag3" value="N" tag = "11">
										<label for="rdoUsage_flag3">실패</label></TD>
								</TR>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>자품목</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT NAME="txtCItemCD" TYPE="Text" MAXLENGTH="20" TAG="11XXXU" SIZE="20"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="ImgCItemCD" ALIGN="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenPopUp frm1.txtCItemCD.value, 2"> <INPUT NAME="txtCItemNM" TYPE="Text" TAG="14" SIZE="25">
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP style="TEXT-ALIGN:center;"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23XXX" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">BOM등록(Multi)</a>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../Blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="HBp_cdFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="HBp_cdTo" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadioFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadioType" tag="24">
			
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>
</BODY>
</HTML>