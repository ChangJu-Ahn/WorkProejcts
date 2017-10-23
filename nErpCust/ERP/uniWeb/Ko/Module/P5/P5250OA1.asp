<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 생산관리 
*  2. Function Name        : 설비관리대장(HB)
*  3. Program ID           : P5250OA1
*  4. Program Name         : 설비관리대장(HB)
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2005/07/29
*  9. Modifier (First)     : Joo Young Hoon
* 10. Modifier (Last)      : Joo Young Hoon
* 11. Comment              :
=======================================================================================================-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>



<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID       = "P5250OB1.asp"						           '☆: Biz Logic ASP Name



Dim C_FACILITY_CD
Dim C_FACILITY_NM
Dim C_EMP_CD
Dim C_SET_DT
Dim C_PROD_CD
Dim C_PROD_AMT
Dim C_PM_DT	
Dim C_BIGO					

Const C_SHEETMAXROWS = 30

Dim iDBSYSDate
Dim EndDate, StartDate,StartDate_,EndDate_,selChk,ACT_ROW

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = DateAdd("d", -7, EndDate)

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  	
	
	C_FACILITY_CD			= 1
	C_FACILITY_NM			= 2
	C_EMP_CD				= 3
	C_SET_DT				= 4
	C_PROD_CD				= 5
	C_PROD_AMT				= 6
	C_PM_DT					= 7
	C_BIGO					= 8

End Sub
	

Sub SetDefaultVal()		

	'frm1.txtReqdlvyFromDt.text = StartDate
	frm1.txtReqdlvyToDt.text = Enddate		
	
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
 Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	
	ggoSpread.Source = frm1.vspdData
	
	ggoSpread.Spreadinit	"V20021105",, parent.gAllowDragDropSpread    
	
	Call AppendNumberPlace("6", "5", "0")
	
	With frm1.vspdData
		
		.ReDraw = False
		  
		.MaxCols = C_BIGO + 1
		.MaxRows = 0
		
		
		Call ggoSpread.ClearSpreadData()	
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit    C_FACILITY_CD			,		"설비코드"			,	15
		ggoSpread.SSSetEdit    C_FACILITY_NM			,		"설비명칭"			,	15
		ggoSpread.SSSetEdit    C_EMP_CD					,		"담당자"			,	15
		ggoSpread.SSSetEdit    C_SET_DT					,		"설치일자"			,	15
		ggoSpread.SSSetEdit    C_PROD_CD				,		"제작업체"			,	15
		ggoSpread.SSSetFloat    C_PROD_AMT				,		"금액"				,	15, "6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"    
		ggoSpread.SSSetDate    C_PM_DT					,		"폐기일자"			,	11, 2, parent.gDateFormat		
		ggoSpread.SSSetEdit    C_BIGO					,		"비고"				,	10
	
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
				
		
		.ReDraw = true
		
	End With
	
	
	
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
		
			C_FACILITY_CD			= iCurColumnPos(1)
			C_FACILITY_NM			= iCurColumnPos(2)
			C_EMP_CD				= iCurColumnPos(3)
			C_SET_DT				= iCurColumnPos(4)
			C_PROD_CD				= iCurColumnPos(5)			
			C_PROD_AMT				= iCurColumnPos(6)
			C_PM_DT					= iCurColumnPos(7)		
			C_BIGO					= iCurColumnPos(8)
		
    End Select    
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
   ' frm1.vspdData.MaxRows = 0
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %> 
End Sub


'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================



'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()


	Err.Clear    
	                                                                    '☜: Clear err status
	Call LoadInfTB19029   
                                                         '☜: Load table , B_numeric_format
 		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec) 'condition
 
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

	
	
	Call InitSpreadSheet  
	
	 
	
	Call SetToolbar("1100000000000111")							 					'⊙: Set ToolBar

	Call InitVariables

	Call InitComboBox
	
	
	call SetDefaultVal
	
	frm1.txtFacilityCd.focus


End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
	
  Dim IntRetCD 
    Dim RetStatus

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    
    
    
    Call ggoSpread.ClearSpreadData()	
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    Call InitVariables                                                       '⊙: Initializes local global variables

    
		
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery() = False Then
        Call RestoreToolBar()
        Exit Function
    End If

    selChk=false
    
    FncQuery = True                                                              '☜: Processing is OK
																'⊙: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    FncNew = True																 '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    dim file_end
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status

End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    FncCopy = True                                                            '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 

    FncCancel = False                                                            '☜: Processing is NG
  
    FncCancel = true                                                         '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : developer describe this line Called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
	On Error Resume Next                                                         '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    FncNext = True                                                               '☜: Processing is OK
	
End Function
'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function


'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery()

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
    
    With frm1

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001		
		strVal = strVal & "&txtReqdlvyFromDt=" & Trim(.txtReqdlvyFromDt.text)
		strVal = strVal & "&txtReqdlvyToDt=" & Trim(.txtReqdlvyToDt.text)			
		strVal = strVal & "&txtFacilityCd=" & Trim(.txtFacilityCd.value)	
		strVal = strVal & "&txtFacilityNm=" & Trim(.txtFacilityNm.value)
		strVal = strVal & "&seltype=" & Trim(.seltype.value)

		StartDate_=frm1.txtReqdlvyFromDt.text
		EndDate_=frm1.txtReqdlvyToDt.text  

 		Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	        
    End With
	    
    DbQuery = True 

End Function


Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False

		ggoSpread.spreadlock C_AS_DT			, -1
		ggoSpread.spreadlock C_AS_NO			, -1
		
		ggoSpread.spreadlock C_AS_BP_CD			, -1
		ggoSpread.spreadlock C_AS_ITEM_CD		, -1
		
		ggoSpread.spreadlock C_AS_NUMBER		, -1
		ggoSpread.spreadlock C_AS_RECEIVER		, -1
		ggoSpread.spreadlock C_AS_TYPE			, -1
		
		ggoSpread.spreadlock C_AS_TEXT			, -1
	
		ggoSpread.spreadlock C_AS_PROCESS		, -1
		ggoSpread.spreadlock C_AS_TREATER		, -1
		ggoSpread.spreadlock C_AS_TREAT_DT		, -1
		ggoSpread.spreadlock C_AS_RESULT		, -1
		ggoSpread.spreadlock C_AS_COWORKER		, -1
		ggoSpread.spreadlock C_BIGO				, -1

    .vspdData.ReDraw = True

    End With

End Sub


'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()

	                                                           '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		

	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
   
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

End Function

	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function



<% '------------------------------------------  OpenRequried()  -------------------------------------------------
'	Name : OpenRequried()
'	Description : Sales Org Display PopUp
'--------------------------------------------------------------------------------------------------------- %>

Function OpenRequried(ByVal iRequried)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iRequried

	Case 1												
	
		arrParam(0) = "금형코드조회"					<%' 팝업 명칭 %>
		arrParam(1) ="Y_CAST"	<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtFacilityCd.value)		<%' Code Condition%>
		'arrParam(3) = Trim(frm1.txtDn_TypeNm.value)		<%' Name Cindition%>
		arrParam(4) = " " 
		arrParam(5) = "금형코드"			  	   <%' TextBox 명칭 %>

		arrField(0) = "CAST_CD"							<%' Field명(0)%>
		arrField(1) = "CAST_NM"							<%' Field명(1)%>

		arrHeader(0) = "금형코드"					<%' Header명(0)%>
		arrHeader(1) = "금형명칭"					<%' Header명(1)%>

			 
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")


	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetRequried(arrRet,iRequried)
	End If	
	
End Function

<% '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= %>
<% '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ %>
<% '------------------------------------------  SetRequried()  --------------------------------------------------
'	Name : SetRequried()
'	Description : 거래처 Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- %>
Function SetRequried(Byval arrRet,ByVal iRequried)

	Select Case iRequried
	Case 1
		
		frm1.txtFacilityCd.value = Trim(arrRet(0))
		frm1.txtFacilityNm.value = Trim(arrRet(1))	
			
	End Select
	
	lgBlnFlgChgValue=true
	

End Function



Function PgmJump1(PGM_JUMP_ID)
    Call BtnDisabled(1)
    
    Call CookiePage(1)  ' Write Cookie
   
    PgmJump(PGM_JUMP_ID)
	Call BtnDisabled(0)
End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
     '----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	
	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
 			lgSortKey = 1
 		End If
 	Else	
		selChk= True	
 	End If
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	If NewRow > 0 And Row <> NewRow Then

	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	
    If Button <> "1" And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
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
    Call InitSpreadSheet()  
    Call ggoSpread.ReOrderingSpreadData
    
End Sub 


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery() = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


'==========================================================================================
'   Event Name : BtnPrint()
'   Event Desc : 
'==========================================================================================
Function BtnPrint()

    Dim StrEbrFile, condvar

  	If frm1.vspdData.ActiveRow < 1 Then
		msgbox "먼저 출력할 설비코드를 클릭하십시요"
		exit function
	End If

	If lgIntFlgMode = parent.OPMD_CMODE Then						'/조회여부 확인 
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If
	  Call PrintCond(strEbrFile, condvar)
	  Call FncEBRprint(EBAction, StrEbrFile, condvar) 	

End Function


'==========================================================================================
'   Event Name : BtnPreView()
'   Event Desc : 
'==========================================================================================
Function BtnPreView()
    
    Dim strEbrFile
    Dim objName
    
	Dim var1
	

	
	dim strUrl
	dim arrParam, arrField, arrHeader

	Call BtnDisabled(1)

	If frm1.vspdData.ActiveRow < 1 Then
		msgbox "먼저 출력할 설비코드를 클릭하십시요"
		exit function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_FACILITY_CD
		var1 = Trim(.Text)
	End With
	
	strUrl = "cast_cd|" & var1 
	
	
	ObjName = AskEBDocumentName("P5250OA1","ebr")

	call FncEBRPreview(objName, strUrl)

	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement 
    
End Function


'==========================================================================================
'   Event Name : PrintCond(strEbrFile, condvar)
'   Event Desc : 
'==========================================================================================
Sub PrintCond(strEbrFile, condvar)

    Dim var1 
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col = C_FACILITY_CD
		var1 = .Text
	End With

    condvar =   " cast_cd|"      & var1
	
    StrEbrFile = "P5250OA1.ebr"    
    

End Sub

Sub txtReqdlvyFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqdlvyFromDt.Action = 7
	End If
End Sub

Sub txtReqdlvyFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call fncquery()
End Sub

Sub txtReqdlvyToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtReqdlvyToDt.Action = 7
	End If
End Sub

Sub txtReqdlvyToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call fncquery()
End Sub


Sub InitComboBox()
	
	
	Call CommonQueryRs(" minor_cd,minor_nm "," B_MINOR "," major_Cd = " & FilterVar("y6002", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.selType, lgF0, lgF1, Chr(11))
    
    frm1.selType.value = ""
    
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설비관리대장</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
								
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100%  height=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>설비코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFacilityCd" ALT="설비코드" TYPE="Text" MAXLENGTH="13" SIZE=20 tag="11XXXU"></td>
									<TD CLASS=TD5 NOWRAP>설비명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFacilityNm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="11"></TD>	
									
								</TR>
								<Tr>
									<Td class=td5>담당자</td>
									<td class=td6><SELECT NAME="selType" ALT="담당자" STYLE="Width: 98px;" tag="11" ><Option value = ""></Option></Select></td>
									<TD CLASS=TD5 NOWRAP>설비설치일자</TD>
									<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtReqdlvyFromDt" CLASS=FPDTYYYYMMDD tag="11X1" ALT="시작일자" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
											<TD>&nbsp;~&nbsp;</TD>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtReqdlvyToDt" CLASS=FPDTYYYYMMDD tag="11X1" ALT="종료일자" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
									</TD>																			
								</tr>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
                    <TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
                        <TABLE <%=LR_SPACE_TYPE_20%>>
                            <TR>
                                <TD WIDTH=100% >
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="13" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
                                </TD>                                
                            </TR>
                        </TABLE>
                    </TD>
                </TR>				
			
			</TABLE>	
		</TD>
		<TR>
			<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
		</TR>
		<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
					     <TD WIDTH = 10 > &nbsp; </TD>
					     <TD>
			               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON>
		                 </TD> 		
 			        </TR>
			    </TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
			</TD>
		</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="hreceivercd"    TAG="24">
<INPUT TYPE=HIDDEN NAME="htreater"		 TAG="24">
<INPUT TYPE=HIDDEN NAME="hCastCd"		 TAG="24">

</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
