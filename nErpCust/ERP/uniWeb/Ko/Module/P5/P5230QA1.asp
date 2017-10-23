<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : p5
*  2. Function Name        : 설비수리내역조회(HB)
*  3. Program ID           : P5210QA1
*  4. Program Name         : 설비수리내역조회(HB)
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2005/07/20
*  9. Modifier (First)     : Joo Young Hoon
* 10. Modifier (Last)      : Chen, Jae Hyun
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


<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID       = "P5230QB1.asp"						           '☆: Biz Logic ASP Name



Dim C_FAC_CAST_CD
Dim C_FACILITY_NM
Dim C_MINOR_NM
Dim C_WORK_DT
Dim C_MINOR_NM2
Dim C_INSP_TEXT
Dim C_BP_NM
Dim C_NAME	
Dim C_BIGO		

Dim fromdate,todate			
				

Const C_SHEETMAXROWS = 30

Dim iDBSYSDate
Dim EndDate, StartDate,Act_Row

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, Parent.gServerDateFormat, Parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = DateAdd("d", -7, EndDate)



'Const IMG_LOAD_PATH = "../../ComAsp/imgTemp.asp?src="
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop,selChk

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  	
	
	C_FAC_CAST_CD				= 1
	C_FACILITY_NM				= 2
	C_MINOR_NM					= 3
	C_WORK_DT					= 4
	C_MINOR_NM2					= 5
	C_INSP_TEXT					= 6
	C_BP_NM						= 7
	C_NAME						= 8	
	C_BIGO						= 9

End Sub

Sub SetDefaultVal()		

	frm1.txtReqdlvyFromDt.text = StartDate
	frm1.txtReqdlvyToDt.text = Enddate		
	
End Sub


Sub SetDefaultVal2()		

	frm1.txtReqdlvyFromDt.text = fromdate
	frm1.txtReqdlvyToDt.text = todate	
	
End Sub


'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
 Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	
	ggoSpread.Source = frm1.vspdData
	
	ggoSpread.Spreadinit	"V20021105",, parent.gAllowDragDropSpread    
	
	
	With frm1.vspdData
		
		.ReDraw = False
		  
		.MaxCols = C_BIGO + 1
		.MaxRows = 0
		
		
		Call ggoSpread.ClearSpreadData()	
		
		'.OperationMode = 3
		
		Call GetSpreadColumnPos("A")
		
				
		ggoSpread.SSSetEdit    C_FAC_CAST_CD			,		"설비코드"			,	10
		ggoSpread.SSSetEdit    C_FACILITY_NM			,		"설비명"			,	15
		ggoSpread.SSSetEdit    C_MINOR_NM				,		"설비유형"			,	10
		ggoSpread.SSSetDate    C_WORK_DT				,		"작업일자"			, 11, 2, parent.gDateFormat
		ggoSpread.SSSetEdit    C_MINOR_NM2				,		"수리부위"			,	10	
		ggoSpread.SSSetEdit    C_INSP_TEXT				,		"점검내역"			,   40		
		ggoSpread.SSSetEdit    C_BP_NM					,		"거래처"			,	10
		ggoSpread.SSSetEdit    C_NAME					,		"작업자"			,	10		
		ggoSpread.SSSetEdit    C_BIGO					,		"비고"				,	10
	
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
				
		
		.ReDraw = true
		
	End With	
	
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If           
       Next          
    End If   
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
		
			C_FAC_CAST_CD			= iCurColumnPos(1)
			C_FACILITY_NM			= iCurColumnPos(2)
			C_MINOR_NM				= iCurColumnPos(3)			
			C_WORK_DT				= iCurColumnPos(4)
			C_MINOR_NM2				= iCurColumnPos(5)
			C_INSP_TEXT				= iCurColumnPos(6)
			C_BP_NM					= iCurColumnPos(7)
			C_NAME					= iCurColumnPos(8)			
			C_BIGO					= iCurColumnPos(9)
									
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
	
	selChk=false
	
	Dim IntRetCD 
    Dim RetStatus

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call ggoSpread.ClearSpreadData()	
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    Call InitVariables                                                       '⊙: Initializes local global variables
	
	'Call ggoOper.ClearField(Document, "1")
    Call ggoOper.ClearField(Document, "2")
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery() = False Then
        Call RestoreToolBar()
        Exit Function
    End If
     

    call SetDefaultVal2
    
    
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
	
	On Error Resume Next
	
    FncNew = True																 '☜: Processing is OK
    
    
    
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
	On Error Resume Next	
    FncDelete = True                                                            '☜: Processing is OK
    
   
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    FncSave = False                                                              '☜: Processing is NG
    On Error Resume Next
    FncSave = True                                                              '☜: Processing is NG
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
   On Error Resume Next
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
	
    FncDeleteRow = False                                                         '☜: Processing is NG
    On Error Resume Next
    FncDeleteRow = True                                                          '☜: Processing is OK
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
    
    fromdate=Trim(frm1.txtReqdlvyFromDt.text)
    todate=Trim(frm1.txtReqdlvyToDt.text)
    
    With frm1

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001	
		strVal = strVal & "&txtReqdlvyFromDt=" & Trim(.txtReqdlvyFromDt.text)
		strVal = strVal & "&txtReqdlvyToDt=" & Trim(.txtReqdlvyToDt.text)
		strVal = strVal & "&txtCastCd=" & Trim(.txtCastCd.value)	
		strVal = strVal & "&selType=" & Trim(.selType.value)
		
 	 Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	        
    End With
	    
    DbQuery = True 

End Function


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

'===========================================================================
' Function Name : OpenCondRepairNo
' Function Desc : OpenCondRepairNo Reference Popup
'===========================================================================
Function OpenCondRepairNo()

	Dim arrRet
	Dim arrParam 
	Dim iCalledAspName
	DIM IntRetCD
	Dim arrpb(0)            
	        
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("XW201RA2_KO244")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "XW201RA_KO244", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent, arrParam ), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0)= "" Then		
		Exit Function
	Else	
		frm1.txtCondRepairNo.value = Trim(arrRet(0))
		Call mainquery()
	End if	

End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================


'==============================================================================
' 현재 경로가 수정되면 
'==============================================================================




Sub SetSpreadColor()

	Dim lsi		

		For lsi = 1 to frm1.vspdData.maxrows
		

			frm1.vspdData.col = 1
			frm1.vspddata.row = lsi
			
			if trim(frm1.vspddata.text) = "품목그룹소계" Then
				frm1.vspddata.col = -1
				frm1.vspddata.row = lsi
				frm1.vspddata.BackColor = RGB(204,255,153) '연두 
			elseif trim(frm1.vspddata.text) = "합계" then
				frm1.vspddata.col = -1
				frm1.vspddata.row = lsi
				frm1.vspddata.BackColor = RGB(176,234,244) '하늘색 
			end if
			
		'	frm1.vspdData.col=2
		'	frm1.vspddata.row=lsi
		'	
		'	if trim(frm1.vspddata.text) = "품목소계" then
		'		frm1.vspddata.col = -1
		'		frm1.vspddata.row = lsi
		'		frm1.vspddata.BackColor = RGB(255,255,0) '노란색  
		'	End if
		Next

    
End Sub

'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================

function check_img_path(path)

	if (ggoSaveFile.fileExists(path) = 0)  then
		Check_img_path = true
	else
		Check_img_path = false
	end if
end function


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
	
		arrParam(0) = "설비코드조회"					<%' 팝업 명칭 %>
		arrParam(1) ="Y_FACILITY"	<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtCastCd.value)		<%' Code Condition%>
		'arrParam(3) = Trim(frm1.txtDn_TypeNm.value)		<%' Name Cindition%>
		arrParam(4) = " " 
		arrParam(5) = "설비코드"			  	   <%' TextBox 명칭 %>

		arrField(0) = "FACILITY_CD"							<%' Field명(0)%>
		arrField(1) = "FACILITY_NM"							<%' Field명(1)%>

		arrHeader(0) = "설비코드"					<%' Header명(0)%>
		arrHeader(1) = "설비명칭"					<%' Header명(1)%>

			 
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
		
		frm1.txtCastCd.value = Trim(arrRet(0))
		frm1.txtCastNM.value = Trim(arrRet(1))	
			
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
 	End If
    
End Sub

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

End Sub


sub selProcessType_OnChange
	lgBlnFlgChgValue=true
End Sub 

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub


'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
	
	Call CommonQueryRs(" minor_cd,minor_nm "," B_MINOR "," major_Cd = " & FilterVar("Z410", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.selType, lgF0, lgF1, Chr(11))
    
    frm1.selType.value = ""
    
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

</HEAD>

<BODY TABINDEX="-1" SCROLL="YES">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>설비수리내역조회</font></td>
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
								<Tr>									
									<TD CLASS=TD5 NOWRAP>작업일자</TD>
									<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtReqdlvyFromDt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="시작일자" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
											<TD>&nbsp;~&nbsp;</TD>
											<TD>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtReqdlvyToDt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="종료일자" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>
									</TD>	
									<td class="td5">설비코드</td>
									<td class="td6"><INPUT NAME="txtCastCd" ALT="설비코드" TYPE="Text" MAXLENGTH="13" SIZE=10 tag="11XXXU"><IMG SRC="../../image/btnPopup.gif" NAME="btnDnHDR" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRequried 1">&nbsp;<INPUT NAME="txtCastNM" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
									
								</tr>
									<TD CLASS=TD5 NOWRAP>설비유형</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="selType" ALT="설비유형" STYLE="Width: 120px;" tag="11" onChange="VBScript:selType_OnChange"><OPTION VALUE=""></OPTION></SELECT></td>
									<TD CLASS=TD5 NOWRAP>&nbsp;</td>
									<TD CLASS=TD6 NOWRAP>&nbsp;</td>
									
									
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
                                <TD HEIGHT="100%" WIDTH=100% >
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="13" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
                                </TD>                                
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
				<TR>
					<TD WIDTH=100% VALIGN=TOP>
						
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
<INPUT TYPE=HIDDEN NAME="hcastcd"		 TAG="24">

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