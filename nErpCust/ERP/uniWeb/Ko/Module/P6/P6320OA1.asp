<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Production
*  2. Function Name        : 금형관리대장(HB)
*  3. Program ID           : P6320OA1
*  4. Program Name         : 금형관리대장(HB)
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2005/06/25
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
Const BIZ_PGM_ID       = "P6320OB1.asp"						           '☆: Biz Logic ASP Name



Dim C_CAST_CD
Dim C_CAST_NM
Dim C_ITEM_CD
Dim C_ITEM_NM
Dim C_ITEM_FORMAL_NM
Dim C_EMP_CD
Dim C_MAKE_DT
Dim C_MAKER
Dim C_PRS_UNIT
Dim C_SPEC		
Dim C_MAT_Q
Dim C_CUR_ACCNT
Dim C_CUSTODY_AREA
Dim C_CLOSE_DT
Dim C_BIGO					

Const C_SHEETMAXROWS = 30

Dim iDBSYSDate
Dim EndDate, StartDate,StartDate_,EndDate_

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
	
	C_CAST_CD				= 1
	C_CAST_NM				= 2
	C_ITEM_CD				= 3
	C_ITEM_NM				= 4
	C_ITEM_FORMAL_NM		= 5
	C_EMP_CD				= 6
	C_MAKE_DT				= 7
	C_MAKER	 				= 8
	C_PRS_UNIT				= 9
	C_SPEC					= 10
	C_MAT_Q					= 11
	C_CUR_ACCNT				= 12
	C_CUSTODY_AREA			= 13
	C_CLOSE_DT				= 14
	C_BIGO					= 15

End Sub
	

Sub SetDefaultVal()		

	frm1.txtReqdlvyFromDt.text = StartDate
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
		
				
		ggoSpread.SSSetEdit    C_CAST_CD			,		"금형코드"			,	15
		ggoSpread.SSSetEdit    C_CAST_NM			,		"금형명칭"			,	15
		ggoSpread.SSSetEdit    C_ITEM_CD			,		"품목코드"			,	15
		ggoSpread.SSSetEdit    C_ITEM_NM			,		"품목명"			,	15
		ggoSpread.SSSetEdit    C_ITEM_FORMAL_NM		,		"모델명"			,	15
		ggoSpread.SSSetEdit    C_EMP_CD				,		"담당자"			,	15
		ggoSpread.SSSetDate    C_MAKE_DT			,		"금형완성일"		,	11, 2, parent.gDateFormat	
		ggoSpread.SSSetEdit    C_MAKER				,		"제작업체"			,   15		
		ggoSpread.SSSetFloat   C_PRS_UNIT			,		"Cavity"			,	15, "6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"    	
		ggoSpread.SSSetEdit    C_SPEC				,		"규격"				,	20
		ggoSpread.SSSetEdit    C_MAT_Q				,		"재질"				,	15
		ggoSpread.SSSetFloat   C_CUR_ACCNT			,		"현재타수"			,	15, "6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"    	
		ggoSpread.SSSetEdit    C_CUSTODY_AREA		,		"보관장소"			,	10
		ggoSpread.SSSetDate    C_CLOSE_DT			,		"금형폐기일자"		,	11, 2, parent.gDateFormat	
		ggoSpread.SSSetEdit    C_BIGO				,		"비고"				,	10
	
		
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
		
			C_CAST_CD			= iCurColumnPos(1)
			C_CAST_NM			= iCurColumnPos(2)
			C_ITEM_CD			= iCurColumnPos(3)
			C_ITEM_NM			= iCurColumnPos(4)
			C_ITEM_FORMAL_NM	= iCurColumnPos(5)			
			C_EMP_CD			= iCurColumnPos(6)
			C_MAKE_DT			= iCurColumnPos(7)
			C_MAKER				= iCurColumnPos(8)
			C_PRS_UNIT			= iCurColumnPos(9)
			C_SPEC				= iCurColumnPos(10)	
			C_MAT_Q				= iCurColumnPos(11)						
			C_CUR_ACCNT			= iCurColumnPos(12)
			C_CUSTODY_AREA		= iCurColumnPos(13)
			C_CLOSE_DT			= iCurColumnPos(14)
			C_BIGO				= iCurColumnPos(15)
									
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

	frm1.txtCastCd.focus	
	
	
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
     
    FncQuery = True                                                              '☜: Processing is OK
																'⊙: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
  
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
   
    
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
	 Dim iDx
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
	 Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

   
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
    
    With frm1

	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001		
		strVal = strVal & "&txtReqdlvyFromDt=" & Trim(.txtReqdlvyFromDt.text)
		strVal = strVal & "&txtReqdlvyToDt=" & Trim(.txtReqdlvyToDt.text)		
		strVal = strVal & "&txtFormalNm=" & Trim(.txtFormalNm.value)
		strVal = strVal & "&txtCastCd=" & Trim(.txtCastCd.value)	
		strVal = strVal & "&txtCastNM=" & Trim(.txtCastNM.value)
		strVal = strVal & "&seltype=" & Trim(.seltype.value)
		strVal = strVal & "&txtCustArea=" & Trim(.txtCustArea.value)

	
 	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 
	        
    End With
	    
    DbQuery = True 

End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery2(Row)
     
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
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
	lgKeyStream = Frm1.txtAsNo.Value & parent.gColSep       'You Must append one character(parent.gColSep)


    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	
	
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
    'Call InitVariables
   ' Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call FncQuery()	
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
		arrParam(2) = Trim(frm1.txtCastCd.value)		<%' Code Condition%>
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
		
		frm1.txtCastCd.value = Trim(arrRet(0))
		frm1.txtCastNM.value = Trim(arrRet(1))	
			
	End Select
	
	lgBlnFlgChgValue=true
	

End Function


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    Call SetPopupMenuItemInf("0000111111")         '화면별 설정   
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
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

    If Button = 2 And gMouseClickStatus = "SPC" Then
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
'   Event Name : FncBtnPrint()
'   Event Desc : 
'==========================================================================================
Function FncBtnPrint()

    Dim StrEbrFile, condvar

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	If lgIntFlgMode = parent.OPMD_CMODE Then						'/조회여부 확인 
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If
	  Call PrintCond(strEbrFile, condvar)
	  Call FncEBRprint(EBAction, StrEbrFile, condvar) 	

End Function


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

sub selType_OnChange
	
end sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>금형관리대장</font></td>
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
									<TD CLASS=TD5 NOWRAP>금형코드</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCastCd" ALT="금형코드" TYPE="Text" MAXLENGTH="13" SIZE=20 tag="11XXXU"></td>
									<TD CLASS=TD5 NOWRAP>금형명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCastNM" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="11"></TD>	
									
								</TR>
								<Tr>
									<Td class=td5>담당자</td>
									<td class=td6><SELECT NAME="selType" ALT="담당자" STYLE="Width: 98px;" tag="11" onChange="VBScript:selType_OnChange"><Option value= ""></Option></Select></td>
									<TD CLASS=TD5 NOWRAP>금형완성일자</TD>
									<TD CLASS=TD6 NOWRAP>
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
												<script language =javascript src='./js/p6320oa1_fpDateTime1_txtReqdlvyFromDt.js'></script>
											</TD>
											<TD>&nbsp;~&nbsp;</TD>
											<TD>
												<script language =javascript src='./js/p6320oa1_fpDateTime2_txtReqdlvyToDt.js'></script>
											</TD>
										</TR>
									</TABLE>
									</TD>										
									
								</tr>
									<TD CLASS=TD5 NOWRAP>모델명</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT NAME="txtFormalNm" TYPE="Text" MAXLENGTH="40" SIZE=20 tag="11">
									</TD>
									<TD CLASS=TD5 NOWRAP>이관장소</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCustArea" ALT="이관장소" TYPE="Text" MAXLENGTH="13" SIZE=20 tag="11XXXU"></td>								
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
                                    <script language =javascript src='./js/p6320oa1_OBJECT1_vspdData.js'></script>
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

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
