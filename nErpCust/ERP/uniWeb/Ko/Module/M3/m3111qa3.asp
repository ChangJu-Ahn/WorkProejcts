<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111qa3
'*  4. Program Name         : 발주번호별진행조회 
'*  5. Program Desc         : 발주번호별진행조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : shin Jin Hyun
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
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 
Dim lgStrPrevKey_A                                          <%'☜: Next Key tag                          %>
Dim lgSortKey_A                                             <%'☜: Sort상태 저장변수                     %> 
Dim lgStrPrevKey_B                                          <%'☜: Next Key tag                          %>
Dim lgSortKey_B                                             <%'☜: Sort상태 저장변수                     %> 

Dim lgKeyPos                                                <%'☜: Key위치                               %>
Dim lgKeyPosVal                                             <%'☜: Key위치 Value                         %>
Dim	lgTopLeft
Dim lgKeyPoNo
Dim IscookieSplit 
Dim lgSaveRow                                           

Dim Query_Msg_Flg
Dim StartDate
Dim EndDate
Dim lgPageNo2

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  


Const BIZ_PGM_ID 		= "m3111qb3.asp"  
Const BIZ_PGM_ID1       = "m3111qb4.asp"
Const BIZ_PGM_JUMP_ID 	= "m3111ma1"
Const BIZ_PGM_JUMP_ID1 	= "m3111ma7"
		             
Const C_MaxKey			  = 20			

'#########################################################################################################
'												2. Function부 
'#########################################################################################################
Function setCookie()

	if lgKeyPoNo <> "" then
		WriteCookie "PoNo", lgKeyPoNo
	end if
	
	if UCase(Trim(frm1.hdnretflg.value)) = "Y" then
	    Call PgmJump(BIZ_PGM_JUMP_ID1)
    else
        Call PgmJump(BIZ_PGM_JUMP_ID)
    end if

End Function

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
	lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey_A   = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1
    lgStrPrevKey_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1
	lgIntFlgMode = Parent.OPMD_CMODE
	Query_Msg_Flg		= False
    lgPageNo         = ""
	lgPageNo2		 = ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtPoFrDt.Text	= StartDate 
	frm1.txtPoToDt.Text	= EndDate 
	frm1.txtBpCd.focus
	Set gActiveElement = document.activeElement
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA")%>
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M3111QA3","S","A","V20030319", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    Call SetZAdoSpreadSheet("M3111QA301","S","B","V20030319", Parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A") 
	Call SetSpreadLock("B") 
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
    Else
      ggoSpread.Source = frm1.vspdData2
      ggoSpread.SpreadLockWithOddEvenRowColor()
    End If   
End Sub

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : Purchase_Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
	Dim IntRetCD
			
	If lgIsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	lgIsOpenPop = True
		
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag

	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function

'------------------------------------------  OpenSubPopup()  -------------------------------------------------
'	Name : OpenSubPopup()
'	Description : Supplier PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenSubPopup()
	Dim strRet
	Dim arrParam(4)
	Dim iCalledAspName
	Dim IntRetCD
		
	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True
	
	frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
	arrParam(0) = GetKeyPosVal("A",1)				'po_no
	frm1.vspdData2.Col = 1
	arrParam(1) = Trim(frm1.vspdData2.Text)			'po_seq_no
	frm1.vspdData2.Col = 2
	arrParam(2) = Trim(frm1.vspdData2.Text)			'item_cd
	frm1.vspdData2.Col = 3
	arrParam(3) = Trim(frm1.vspdData2.Text)			'item_nm
	frm1.vspdData2.Col = 16
	arrParam(4) = Trim(frm1.vspdData2.Text)			'currency

	iCalledAspName = AskPRAspName("M3111PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
End Function
'------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenSppl()
'	Description : Supplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBpCd.Value)		
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "공급처"					
	
    arrField(0) = "BP_CD"						
    arrField(1) = "BP_NM"						
    
    arrHeader(0) = "공급처"					
    arrHeader(1) = "공급처명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function

'------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : PoType PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "발주형태"					
	arrParam(1) = "M_CONFIG_PROCESS"			
	arrParam(2) = Trim(frm1.txtPoType.Value)	
'	arrParam(3) = Trim(frm1.txtPoTypeNm.Value)	
	arrParam(4) = ""							
	arrParam(5) = "발주형태"					
	
    arrField(0) = "PO_TYPE_CD"						
    arrField(1) = "PO_TYPE_NM"						
        
    arrHeader(0) = "발주형태"					
    arrHeader(1) = "발주형태명"					
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoType.focus
		Exit Function
	Else
		frm1.txtPoType.Value = arrRet(0)
		frm1.txtPoTypeNm.Value = arrRet(1)
		frm1.txtPoType.focus
	End If	
End Function

'------------------------------------------  OpenPurGrp()  -------------------------------------------------
'	Name : OpenPurGrp()
'	Description : PurGrp PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus
	End If	

End Function 

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
'	Name : PopZAdoConfigGrid()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	If gActiveSpdSheet.Id="A" Then
		Call OpenGroupPopup(gActiveSpdSheet.Id)
	Else
		Call OpenOrderBy(gActiveSpdSheet.Id)
	End If
End Sub

'------------------------------------  OpenGroupPopup()  ----------------------------------------------
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData(pSpdNo),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Function

'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Function OpenOrderBy(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pSpdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


'------------------------------------  Setretflg()  ----------------------------------------------
'	Name : Setretflg()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Function Setretflg()
	
    Setretflg = False
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iretflg
    Dim iPlsFlg
	
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = 1
    Err.Clear

	Call CommonQueryRs(" ret_flg ", " m_pur_ord_hdr ", " po_no = " & FilterVar(Trim(frm1.vspdData.Text), " " , "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)


    iretflg = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, VbInformation, parent.gLogoName
		Err.Clear 
		Exit Function
	End If

    if Trim(lgF0) <> "" then
        frm1.hdnretflg.value = UCase(Trim(iretflg(0)))        
        if UCase(Trim(iretflg(0))) = "Y" then
            Setretflg = False
            Exit Function 
        end if
    end if 

    Setretflg = True
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    ReDim lgKeyPos(C_MaxKey)
    ReDim lgKeyPosVal(C_MaxKey)
	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")											
	Set gActiveElement = document.activeElement
    lblJump.innerHTML = "발주등록"
End Sub

'==========================================================================================
'   Event Name : vspdData_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub   

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Function

'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtPoFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtPoToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtPoToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoToDt.Focus
	End If
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtPoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtPoToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name : vspdData2_DblClick
'   Event Desc :
'==========================================================================================
Function vspdData2_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OpenSubPopup()
		End If
	End If
End Function
	
'========================================================================================== 
' Event Name : vspdData_LeaveCell 
' Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
		Call vspdData_Click(NewCol, NewRow)
		frm1.vspdData2.MaxRows = 0
		lgPageNo2		 = ""
		Call DbQuery("2", NewRow)
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click( Col,  Row)
    Dim ii

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 1
        End If    
        Exit Sub
    End If

	Call SetSpreadColumnValue("A",Frm1.vspdData, Col, Row)
		
    frm1.vspdData.Col = GetKeyPos("A",1)
    frm1.vspdData.Row = Row
    lgKeyPoNo      = frm1.vspdData.text
    lgStrPrevKey_B   = ""     
    lgSortKey_B      = 1

    If Setretflg() = False Then
		lblJump.innerHTML = "반품발주등록"
	Else 
		lblJump.innerHTML = "발주등록"
	End if	

End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click( Col,  Row)
	Call SetPopupMenuItemInf("00000000001")
	Set gActiveSpdSheet = frm1.vspdData2
	gMouseClickStatus = "SP2C"

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If

	Call SetSpreadColumnValue("B",Frm1.vspdData2, Col, Row)
End Sub

	
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
    	If lgPageNo <> "" Then								
			lgTopLeft = "Y"
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery("1", 0) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If			         
		End If
   End if
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo2 <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
				
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery("2", frm1.vspdData2.ActiveRow) = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
   End if
End Sub


'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear     

     '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    Call InitVariables 											
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	with frm1
		if (UniConvDateToYYYYMMDD(.txtPoFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtPoToDt.text,Parent.gDateFormat,"")) And Trim(.txtPoFrDt.text) <> "" And Trim(.txtPoToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")	
			Exit Function
		End if   
	End with

    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery("1", 0) = False then Exit Function    							

	Set gActiveElement = document.activeElement
    FncQuery = True		
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                            
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery(iOpt, currRow) 
	Dim strVal
	Dim strCfmFlg
    DbQuery = False
    
    If iOpt <> "1" and frm1.vspdData.MaxRows < 1 Then Exit Function
		
    Err.Clear                                                       
	If LayerShowHide(1) = False Then Exit Function
        
    With frm1

		if .rdoCfmFlg0.checked then
			strCfmFlg = ""
		elseif .rdoCfmFlg1.checked then
			strCfmFlg = "Y"
		else
			strCfmFlg = "N"
		end if
			
        If iOpt = "1" Then
			
			If lgIntFlgMode = Parent.OPMD_UMODE Then

				strVal = BIZ_PGM_ID & "?txtPurGrpCd=" & Trim(.hdnPurGrpCd.value)
				strVal = strVal & "&txtBpCd=" & Trim(.hdnBpCd.value)
				strVal = strVal & "&txtPoFrDt=" & Trim(.hdnPoFrDt.value)
				strVal = strVal & "&txtPoToDt=" & Trim(.hdnPoToDt.value)
				strVal = strVal & "&txtPoType=" & Trim(.hdnPoType.value)
				strVal = strVal & "&txtPoNo=" & Trim(.hdnPoNo.value)
				strVal = strVal & "&txtCfmFlg=" & Trim(.hdnstrCfmFlg.value)
		        strVal = strVal & "&lgPageNo="   & lgPageNo         
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
			else				
				strVal = BIZ_PGM_ID & "?txtPurGrpCd=" & Trim(.txtPurGrpCd.value)
				strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)
				strVal = strVal & "&txtPoFrDt=" & Trim(.txtPoFrDt.Text)
				strVal = strVal & "&txtPoToDt=" & Trim(.txtPoToDt.Text)
				strVal = strVal & "&txtPoType=" & Trim(.txtPoType.value)
				strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
				strVal = strVal & "&txtCfmFlg=" & strCfmFlg
		        strVal = strVal & "&lgPageNo="   & lgPageNo         
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

			end if
		
		Else   
			strVal = BIZ_PGM_ID1 & "?txtPoNo="	 & GetKeyPosVal("A",1)
		    strVal = strVal & "&lgPageNo="   & lgPageNo2         
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
        End If   
		
		if Query_Msg_Flg = false then
			strVal = strVal & "&Query_Msg_Flg=" & "F"
		else
			strVal = strVal & "&Query_Msg_Flg=" & "T"
		end if

        Call RunMyBizASP(MyBizASP, strVal)							
        
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")								
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk( iOpt)												
    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = Parent.OPMD_UMODE 
	
	If iOpt = 1 Then
		If lgTopLeft <> "Y" Then
			Call vspdData_Click(1, 1)
			Call DbQuery("2", 1)
		End If
		lgTopLeft = "N"
	else
		Query_Msg_Flg = true
	End If							                                     '⊙: This function lock the suitable field
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtBpCd.focus
	End If
	Set gActiveElement = document.activeElement
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
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
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>발주번호별진행</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right>&nbsp;</td>
					<TD WIDTH=10>&nbsp;</TD>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>공급처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공급처" NAME="txtBpCd"  SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSppl()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=4  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrp()">
														   <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 tag="14"></TD>					   
								</TR>					   
								<TR>
									<TD CLASS="TD5" NOWRAP>발주형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발주형태"  NAME="txtPoType" SIZE=10 LANG="ko" MAXLENGTH=5 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPoType()">
														   <INPUT TYPE=TEXT NAME="txtPoTypeNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>발주일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m3111qa3_fpDateTime2_txtPoFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m3111qa3_fpDateTime2_txtPoToDt.js'></script>
												</td>
											</tr>
										</table>
							         </TD>
	                            </TR>
	                            <TR>
									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo"  SIZE=29 MAXLENGTH=18 ALT="발주번호" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
<!--									<TD CLASS="TD5" NOWRAP>발주번호</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발주번호"  NAME="txtPoNo" SIZE=34 LANG="ko" MAXLENGTH=18 tag="11"></TD>-->
									<TD CLASS="TD5" NOWRAP>확정여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg0" CLASS="RADIO" value = "A" tag="11"><label for="rdoCfmFlg0">&nbsp;전체&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg1" CLASS="RADIO" value = "Y" tag="11"><label for="rdoCfmFlg1">&nbsp;확정&nbsp;</label>
														   <INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg2" CLASS="RADIO" value = "N" tag="11" checked><label for="rdoCfmFlg2">&nbsp;미확정&nbsp;&nbsp;</label></TD>
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
									<script language =javascript src='./js/m3111qa3_A_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/m3111qa3_B_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
    <TR>
      <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:setCookie()"><SPAN ID="lblJump">&nbsp;</SPAN></a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnPurGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnstrCfmFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnretflg" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
