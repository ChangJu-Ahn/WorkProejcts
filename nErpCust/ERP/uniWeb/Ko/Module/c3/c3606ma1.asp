
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : 실제원가관리 
'*  3. Program ID           : c3606ma1
'*  4. Program Name         : 품목별 배부내역 조회 
'*  5. Program Desc         : 품목별 배부내역 및 배부근거 조회 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/18
'*  8. Modified date(Last)  : 2001/03/5
'*  9. Modifier (First)     : Cho, Ig sung
'* 10. Modifier (Last)      : Lee Tae Soo
'* 11. Comment              : ahn do hyun =>ado변환 
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit									'☜: indicates that All variables must be declared in advance
	
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID		= "C3606MB1.asp"			'☆: 비지니스 로직 ASP명 
Const BIZ_PGM_ID1		= "C3606MB2.asp"			'☆: 비지니스 로직 ASP명 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

'Const C_SHEETMAXROWS_D_A  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 

'Const C_SHEETMAXROWS_D_B  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey            = 3                                     '☆☆☆☆: Max key value

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim lgIsOpenPop                                             '☜: Popup status                           
'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   

Dim lgPageNo_A                                              '☜: Next Key tag                          
Dim lgSortKey_A                                             '☜: Sort상태 저장변수                      

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   

Dim lgPageNo_B                                              '☜: Next Key tag                          
Dim lgSortKey_B                                             '☜: Sort상태 저장변수                      

'☜:--------Spreadsheet temp---------------------------------------------------------------------------   
                                                              '☜:--------Buffer for Spreadsheet -----   
'Dim lgKeyPos                                                '☜: Key위치                               
'Dim lgKeyPosVal                                             '☜: Key위치 Value                         


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode

    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B		 = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------
	Dim BaseDate, DateOfDB	
	Dim strYear ,strMonth ,strDay

	BaseDate	= "<%=GetSvrDate%>"                                                                  'Get DB Server Date
	DateOfDB	= UniConvDateAToB(BaseDate ,Parent.gServerDateFormat,Parent.gDateFormat)
	
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
	Call ExtractDateFrom(DateOfDB,	Parent.gDateFormat ,Parent.gComDateType ,strYear ,strMonth ,strDay) 
	frm1.txtYYYYMM.Year		= strYear
	frm1.txtYYYYMM.Month	= strMonth
	frm1.txtYYYYMM.Day		= strDay
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "C", "NOCOOKIE", "QA") %>                                '☆: 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
  	Call SetZAdoSpreadSheet("C3606MA101","G","A","V20021213",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock("A") 

	Call SetZAdoSpreadSheet("C3606MA101","G","B","V20021213",parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X","X")
    Call SetSpreadLock ("B")
    
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock( iOpt )
    If iOpt = "A" Then
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
    Else
	  ggoSpread.Source = frm1.vspdData2
      ggoSpread.SpreadLockWithOddEvenRowColor()
    End If   
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
    Call InitSpreadSheet()      
'   Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
'	Call initMinor()
End Sub



'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    '--------- Developer Coding Part (Start) ----------------------------------------------------------

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                      ' ⊙: Lock  Suitable  Field
   

	Call InitVariables														'⊙: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()

    Call SetToolbar("11000000000011")								        '⊙: 버튼 툴바 제어 
    frm1.txtYYYYMM.focus 
    frm1.txtTotAmt.allownull = False
    frm1.txtTotWorkinAmt.allownull = False
    frm1.txtTotItemAmt.allownull = False
    frm1.txtTotWorkinAmtSum.allownull = False
    frm1.txtTotItemAmtSum.allownull = False
    '--------- Developer Coding Part (End  ) ----------------------------------------------------------
    Set gActiveElement = document.activeElement 
End Sub

'========================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
    
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
    Call ggoOper.ClearField(Document, "2")									'⊙: Clear Contents  Field
   
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		

    Call InitVariables 														'⊙: Initializes local global variables
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    frm1.txtTotAmt = 0
	frm1.txtTotWorkinAmt = 0
	frm1.txtTotItemAmt = 0
    '--------- Developer Coding Part (End) ----------------------------------------------------------

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								        '⊙: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery("MQ") = False Then   
       Exit Function           
    End If     							

    FncQuery = True		
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

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


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect) 
	Dim strVal
	
                                                                       '☜: Clear err status
    On Error Resume Next
    Err.Clear 
    
    DbQuery = False                                                              '☜: Processing is NG

    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    Select Case pDirect
        Case "MQ","MN"

                With Frm1
        		    strVal = BIZ_PGM_ID  & "?txtMode="          & Parent.UID_M0001						         
                    strVal = strVal      & "&txtCostCd="		& .txtCostCd.Value              '☜: Query Key
                    strVal = strVal      & "&txtYYYYMM="		&  frm1.txtYYYYMM.Year & Right("0" & frm1.txtYYYYMM.Month,2)           '☜: Query Key
                    strVal = strVal      & "&txtMaxRows="       & .vspdData.MaxRows

    '--------- Developer Coding Part (End) ----------------------------------------------------------
                    strVal = strVal      & "&lgPageNo="          & lgPageNo_A                          '☜: Next key tag
                    strVal = strVal      & "&lgSelectListDT="    & GetSQLSelectListDataType("A")
                    strVal = strVal      & "&lgTailList="        & MakeSQLGroupOrderByList("A")
                    strVal = strVal      & "&lgSelectList="      & EnCoding(GetSQLSelectList("A"))
'                   strVal = strVal      & "&lgMaxCount="        & CStr(C_SHEETMAXROWS_D_A)            '☜: 한번에 가져올수 있는 데이타 건수 
                End With

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        Case "M1Q","M1N"

                With Frm1
        		    strVal = BIZ_PGM_ID1 & "?txtMode="           & Parent.UID_M0001						         
                    strVal = strVal      & "&txtCostCd="		 & GetKeyPosVal("A",1)                      '☜: Query Key
                    strVal = strVal      & "&txtAcctCd="		 & GetKeyPosVal("A",3)                      '☜: Query Key
                    strVal = strVal      & "&txtDiflag="		 & GetKeyPosVal("A",2)                      '☜: Query Key
                    strVal = strVal      & "&txtYYYYMM="		 & frm1.txtYYYYMM.Year & Right("0" & frm1.txtYYYYMM.Month,2)           '☜: Query Key
                    strVal = strVal      & "&txtMaxRows="        & .vspdData2.MaxRows

    '--------- Developer Coding Part (End) ----------------------------------------------------------
                    strVal = strVal      & "&lgPageNo="          & lgPageNo_B                          '☜: Next key tag
                    strVal = strVal      & "&lgSelectListDT="    & GetSQLSelectListDataType("B")
                    strVal = strVal      & "&lgTailList="        & MakeSQLGroupOrderByList("B")
                    strVal = strVal      & "&lgSelectList="      & EnCoding(GetSQLSelectList("B"))
'                   strVal = strVal      & "&lgMaxCount="        & CStr(C_SHEETMAXROWS_D_B)            '☜: 한번에 가져올수 있는 데이타 건수 
                End With
				
		End Select

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk( iOpt)											 '☆: 조회 성공후 실행로직 
	
    lgIntFlgMode     = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
	If iOpt = 1 Then
		frm1.vspdData.focus
       Call vspdData_Click(1,1)
	End If							                                     '⊙: This function lock the suitable field

	Call ggoOper.LockField(Document, "Q")								 '⊙: This function lock the suitable field 
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'	Name : OpenConItemCd()
'	Description : Item PopUp
'========================================================================================================
Function OpenConItemCd()
End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	Dim gPos
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       Case "VSPDDATA2"                  
	            gPos = "B"
    End Select     
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(gPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "3")	   
	   Exit Function        
	Else
	   Call ggoSpread.SaveXMLData(gPos,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	gMouseClickStatus = "SPC"	'Split 상태코드    

    If Row <> NewRow And NewRow > 0 Then	
	
		Call SetSpreadColumnValue("A", frm1.vspdData, NewCol, NewRow)
	
		IF Row <> 0 Then
			 ggoSpread.Source = frm1.vspdData2
			 ggoSpread.ClearSpreadData
		     lgPageNo_B       = ""                                  'initializes Previous Key
		     lgSortKey_B      = 1
		'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
			 frm1.txtTotWorkinAmt = 0
			 frm1.txtTotItemAmt = 0
	
			Call DbQuery("M1Q")
		ENd IF
	End If    
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_Click( ByVal Col, ByVal Row)
 '   Dim ii
	Call SetPopupMenuItemInf("00000000001") 
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
 
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
'	 For ii = 1 to UBound(lgKeyPos)
'        frm1.vspdData.Col = lgKeyPos(ii)
'        frm1.vspdData.Row = Row
'        lgKeyPosVal(ii)   = frm1.vspdData.text
'	 Next

	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
	
	IF Row <> 0 Then
		 ggoSpread.Source = frm1.vspdData2
		 ggoSpread.ClearSpreadData
	     lgPageNo_B       = ""                                  'initializes Previous Key
	     lgSortKey_B      = 1
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
		 frm1.txtTotWorkinAmt = 0
		 frm1.txtTotItemAmt = 0
	
		Call DbQuery("M1Q")
	ENd IF
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
'    Dim ii

	Call SetPopupMenuItemInf("00000000001") 
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2

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
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    Call SetSpreadColumnValue("B", frm1.vspdData2, Col, Row)
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

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : This function is called where spread sheet column width change
'========================================================================================================
sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData2
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_A <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
		End If
   End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery("M1Q") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
		End If
   End if
    
End Sub

'========================================================================================================
'   Event Name : txtYYYYMM_DblClick
'   Event Desc :
'=========================================================================================================
Sub txtYYYYMM_DblClick(Button)
	If Button = 1 then
       frm1.txtYYYYMM.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtYYYYMM.Focus
	End if
End Sub


'=======================================================================================================
'   Event Name : txtYYYYMM_Keypress(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtYYYYMM_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'=======================================================================================================
'	Name : OpenCost()
'	Description : Condition Plant PopUp
'=======================================================================================================
Function OpenCost()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "코스트센타팝업"			'팝업 명칭 
	arrParam(1) = "b_cost_center"				'TABLE 명칭 
	arrParam(2) = Trim(frm1.txtCostCd.Value)	'Code Condition
	arrParam(3) = ""							'Name Cindition
	arrParam(4) = "cost_type = " & FilterVar("M", "''", "S") & " "				'Where Condition
	arrParam(5) = "코스트센타"				'TextBox 명칭 
	
    arrField(0) = "cost_cd"						'Field명(0)
    arrField(1) = "cost_nm"						'Field명(1)
    
    arrHeader(0) = "코스트센타"				'Header명(0)
    arrHeader(1) = "코스트센타명"			'Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCostCd.focus
		Exit Function
	Else
		Call SetCost(arrRet)
	End If
		
End Function

'=======================================================================================================
'	Name : SetCost()
'	Description : Acct Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetCost(byval arrRet)
	frm1.txtCostCd.focus
	frm1.txtCostCd.Value = arrRet(0)		
	frm1.txtCostNm.Value = arrRet(1)		
	lgBlnFlgChgValue = True
	
End Function


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'########################################################################################################  -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별배부내역조회</font></td>
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
									<TD CLASS="TD5">작업년월</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/c3606ma1_txtYYYYMM_txtYYYYMM.js'></script>
									</TD>
									<TD CLASS="TD5">코스트센타</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtCostCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="코스트센타"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCost()">
										 <INPUT TYPE=TEXT ID="txtCostNm" NAME="txtCosttNm" SIZE=30 tag="14X">
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
					<TD WIDTH=100% HEIGHT="40%" valign=top>
						<script language =javascript src='./js/c3606ma1_I820745986_vspdData.js'></script>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS=TDT NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
									<TD CLASS=TD5 NOWRAP>합계액</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3606ma1_fpDoubleSingle1_txtTotAmt.js'></script>
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
						<script language =javascript src='./js/c3606ma1_I844672524_vspdData2.js'></script>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>재공배부 합계액</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3606ma1_fpDoubleSingle2_txtTotWorkinAmt.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>합계액</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3606ma1_fpDoubleSingle3_txtTotItemAmt.js'></script>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>총 재공배부 합계액</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3606ma1_fpDoubleSingle2_txtTotWorkinAmtSum.js'></script>
									</TD>
									<TD CLASS=TD5 NOWRAP>총 배부합계액</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/c3606ma1_fpDoubleSingle3_txtTotItemAmtSum.js'></script>
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
               </TR>  
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1" ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="htxtYYYYMM" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="htxtCostCd" tag="24" TABINDEX = "-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

