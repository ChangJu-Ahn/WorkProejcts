<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :  Ado query Sample with DBAgent(Multi + Multi)
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
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
<!-- #Include file="../../inc/IncSvrHTML.inc" -->


<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                  '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID		= "C3901MB1.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_SUB_ID	= "C3901MB2.asp"                         '☆: Biz logic spread sheet for #2

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Const C_MaxKey            = 2                                    '☆☆☆☆: Max key value

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop        
Dim IsOpenPop                                       '☜: Popup status                           
'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   

Dim lgPageNo_A                                              '☜: Next Key tag                          

'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   

Dim lgPageNo_B                                              '☜: Next Key tag                          
Dim lgPageNo_C                                              '☜: Next Key tag                          



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
    lgPageNo_B       = ""                                  'initializes Previous Key for spreadsheet #2
    lgPageNo_C       = ""                                  'initializes Previous Key for spreadsheet #2

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------

	Dim StartDate
	StartDate	= "<%=GetSvrDate%>"

	frm1.txtYYYYMM.text	= UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat)
	Call ggoOper.FormatDate(frm1.txtYYYYMM, Parent.gDateFormat, 2)
	
	
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

	<% Call loadInfTB19029A("Q", "C","NOCOOKIE","QA") %>                                '☆: 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("C3901MA1","S","A", "V20021201", Parent.C_SORT_DBAGENT,frm1.vspdData ,C_MaxKey, "X", "X")
    Call SetZAdoSpreadSheet("C3901MA1","S","B", "V20021201", Parent.C_SORT_DBAGENT,frm1.vspdData1,C_MaxKey, "X", "X")
    Call SetZAdoSpreadSheet("C3901MA1","S","C", "V20021201", Parent.C_SORT_DBAGENT,frm1.vspdData2,C_MaxKey, "X", "X")

    Call SetSpreadLock ("A")
    Call SetSpreadLock ("B")
    Call SetSpreadLock ("C")
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock(ByVal pOpt )
    If pOpt = "A" Then
       With frm1
          .vspdData.ReDraw = False
          ggoSpread.Source = .vspdData 
          ggoSpread.SpreadLock 1 , -1
          .vspdData.ReDraw = True
       End With
    ElseIf pOpt = "B" Then
       With frm1
            .vspdData1.ReDraw = False
            ggoSpread.Source = .vspdData1 
            ggoSpread.SpreadLock 1, -1
            .vspdData1.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2 
            ggoSpread.SpreadLock 1, -1
            .vspdData2.ReDraw = True
       End With
    End If   
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
	Call InitComboBox
    Call InitSpreadSheet()
    Call SetToolbar("11000000000011")								        '⊙: 버튼 툴바 제어 
	frm1.txtYyyymm.focus

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

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False                                                              '⊙: Processing is NG

    Call ggoOper.ClearField(Document, "2")								          '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.ClearSpreadData()
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.ClearSpreadData()
    
    Call InitVariables                                                            '⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then	                                          '⊙: This function check indispensable field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If DbQuery("MQ") = False Then   
       Exit Function           
    End If     							
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncQuery = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False																  '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncNew = True                                                              '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncDelete = True                                                           '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False                                                               '☜: Processing is NG
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncSave = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncCopy = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncCancel = True                                                           '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncInsertRow = False                                                          '☜: Processing is NG

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                        '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncDeleteRow = True                                                        '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncPrint = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncPrev = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncNext = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                             '☜: Processing is NG

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncExport(Parent.C_MULTI)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncExcel = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncFind(Parent.C_MULTI, True)

    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncFind = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

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

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncExit = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

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
	Dim strYear,strMonth,strDay, strProcurType

	On Error Resume Next                                                          '☜: If process fails
	Err.Clear                                                                     '☜: Clear error status

	DbQuery = False                                                              '☜: Processing is NG

	Call DisableToolBar(parent.TBC_QUERY)                                        '☜: Disable Query Button Of ToolBar
	Call LayerShowHide(1)                                                        '☜: Show Processing Message

	'--------- Developer Coding Part (Start) ----------------------------------------------------------

	Call ExtractDateFrom(Frm1.txtYyyymm.Text,Frm1.txtYyyymm.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

	If Trim(frm1.cboProcurType.value) = "P" Then
		strProcurType	= "(" & FilterVar("P", "''", "S") & " )"
	ElseIf Trim(frm1.cboProcurType.value) = "M" Then
		strProcurType	= "(" & FilterVar("M", "''", "S") & " ," & FilterVar("O", "''", "S") & " )"
	Else
		strProcurType	= ""
	End If
	
	Select Case pDirect
	    Case "MQ","MN"
            With Frm1
    		    strVal = BIZ_PGM_ID  & "?txtMode="		& parent.UID_M0001						         
                strVal = strVal      & "&txtYyyymm="	& strYear&strMonth
                strVal = strVal      & "&txtPlantCd="	& .txtPlantCD.value
                strVal = strVal      & "&strProcurType="	& strProcurType
                strVal = strVal      & "&txtItemAcct="	& .txtItemAcctCd.value
                strVal = strVal      & "&txtItemCd="	& .txtItemCd.value
                strVal = strVal      & "&txtMaxRows="	& .vspdData.MaxRows


'--------- Developer Coding Part (End) ----------------------------------------------------------
                strVal = strVal      & "&lgPageNo="          & lgPageNo_A                          '☜: Next key tag
                strVal = strVal      & "&lgSelectListDT="    & GetSQLSelectListDataType("A")
                strVal = strVal      & "&lgTailList="        & MakeSQLGroupOrderByList("A")
                strVal = strVal      & "&lgSelectList="      & EnCoding(GetSQLSelectList("A"))
            End With

	'--------- Developer Coding Part (Start) ----------------------------------------------------------
	    Case "M1Q","M1N"
            With Frm1
    		    strVal = BIZ_PGM_SUB_ID & "?txtMode="	& parent.UID_M0001						         
                strVal = strVal      & "&txtYyyymm="	& strYear&strMonth
                strVal = strVal      & "&txtPlantCd="	& GetKeyPosVal("A",1)
                strVal = strVal      & "&txtItemCd="	& GetKeyPosVal("A",2)
                strVal = strVal      & "&txtMaxRows="	& .vspdData1.MaxRows

'--------- Developer Coding Part (End) ----------------------------------------------------------
                strVal = strVal      & "&lgPageNo="          & lgPageNo_B                          '☜: Next key tag
                strVal = strVal      & "&lgSelectListDT="    & GetSQLSelectListDataType("B")
                strVal = strVal      & "&lgTailList="        & MakeSQLGroupOrderByList("B")
                strVal = strVal      & "&lgSelectList="      & EnCoding(GetSQLSelectList("B"))

                strVal = strVal      & "&txtMaxRows_C="	& .vspdData2.MaxRows
                strVal = strVal      & "&lgPageNo_C="          & lgPageNo_C                          '☜: Next key tag
                strVal = strVal      & "&lgSelectListDT_C="    & GetSQLSelectListDataType("C")
                strVal = strVal      & "&lgTailList_C="        & MakeSQLGroupOrderByList("C")
                strVal = strVal      & "&lgSelectList_C="      & EnCoding(GetSQLSelectList("C"))
           End With
	End Select		

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

	If Err.number = 0 Then
	   DbQuery = True                                                             '⊙: Processing is OK
	End If   

	Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk(ByVal pOpt)											 '☆: 조회 성공후 실행로직 
	
    lgIntFlgMode     = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode

	
	If pOpt = 1 Then
       Call vspdData_Click(1,1)
		
	End If							                                     '⊙: This function lock the suitable field

	frm1.vspdData.focus		
	Call ggoOper.LockField(Document, "Q")								 '⊙: This function lock the suitable field 

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'========================================================================================================= 
Function OpenPopup(ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
	select case iWhere
		case 1
			arrParam(0) = "공장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_PLANT"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""		' Where Condition
			arrParam(5) = "공장"			
	
			arrField(0) = "PLANT_CD"					' Field명(0)
			arrField(1) = "PLANT_NM"					' Field명(1)
    
			arrHeader(0) = "공장"				' Header명(0)
			arrHeader(1) = "공장명"				' Header명(1)
		case 2
			arrParam(0) = "품목계정 팝업"				' 팝업 명칭 
			arrParam(1) = "B_MINOR a,b_item_acct_inf b"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtItemAcctCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			IF frm1.cboProcurType.value = "P" Then
				arrParam(4) = "a.major_cd = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group in ('3RAW','4SUB','5GOODS') "
			ELSEIF  frm1.cboProcurType.value = "M" Then
				arrParam(4) = "a.major_cd = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group in ('1FINAL','2SEMI') "
			ELSE
				arrParam(4) = "a.major_cd = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group <> '6MRO' "
			END IF			

			arrParam(5) = "품목계정"			
	
			arrField(0) = "MINOR_CD"					' Field명(0)
			arrField(1) = "MINOR_NM"					' Field명(1)
    
			arrHeader(0) = "품목계정"				' Header명(0)
			arrHeader(1) = "품목계정명"				' Header명(1)
		case 3
			arrParam(0) = "품목 팝업"				' 팝업 명칭 
			arrParam(1) = "B_ITEM a,B_ITEM_BY_PLANT b"					' TABLE 명칭 
			arrParam(2) = Trim(frm1.txtItemCd.Value)	' Code Condition
			arrParam(3) = ""							' Name Cindition
			
			IF frm1.cboProcurType.value = "P" Then
				arrParam(4) = "a.item_cd = b.item_cd and b.procur_type = " & FilterVar("P", "''", "S") & "  "		' Where Condition
			ELSEIF  frm1.cboProcurType.value = "M" Then
				arrParam(4) = "a.item_cd = b.item_cd and b.procur_type <> " & FilterVar("P", "''", "S") & "  "		' Where Condition
			ELSE
				arrParam(4) = "a.item_cd = b.item_cd "		' Where Condition
			END IF
			
			IF Trim(frm1.txtItemAcctCd.value) <> "" Then
				arrParam(4) = arrParam(4) & " and b.item_acct =  " & FilterVar(frm1.txtItemAcctCd.value, "''", "S") & " "		' Where Condition
			END IF
			
			IF Trim(frm1.txtPlantCd.value) <> "" Then
				arrParam(4) = arrParam(4) & " and b.plant_cd =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "		' Where Condition
			END IF

			arrParam(5) = "품목"			
	
			arrField(0) = "a.ITEM_CD"					' Field명(0)
			arrField(1) = "a.ITEM_NM"					' Field명(1)
    
			arrHeader(0) = "품목"				' Header명(0)
			arrHeader(1) = "품목명"				' Header명(1)
	end select
		
	    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
      Select case iWhere
		case 1
			frm1.txtPlantCD.focus		
		case 2
			frm1.txtItemAcctCd.focus		
		case 3
			frm1.txtItemCd.focus
	  End Select		
		Exit Function
	Else
		Call SetReturnVal(iWhere,arrRet)
	End If	

End Function'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'========================================================================================================= 

Function SetReturnVal(byval iwhere,byval arrRet)
	With frm1
		select case iWhere	
			case 1
				.txtPlantCD.focus	
				.txtplantCd.Value	= arrRet(0)
				.txtPlantNm.Value	= arrRet(1)
			case 2
				.txtItemAcctCd.focus
				.txtItemAcctCd.Value	= arrRet(0)
				.txtItemAcctNm.Value	= arrRet(1)
			case 3
				.txtItemCd.focus
				.txtItemCd.Value	= arrRet(0)
				.txtItemNm.Value	= arrRet(1)
		end select 		
	End With

End Function


'======================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=======================================================================================================
Sub InitComboBox()
	Call SetCombo(frm1.cboProcurType, "", "")								
	Call SetCombo(frm1.cboProcurType, "P", "구매품")								
	Call SetCombo(frm1.cboProcurType, "M", "가공품")
End Sub

'==================================================================================
' Name : PopZAdoConfigGrid
' Desc :
'==================================================================================
Sub PopZAdoConfigGrid()

  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
     Exit Sub
  End If
  
  If gActiveSpdSheet.Id = "vspdData" Then
     Call OpenOrderByPopup("A")
  ElseIf gActiveSpdSheet.Id = "vspdData1" Then
     Call OpenOrderByPopup("B")
  Else
     Call OpenOrderByPopup("C")
  End If

End Sub


'========================================================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'========================================================================================================
Sub OpenOrderByPopup(ByVal pvPsdNo)
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If

End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생   16512285813101
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData

  	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       Exit Sub
    End If

    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)

    frm1.vspdData1.MaxRows = 0
    lgPageNo_B             = ""                                  'initializes Previous Key

    frm1.vspdData2.MaxRows = 0
    lgPageNo_C             = ""                                  'initializes Previous Key

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     Call DbQuery("M1Q")
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
End Sub


'=======================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

	gMouseClickStatus = "SPC"	'Split 상태코드    

    If Row <> NewRow And NewRow > 0 Then
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
		Call SetSpreadColumnValue("A", frm1.vspdData, NewCol, NewRow)
     
		ggoSpread.Source = frm1.vspdData1 
		ggoSpread.ClearSpreadData

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	
		lgPageNo_B       = ""                                  'initializes Previous Key

	    Call DbQuery("M1Q")
	End If    
	    

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    
    If Row <= 0 Then
        
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
'   Event Name : vspdData1_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData1
    

  	If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData1
       Exit Sub
    End If
    
    Call SetSpreadColumnValue("B",frm1.vspdData1,Col,Row)
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
End Sub

'========================================================================================================
'   Event Name : vspdData1_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				'8) 컬럼 title 변경 
    
    If Row <= 0 Then
        
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
End Sub


'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SP2C"
	Set gActiveSpdSheet = frm1.vspdData2
    

  	If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2
       Exit Sub
    End If
    
    Call SetSpreadColumnValue("C",frm1.vspdData2,Col,Row)
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
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
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub 


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery("M1N") = False Then
              Call RestoreToolBar()
              Exit Sub
          End if
		End If
   End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_C <> "" Then                            '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery("M1N") = False Then
              Call RestoreToolBar()
              Exit Sub
          End if
		End If
   End if
    
End Sub


'========================================================================================================
'   Event Name : txtYyyymm
'   Event Desc :
'=========================================================================================================
Sub txtYyyymm_DblClick(Button)
	If Button = 1 then
       frm1.txtYyyymm.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtYyyymm.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_KeyDown(keycode, shift)
'   Event Desc : Enter Event시 FncQuery한다.
'=======================================================================================================
Sub txtYyyymm_Keypress(ByVal KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고차이배부내역</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD>
					<TD>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>작업년월</TD> 
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/c3901ma1_OBJECT1_txtYyyymm.js'></script>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlantCD"  SIZE=10  ALT ="공장" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(1)">
														<INPUT NAME="txtPlantNM"  SIZE=30  ALT ="공장명" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>조달구분</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboProcurType" tag="11X" STYLE="WIDTH:82px:" ALT="조달구분"></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>품목계정</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemAcctCd" SIZE=10  tag="11XXXU" ALT="품목계정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(2)">
														<INPUT TYPE=TEXT NAME="txtItemAcctNm" SIZE=30 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemCd" SIZE=20  tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(3)">
														<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14"></TD>
								    <TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>
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
						<TR HEIGHT="50%">
							<TD WIDTH="100%" colspan=4>
							<script language =javascript src='./js/c3901ma1_vspdData_vspdData.js'></script>
							</TD>
						</TR>
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
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>									
									<TD CLASS="TD5" NOWRAP>재고차이금액</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/c3901ma1_fpDoubleSingle2_txtDiffAmt.js'></script>&nbsp;
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
						<TR HEIGHT="20%">
							<TD WIDTH="100%" colspan=4>
							<script language =javascript src='./js/c3901ma1_vspdData1_vspdData1.js'></script>
							</TD>
						</TR>
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
									<TD CLASS="TD5" NOWRAP>배부된 차이금액</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/c3901ma1_fpDoubleSingle2_txtAllcAmt.js'></script>&nbsp;
									</TD>
									<TD CLASS="TD5" NOWRAP>총 배부된 차이금액</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/c3901ma1_fpDoubleSingle2_txtTotAllcAmt.js'></script>&nbsp;
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
						<TR HEIGHT="30%">
							<TD WIDTH="100%" colspan=4>
							<script language =javascript src='./js/c3901ma1_vspdData2_vspdData2.js'></script>
							</TD>
						</TR>
					</TD>
				</TR>
			</TABLE>
			</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"         tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      tag="24" TABINDEX="-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>
