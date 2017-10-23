<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--'**********************************************************************************************
*  1. Module Name          : ȸ��
*  2. Function Name        : �ڻ�
*  3. Program ID           : a7116ma1
*  4. Program Name         : 
*  5. Program Desc         : �������ڻ꺰��ȸ
*  6. Comproxy List        :
*  7. Modified date(First) : 2003/12/01
*  8. Modified date(Last)  : 2005/05/04
*  9. Modifier (First)     : JooJiYeong
* 10. Modifier (Last)      : Joo,Sungho
* 11. Comment              : 
* 2005.05.04  Joo,  Sungho	  �Ⱓ(��)�� ������ �ݾװ� ����� ��ȸ ���� �߰�. �̸����� �߰�
'********************************************************************************************** -->
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
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"	 SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"      SRC="../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit									'��: indicates that All variables must be declared in advance
	
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "a7116mb1.asp"			'��: �����Ͻ� ���� ASP�� 

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 100                                          '��: Fetch max count at once
Const C_MaxKey          = 5                                           '��: key count of SpreadSheet

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Dim lgIsOpenPop
Dim lgSaveRow

' ���Ѱ��� �߰�
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' �����
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ����

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
    lgStrPrevKey		= ""
    lgPageNo			= ""
    lgIntFlgMode		= parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue	= False                    'Indicates that no value changed
	lgSortKey			= 1
	lgSaveRow			= 0

End Sub


'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)

    Call ggoOper.FormatDate(frm1.txtFr_dt, gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtTo_dt, gDateFormat, 2)
	
	frm1.txtFr_dt.focus
	frm1.txtFr_dt.Year		= strYear 		'��� default value setting
	frm1.txtFr_dt.Month		= "01"

	frm1.txtTo_dt.Year		= strYear 	
	frm1.txtTo_dt.Month		= strMonth 	

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>                                '��: 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub

'====================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================

Sub InitSpreadSheet()
    'frm1.vspdData.operationmode = 5
    Call SetZAdoSpreadSheet("A7116MA1","S","A","V20080702",Parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock("A")
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================
Sub SetSpreadLock(Byval pOpt)
	if pOpt = "A" then
		ggoSpread.Source = frm1.vspdData
	    With frm1
			.vspdData.ReDraw = False
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
	    End With
	end if
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================

Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
End Sub

'==================================================================================
' Name : PopZAdoConfigGrid
' Desc :
'==================================================================================
Sub PopZAdoConfigGrid()

	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")

End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	'Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   
    ''''Call parent.GetAdoFieldInf("a7104ma101","S","A")								' G for Group , A for SpreadSheet No('A','B',....      
    Call InitVariables																	'��: Initializes local global variables
    Call SetDefaultVal
    Call InitSpreadSheet()
    Call SetToolbar("110000000001111")													'��: ��ư ���� ����	
   '---------Developer Coding part (Start)----------------------------------------------------------------

    frm1.txtFr_dt.focus 

	' ���Ѱ��� �߰�
	Dim xmlDoc

	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc)

	' �����
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ�
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text

	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text

	' ����
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text

	Set xmlDoc = Nothing

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
Dim IntRetCd

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
       Exit Function
    End If

	If frm1.txtDeptCd.value = "" Then
		frm1.txtDeptNm.value = ""
	End If
	
	If frm1.txtAcctCd.value = "" Then
		frm1.txtAcctNm.value = ""
	End If

	If frm1.txtCondAsstNo.value = "" Then
		frm1.txtCondAsstNm.value = ""
	End If
	
	If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If frm1.txtBizAreaCd1.value = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
	If frm1.txtBizUnitCd.value = "" Then
		frm1.txtBizUnitNm.value = ""
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" And   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
	If UCase(Trim(frm1.txtBizAreaCd.value)) > UCase(Trim(frm1.txtBizAreaCd1.value)) Then
		IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
	  	frm1.txtBizAreaCd.focus
	  	Exit Function
	End If
	End If
	
    If DbQuery = False Then Exit Function

    FncQuery = True													

End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '��: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '��: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
	Dim imRow
	FncInsertRow = False
'	imRow = AskSpdSheetAddRowCount()
'	If imRow = "" then
'		Exit Function
'	End If

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If     
 With frm1
	.vspdData.focus
	ggoSpread.Source = .vspdData
	'.vspdData.EditMode = True
	.vspdData.ReDraw = False
	ggoSpread.InsertRow ,imRow
	SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	.vspdData.ReDraw = True
 End With
 Call SetToolbar("11001111001111")
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement  
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '��: Protect system from crashing
    FncPrint = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call Parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call Parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '��: Processing is OK
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

    FncExit = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '��: Processing is OK
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery()
	Dim strVal
    Dim DurYrsFg

	Err.Clear                                                               '��: Protect system from crashing
	DbQuery = False

	Call LayerShowHide(1)

    With frm1
        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	    If lgIntFlgMode <> parent.OPMD_UMODE Then

		        if .rdoDurYrsFg(0).checked then DurYrsFg = "C"
    	                if .rdoDurYrsFg(1).checked then DurYrsFg = "T"

			strVal = strVal & "?txtDeptCd="			& Trim(.txtDeptCd.value)				'��: 
			strVal = strVal & "&txtFryymm="			& Trim(frm1.txtFr_dt.year & Right("0" & frm1.txtFr_dt.month,2) )
			strVal = strVal & "&txtToyymm="			& Trim(frm1.txtTo_dt.year & Right("0" & frm1.txtTo_dt.month,2) )
			strVal = strVal & "&txtAcctCd="			& Trim(.txtAcctCd.value)
			strVal = strVal & "&txtCondAsstNo="		& Trim(.txtCondAsstNo.value)
			strVal = strVal & "&txtBizUnitCd="		& Trim(.txtBizUnitCd.value)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
			strVal = strVal & "&txtBizAreaCd_Alt="		& Trim(frm1.txtBizAreaCd.alt)
			strVal = strVal & "&txtBizAreaCd1_Alt="		& Trim(frm1.txtBizAreaCd1.alt)
			strVal = strVal & "&DurYrsFg="                  & DurYrsFg

		Else
			strVal = strVal & "?txtDeptCd="			& Trim(.htxtDeptCd.value)
			strVal = strVal & "&txtFryymm="			& Trim(frm1.txtFr_dt.year & Right("0" & frm1.txtFr_dt.month,2) )
			strVal = strVal & "&txtToyymm="			& Trim(frm1.txtTo_dt.year & Right("0" & frm1.txtTo_dt.month,2) )
			strVal = strVal & "&txtAcctCd="			& Trim(.htxtAcctCd.value)
			strVal = strVal & "&txtCondAsstNo="		& Trim(.htxtCondAsstNo.value)
			strVal = strVal & "&txtBizUnitCd="		& Trim(.htxtBizUnitCd.value)
			strVal = strVal & "&txtBizAreaCd="		& Trim(.htxtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.htxtBizAreaCd1.value)
			strVal = strVal & "&DurYrsFg="                & Trim(.hDurYrsFg.value)
			strVal = strVal & "&txtBizAreaCd_Alt="		& Trim(frm1.htxtBizAreaCd.alt)
			strVal = strVal & "&txtBizAreaCd1_Alt="		& Trim(frm1.htxtBizAreaCd1.alt)
		End If    

    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&txtinternalCd="       & frm1.hINternalCD.value 
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")		'lgSelectListDT
         
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")		'lgMaxFieldCount,lgPopUpR,parent.gFieldCD,parent.gNextSeq,parent.gTypeCD(0),parent.C_MaxSelList)
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' ���Ѱ��� �߰�
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' �����
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ�
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ����

        Call RunMyBizASP(MyBizASP, strVal)							

    End With
  
    If Err.Number = 0 Then
		DBQuery = True
	End If

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    lgSaveRow        = 1

End Function


'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================

Sub OpenOrderByPopup(ByVal pSpdNo)

	Dim arrRet
	On Error Resume Next
	If lgIsOpenPop = True Then Exit Sub

	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp", _
									Array(ggoSpread.GetXMLData("A"),gMethodText), _
									"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & _
									parent.SORTW_HEIGHT & "px; ; center: Yes; help: No; resizable: No; status: No;")
									
	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
		Exit Sub
	Else
		Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
		Call InitVariables
		Call InitSpreadSheet()
	End If

End Sub


'------------------------------------------  OpenMasterRef()  -------------------------------------------------
'	Name : OpenMasterRef()
'	Description : Asset Master Condition PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMasterRef()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If lgIsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		lgIsOpenPop = False
		Exit Function
	End If
	
	' ���Ѱ��� �߰�
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	lgIsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPoRef(arrRet)
	End If	

	
End Function

 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(strRet)
       
	frm1.txtCondAsstNo.value     = strRet(0)
	frm1.txtcondAsstNm.value	 = strRet(1)
		
End Sub


Function OpenAcctDeptPopUp(Byval strCode, Byval Cond)
	Dim arrRet
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrParam(8)

	If lgIsOpenPop = True Then Exit Function

	
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If

	
	arrParam(0) = strCode								 '  Code Condition
   	arrParam(1) = ""
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = "F"									' T / F �������� ���� Condition  

	' ���Ѱ��� �߰�
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID


	lgIsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet, 4)
	End If	
	
End Function


'------------------------------------------  OpenAcct()  -------------------------------------------------
'	Name : OpenAcct()
'	Description : Account PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcctPopUp(Byval strCode, Byval Cond)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	

	arrParam(0) = "�����ڵ��˾�"			' �˾� ��Ī 
	arrParam(1) = "a_acct"						' TABLE ��Ī 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "acct_type = 'K0'"			' Where Condition
	arrParam(5) = "�����ڵ�"				' �����ʵ��� �� ��Ī 
	
    arrField(0) = "acct_cd"						' Field��(0)
    arrField(1) = "acct_nm"						' Field��(1)
    
    arrHeader(0) = "�����ڵ�"				' Header��(0)
    arrHeader(1) = "������"					' Header��(1)
    
	lgIsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet, 3)
	End If	
	
End Function


'==========================================  SetAcct()  ==================================================
'	Name : SetAcct()
'	Description : Account Popup���� Return�Ǵ� �� setting
'=========================================================================================================
'Function SetReturnVal(ByVal arrRet, ByVal field_fg)
'	
'	Select case field_fg
'		case 3	'OpenAcctCd
'			frm1.txtAcctCd.Value		= arrRet(0)
'			frm1.txtAcctNm.Value		= arrRet(1)
'		case 4	'OpenAcctDeptPopUp
'			frm1.txtDeptCd.Value	= Trim(arrRet(0))
'			frm1.txtDeptNm.Value	= arrRet(1)
'			frm1.hOrgChangeId.value  = arrRet(2)
'			frm1.txtDeptCd.focus
'			lgBlnFlgChgValue = True
'	End select	
'
'End Function


'------------------------------------------  OpenBizUnit()  -------------------------------------------------
'	Name : OpenBizUnit()
'	Description : Account PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizUnitPopup(Byval strCode, Byval Cond)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	arrParam(0) = "������˾�"					' �˾� ��Ī 
	arrParam(1) = "b_biz_unit"					' TABLE ��Ī 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Condition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "������ڵ�"					' �����ʵ��� �� ��Ī 
	
    arrField(0) = "Biz_unit_cd"						' Field��(0)
    arrField(1) = "Biz_unit_nm"						' Field��(1)
    
    arrHeader(0) = "������ڵ�" 				' Header��(0)
    arrHeader(1) = "����θ�"					' Header��(1)
    
	lgIsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet, 5)
	End If	
	
End Function
'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition

	' ���Ѱ��� �߰�
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

	arrParam(5) = "����� �ڵ�"			

    arrField(0) = "BIZ_AREA_CD"					' Field��(0)
    arrField(1) = "BIZ_AREA_NM"					' Field��(1)

    arrHeader(0) = "������ڵ�"				' Header��(0)
	arrHeader(1) = "������"				' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If
End Function


'=======================================================================================================
'	Name : SetReturnVal()
'	Description : 
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		case 0
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 1
			frm1.txtBizAreaCd1.Value = arrRet(0)
			frm1.txtBizAreaNm1.Value = arrRet(1)
			frm1.txtBizAreaCd1.focus
		case 3	'OpenAcctCd
			frm1.txtAcctCd.Value		= arrRet(0)
			frm1.txtAcctNm.Value		= arrRet(1)
		case 4	'OpenAcctDeptPopUp
			frm1.txtDeptCd.Value	= Trim(arrRet(0))
			frm1.txtDeptNm.Value	= arrRet(1)
			frm1.hINternalCD.value  = arrRet(2)
			'frm1.txtDeptCd.focus
		case 5
			frm1.txtBizUnitCd.Value		= arrRet(0)
			frm1.txtBizUnitNm.Value		= arrRet(1)
			
			lgBlnFlgChgValue = True
			
	End Select
	
	lgBlnFlgChgValue = True

End Function


Sub txtFr_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtFr_dt.Action = 7
        frm1.txtFr_dt.focus
    End If
    lgBlnFlgChgValue = True
End Sub

Sub txtFr_dt_Change()
    lgBlnFlgChgValue = True
End Sub

Sub txtFr_dt_Keypress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Sub


Sub txtTo_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtTo_dt.Action = 7
        frm1.txtTo_dt.focus
    End If
    lgBlnFlgChgValue = True
End Sub

Sub txtTo_dt_Change()
    lgBlnFlgChgValue = True
  
End Sub

Sub txtTo_dt_Keypress(KeyAscii)
    If KeyAscii = 13 Then
        Call MainQuery()
    End If
End Sub
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
 	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub
	

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001")
	Dim ii

	gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
    
	If Row = 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort
			lgSortKey = 2
		Else
			ggoSpread.SSSort ,lgSortKey
			lgSortKey = 1
		End If    
	End If
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

Sub vspdData_Change(ByVal Col, ByVal Row)
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)  
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
'Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
'End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : 
'==========================================================================================
'Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
'End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub

'========================================================================================================
'   Event Name : txtDeptCd_Keypress()
'   Event Desc : 
'========================================================================================================
Sub txtDeptCd_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()		
	End If   
End Sub

'========================================================================================================
'   Event Name : txtAcctCd_Keypress()
'   Event Desc : 
'========================================================================================================
Sub txtAcctCd_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()		
	End If   
End Sub

'========================================================================================================
'   Event Name : txtCondAsstNo_Keypress()
'   Event Desc : 
'========================================================================================================
Sub txtCondAsstNo_Keypress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()		
	End If   
End Sub

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
Function txtCondAsstNo_onblur
	if frm1.txtCondAsstNo.value = "" then
		frm1.txtCondAsstNm.value = "" 
	end if
End Function

Function txtDeptCd_onblur
	if frm1.txtDeptCd.value  = "" then
		frm1.txtDeptNm.value = "" 
		frm1.hINternalCD.value =""
	end if
End Function

Function txtAcctCd_onblur
	if frm1.txtAcctCd.value  = "" then
		frm1.txtAcctNm.value = "" 
	end if
End Function
'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================

Sub txtDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj
	Dim strSvrDate
	Dim strCliDate
	Dim sDate

    lgBlnFlgChgValue = True
	if Trim(frm1.txtDeptCd.value) ="" then exit sub
	strSvrDate = "<%=GetSvrDate%>"	

	sDate = cdate(frm1.txtTo_dt.year & "-" & frm1.txtTo_dt.month & "-01")

		'----------------------------------------------------------------------------------------
		strSelect	=			 " top 1 dept_cd, org_change_id, internal_cd ,dept_nm"    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " &  FilterVar(Ltrim(Rtrim(frm1.txtDeptCd.value)) ,"''","S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <='" & sDate & "'))"				
	

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			If lgIntFlgMode <> parent.OPMD_UMODE Then
				IntRetCD = DisplayMsgBox("124600","X","X","X")  
			End If
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.hInternalCd.value = ""

		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				frm1.hInternalCd.value =Trim(arrVal2(3))
				frm1.txtDeptNm.value =Trim(arrVal2(4))
			Next	
			
		End If
	
		'----------------------------------------------------------------------------------------

End Sub


'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub



'========================================
Function BtnPreview_OnClick()
	Call BtnPrint("N")
End Function

'====================================================
Function BtnPrint_OnClick()
	Call BtnPrint("Y")
End Function

'=====================================================
Function BtnPrint(ByVal pvStrPrint) 

	If Not chkField(Document, "1") Then Exit Function

	Dim iStrUrl
	Dim varDept, varFrDt, varToDt, varDurYrs, varAcct, varAsstNo, varBizArea, varBizArea1, varBizUnit
	Dim IntRetCD

	varFrDt			= frm1.txtFr_dt.year & Right("0" & frm1.txtFr_dt.month,2)
	varToDt			= frm1.txtTo_dt.year & Right("0" & frm1.txtTo_dt.month,2)
	varDept			= Ucase(Trim(frm1.txtDeptCd.value))
	varAcct			= Ucase(Trim(frm1.txtAcctCd.value))
	varAsstNo		= Ucase(Trim(frm1.txtCondAsstNo.value))
	varBizArea		= Ucase(Trim(frm1.txtBizAreaCd.value))
	varBizArea1		= Ucase(Trim(frm1.txtBizAreaCd1.value))
	varBizUnit		= Ucase(Trim(frm1.txtBizUnitCd.value))
			
	if frm1.rdoDurYrsFg(0).checked then 
		varDurYrs = "C"
	else
		varDurYrs = "T"
	end if

	if varDept		= "" then		varDept		= "%"
	if varAcct		= "" then		varAcct		= "%"
	if varAsstNo	= "" then		varAsstNo	= "%"
	if varBizArea	= "" then		varBizArea	= "0"
	if varBizArea1	= "" then		varBizArea1	= "zzzzzzz"
	if varBizUnit	= "" then		varBizUnit	= "%"
       




	Dim strwhere
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6

	strwhere	=  " B.ASST_NO = A.ASST_NO "
	strwhere	= strwhere & " AND DEPR_YYYYMM between '" & varFrDt & "' and '" & varToDt & "' "
	strwhere	= strwhere & " AND DUR_YRS_FG = '" & varDurYrs & "' "

	if varDept <> "%" then
		strwhere	= strwhere & " AND B.DEPT_CD = '" & varDept & "' "
	end if
	if varAcct <> "%" then
		strwhere	= strwhere & " AND B.ACCT_CD ='" & varAcct & "' "
	end if
	if varAsstNo <> "%" then
		strwhere	= strwhere & " AND A.ASST_NO = '" & varAsstNo & "' "
	end if
	if varBizArea <> "0" then
		strwhere	= strwhere & " AND B.BIZ_AREA_CD >='" & varBizArea & "' "
	end if
	if varBizArea1 <> "zzzzzzz" then
		strwhere	= strwhere & " AND B.BIZ_AREA_CD <='" & varBizArea1 & "' "
	end if

	Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					

	' ���Ѱ��� �߰�
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND B.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		lgInternalCdAuthSQL		= " AND B.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		lgSubInternalCdAuthSQL	= " AND B.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		lgAuthUsrIDAuthSQL		= " AND B.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	strwhere	= strwhere	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

	Call CommonQueryRs("COUNT(*) ", "A_ASSET_DEPR_MASTER A(nolock), A_ASSET_MASTER B(nolock) ", strwhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	if (lgF0 = "") or Left(lgF0, Len(lgF0)-1) = "0" then 
		IntRetCD = DisplayMsgBox("800506","X","X","X")					'���ǿ� �ش��ϴ� �����Ͱ� �����ϴ�.(�ش絥���Ͱ� �ִ��� Ȯ���ϴ� ����)
		Exit Function
	end if



	Dim	strAuthCond
	' ���Ѱ��� �߰�
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	iStrUrl = "DurYrs|" & varDurYrs
	iStrUrl = iStrUrl & "|Dept|"		& varDept
	iStrUrl = iStrUrl & "|Acct|"		& varAcct
	iStrUrl = iStrUrl & "|AsstNo|"		& varAsstNo
	iStrUrl = iStrUrl & "|Bizarea|"		& varBizArea 
	iStrUrl = iStrUrl & "|Bizarea1|"	& varBizArea1
	iStrUrl = iStrUrl & "|BizUnit|"		& varBizUnit
	iStrUrl = iStrUrl & "|FrDt|"		& varFrDt
	iStrUrl = iStrUrl & "|ToDt|"		& varToDt
	
	iStrUrl = iStrUrl & "|strAuthCond|"	& strAuthCond

	OBjName = AskEBDocumentName("a7116oa1","ebr")    

	If pvStrPrint = "N" Then
		Call FncEBRPreview(ObjName, iStrUrl)
	Else
		Call FncEBRprint(EBAction, ObjName, iStrUrl)
	End If
		
End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						  <TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
			            				<TD CLASS="TD5" NOWRAP>�󰢳��</TD>
			            				<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtFr_dt name=txtFr_dt title=FPDATETIME tag="12X1" ALT="���ۿ�" CLASS=FPDTYYYYMM></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
           									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtTo_dt name=txtTo_dt title=FPDATETIME tag="12X1" ALT="�����" CLASS=FPDTYYYYMM></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>����������</TD>
								<TD CLASS="TD6" COLSPAN="3" NOWRAP> <SPAN STYLE="width:120;"><INPUT TYPE="RADIO" CLASS="RADIO" checked NAME="rdoDurYrsFg" TAG="12" ID="rbYrsTax"><LABEL TYPE="HIDDEN" FOR="rbYrsCas">���ȸ�����</LABEL></SPAN>
																		<SPAN STYLE="width:120;"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDurYrsFg" TAG="12" ID="rbYrsCas"><LABEL TYPE="HIDDEN" FOR="rbYrsTax">��������</LABEL></SPAN></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ڻ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCondAsstNo" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="�ڻ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMasterRef()"> <INPUT TYPE="Text" NAME="txtCondAsstNm" SIZE=25 MAXLENGTH=30 tag="14" ALT="�ڻ��"></TD>
								<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCd" SIZE=10 MAXLENGTH=15 tag="11XXXU" ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenAcctPopup(frm1.txtAcctCd.Value, '')"> <INPUT TYPE="Text" NAME="txtAcctNm" SIZE=25 MAXLENGTH=30 tag="14" ALT="������"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����μ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeptCd" SIZE=15 MAXLENGTH=15 tag="11XXXU" ALT="�μ��ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenAcctDeptPopUp(frm1.txtDeptCd.Value, '005')"> <INPUT TYPE="Text" NAME="txtDeptNm" SIZE=25 MAXLENGTH=30 tag="14" ALT="�μ���"></TD>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd.value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD> 
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizUnitCd" SIZE=15 MAXLENGTH=15 tag="11XXXU" ALT="������ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizUnitCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBizUnitPopup(frm1.txtBizUnitCd.Value, '')"> <INPUT TYPE="Text" NAME="txtBizUnitNm" SIZE=25 MAXLENGTH=30 tag="14" ALT="����θ�"></TD>
								<TD CLASS="TD5" NOWRAP></TD> 
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd1.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=25 tag="14"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR HEIGHT =100%>
								<TD WIDTH="100%" colspan=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							 </TR>
							 <TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
								<TD CLASS="TD5" NOWRAP>���󰢾�</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSum1 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:180px" title="FPDOUBLESINGLE" ALT="���󰢾�" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
							</TR>
						</TABLE>
					</TD>
			
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN"  Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN"  Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>	
		</TD>
	</TR>			
	<TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="htxtFr_dt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtTo_dt"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtDeptCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtAcctCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtCondAsstNo"	tag="24">
<INPUT TYPE=HIDDEN NAME="hDurYrsFg"	        tag="24">

<INPUT TYPE=HIDDEN NAME="htxtBizUnitCd"	    tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"		tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1"	tag="24">

<INPUT TYPE=TEXT NAME="hOrgChangeId"		tag="14" TABINDEX="-1">
<INPUT TYPE=TEXT NAME="hINternalCD"		tag="14" TABINDEX="-1">
				
				
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TabIndex="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TabIndex="-1">	
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

