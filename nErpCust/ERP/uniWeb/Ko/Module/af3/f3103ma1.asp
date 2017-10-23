
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : f3103ma1
'*  4. Program Name         : �������ܰ���ȸ 
'*  5. Program Desc         : Query of Deposit Balance
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.01.11
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song, Mun Gil  ������ڵ�, �����ڵ� ���� Check
'======================================================================================================= -->
<HTML> 
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->					<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->		
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	
'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Dim lgBlnFlgChgValue                                        '��: Variable is for Dirty flag            
'Dim lgStrPrevKey                                            '��: Next Key tag                          
'Dim lgSortKey                                               '��: Sort���� ���庯��                      
Dim IsOpenPop                                               '��: Popup status                           

Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark    
Dim lgIsOpenPop                                             '��: ��ũ                                  

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

CONST C_MaxKey = 4 
<%

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
  'Call GetAdoFiledInf("F3103MA1","G", "A")						'��: spread sheet �ʵ����� query   -----
																' G is for Qroup , S is for Sort
																' A is spreadsheet No
'--------------- ������ coding part(�������,End)-------------------------------------------------------------
%>

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "f3103mb1.asp"

'Dim lsPoNo                                                 '��: Jump�� Cookie�� ���� Grid value
'--------------- ������ coding part(��������,End)-------------------------------------------------------------

 '#########################################################################################################
'												2. Function�� 
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

 '==========================================  2.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgStrPrevKey     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()
	
	
'--------------- ������ coding part(�������,Start)--------------------------------------------------

	frm1.txtDateMid.Text = UniConvDateAToB("<%=GetSvrDate%>" ,Parent.gServerDateFormat,Parent.gDateFormat) 

	'frm1.txtBizAreaCd.value	= Parent.gBizArea	
'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "F","NOCOOKIE","QA") %>	
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","MA") %>

	
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
        Call SetZAdoSpreadSheet("f3103ma1","G","A","V20021215",Parent.C_GROUP_DBAGENT, frm1.vspdData, C_MaxKey,"X","X")
 	    Call SetSpreadLock("A") 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(Byval iOpt)
   If iOpt = "A" Then
		With frm1
			.vspdData.ReDraw = False
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
		End With
    End If
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3012", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDpstType ,lgF0  ,lgF1  ,Chr(11))
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F3014", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboTransSts ,lgF0  ,lgF1  ,Chr(11))
End Sub

'**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	frm1.hOrgChangeId.value = Parent.gChangeOrgId

	Select Case iWhere
		Case 0, 3
			arrParam(0) = "������ڵ� �˾�"								' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA" 										' TABLE ��Ī 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Cindition

			' ���Ѱ��� �߰� 
			If lgAuthBizAreaCd <> "" Then
				arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
			Else
				arrParam(4) = ""
			End If

			arrParam(5) = "������ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"										' Field��(0)
			arrField(1) = "BIZ_AREA_NM"										' Field��(1)
    
			arrHeader(0) = "������ڵ�"									' Header��(0)
			arrHeader(1) = "������"									' Header��(1)
			
		Case 1
			arrParam(0) = "�����ڵ� �˾�"								' �˾� ��Ī 
			arrParam(1) = "B_BANK B"	'" B_BANK B, B_BANK_ACCT A"		' TABLE ��Ī 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = ""	'"B.BANK_CD = A.BANK_CD"							' Where Condition
			arrParam(5) = "�����ڵ�"									' �����ʵ��� �� ��Ī 
	
			arrField(0) = "B.BANK_CD"						' Field��(0)
			arrField(1) = "B.BANK_NM"						' Field��(1)
'			arrField(1) = "A.BANK_ACCT_NO"						' Field��(1)
    
			arrHeader(0) = "�����ڵ�"					' Header��(0)
			arrHeader(1) = "�����"						' Header��(1)
'			arrHeader(1) = "���¹�ȣ"						' Header��(1)
		
		Case 2
			arrParam(0) = "��ȭ�ڵ� �˾�"				' �˾� ��Ī 
			arrParam(1) = " B_CURRENCY A"					' TABLE ��Ī 
			arrParam(2) = Trim(strCode)						' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "��ȭ�ڵ�"					' �����ʵ��� �� ��Ī 
	
			arrField(0) = "A.CURRENCY"						' Field��(0)
			arrField(1) = "A.CURRENCY_DESC"					' Field��(1)
    
			arrHeader(0) = "��ȭ�ڵ�"					' Header��(0)
			arrHeader(1) = "��ȭ��"						' Header��(1)
		
		Case Else
			Exit Function
	End Select
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Select Case iWhere
		Case 0
			frm1.txtBizAreaCd.value = arrRet(0)
			frm1.txtBizAreaNm.value = arrRet(1)
			
			frm1.txtBizAreaCd.focus
		Case 1
			frm1.txtBankCd.value = arrRet(0)
			frm1.txtBankNm.value = arrRet(1)
			
			frm1.txtBankCd.focus
		Case 2
			frm1.txtDocCur.value = arrRet(0)
			frm1.txtDocCur.focus
		Case 3
			frm1.txtBizAreaCd1.value = arrRet(0)
			frm1.txtBizAreaNm1.value = arrRet(1)
		End Select
	End If	

End Function

'===========================================================================
' Function Name : OpenGroupPopup
' Function Desc : OpenGroupPopup Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
 '******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)	
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call InitComboBox()
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Call FncSetToolBar("New")
	frm1.txtDateMid.focus

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

'--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 


'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'==========================================================================================

Sub txtDateMid_DblClick(Button)
	if Button = 1 then
		frm1.txtDateMid.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtDateMid.Focus
	End if
End Sub

Sub txtDateMid_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData
        
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
'	lsPoNo=frm1.vspdData.Text
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then							'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DbQuery
		End If
   End if
    
End Sub


 '#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


 '#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggospread.ClearSpreadData									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If

	Call FncSetToolBar("New")    
    '-----------------------
    'Query function call area
    '-----------------------
		
    Call DbQuery															'��: Query db data

    FncQuery = True		
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
	Dim ACol
	Dim ARow
	Dim iRet
	Dim iColumnLimit
	
	iColumnLimit = frm1.vspdData.MaxCols
	
	ACol = frm1.vspdData.ActiveCol
	ARow = frm1.vspdData.ActiveRow
	
	If ACol > iColumnLimit Then
		iRet = DisplayMsgBox("900030", "X", iColumnLimit, "X")
		Exit Function
	End If
	
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit(ACol)
	
	frm1.vspdData.Col = ACol
	frm1.vspdData.Row = ARow
	frm1.vspdData.Action = 0
	frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtBizAreaCd="	& Trim(.txtBizAreaCd.value)
		strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
		strVal = strVal & "&txtBankCd="			& Trim(.txtBankCd.value)
		strVal = strVal & "&cboDpstType="		& Trim(.cboDpstType.value)
		strVal = strVal & "&txtDateMid="		& Trim(.txtDateMid.Text)
		strVal = strVal & "&cboTransSts="		& Trim(.cboTransSts.value)
		strVal = strVal & "&txtDocCur="			& Trim(.txtDocCur.value)
		strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtBizAreaCd_Alt1="	& Trim(.txtBizAreaCd1.Alt)
		strVal = strVal & "&txtBankCd_Alt="		& Trim(.txtBankCd.Alt)
		strVal = strVal & "&txtDocCur_Alt="		& Trim(.txtDocCur.Alt)
'--------------- ������ coding part(�������,End)------------------------------------------------

        strVal = strVal & "&lgStrPrevKey="		& lgStrPrevKey                      '��: Next key tag
        strVal = strVal & "&lgPageNo="			& lgPageNo                      '��: Next key tag		
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))
		
		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True


End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call FncSetToolBar("Query")
	Call CurFormatNumericOCX()
	
	'SetGridFocus
	Set gActiveElement = document.activeElement 
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	Dim intRetCD
	Dim strGBCurrency
	
	With frm1

		If Trim(.txtDocCur.value) = "" Then
		
            intRetCD = CommonQueryRs("top 1 currency"," b_numeric_format "," decimals  = (select max(decimals) from b_numeric_format where data_type = 2 ) and data_type=2 and form_type = " & FilterVar("Q", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            
		  If intRetCD = True Then	
		  	strGBCurrency = Trim(Replace(lgF0,Chr(11),""))
		  Else
		    strGBCurrency = 	parent.gCurrency
		  End If					
			ggoOper.FormatFieldByObjectOfCur .txtRcptAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBalAmt,	strGBCurrency, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		  
		Else
			ggoOper.FormatFieldByObjectOfCur .txtRcptAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtPaymAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
			ggoOper.FormatFieldByObjectOfCur .txtBalAmt,	.txtDocCur.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
		End If
	End With

End Sub



'==========================================================
'���ٹ�ư ���� 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDateMid" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=�������� id=fpDateMid></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>����������</TD>
								<TD CLASS="TD6" NOWRAP><SELECT ID="cboDpstType" NAME="cboDpstType" ALT="�����ݱ���" STYLE="WIDTH: 132px" tag="11X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBankCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�����ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBankCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBankCD.Value,1)">&nbsp;<INPUT TYPE="Text" NAME="txtBankNm" SIZE=25 tag="24X" ALT="�����"></TD>
								<TD CLASS="TD5" NOWRAP>�ŷ�����</TD>
								<TD CLASS="TD6" NOWRAP><SELECT ID="cboTransSts" NAME="cboTransSts" ALT="�ŷ�����" STYLE="WIDTH: 132px" tag="11X"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD.Value,0)">&nbsp;
													   <INPUT TYPE="Text" NAME="txtBizAreaNm" SIZE=25 tag="24X" ALT="������">&nbsp;~&nbsp;</TD>
								<TD CLASS="TD5" NOWRAP>��ȭ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="��ȭ" SIZE = "10" MAXLENGTH="3"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCurCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtDocCur.value, 2)"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCD1.Value,3)">&nbsp;
													   <INPUT TYPE="Text" NAME="txtBizAreaNm1" SIZE=25 tag="24X" ALT="������"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=* WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100%>
									<TABLE <%=LR_SPACE_TYPE_60%>>
										<TR>
											<TD CLASS="TD5" NOWRAP>�Ա��հ�</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtRcptAmt" title=FPDOUBLESINGLE ALT="�Ա��հ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
											<TD CLASS="TD5" NOWRAP>�Ա��հ�(�ڱ�)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtRcptLocAmt" title=FPDOUBLESINGLE ALT="�Ա��հ�(�ڱ�)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>																						
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>����հ�</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPaymAmt" title=FPDOUBLESINGLE ALT="����հ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>											
											<TD CLASS="TD5" NOWRAP>����հ�(�ڱ�)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtPaymLocAmt" title=FPDOUBLESINGLE ALT="����հ�(�ڱ�)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>																						
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>�ܾ�</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtBalAmt" title=FPDOUBLESINGLE ALT="�ܾ�" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>																																	
											<TD CLASS="TD5" NOWRAP>�ܾ�(�ڱ�)</TD>
											<TD class=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> CLASS=FPDS140 name="txtBalLocAmt" title=FPDOUBLESINGLE ALT="�ܾ�(�ڱ�)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>																																	
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></td>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

