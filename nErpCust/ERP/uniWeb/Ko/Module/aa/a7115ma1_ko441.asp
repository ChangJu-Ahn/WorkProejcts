
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset
'*  3. Program ID           : a7115ma1
'*  4. Program Name         : �����󰢰�������ȸ 
'*  5. Program Desc         : List Depreciation by Account
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000.11.20
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript"   SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript" SRC = "../../inc/incImage.js">				</SCRIPT>

<SCRIPT LANGUAGE= "VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Dim lgBlnFlgChgValue                                        '��: Variable is for Dirty flag            
'Dim lgStrPrevKey                                            '��: Next Key tag                          
'Dim lgSortKey                                               '��: Sort���� ���庯��                     
Dim IsOpenPop                                               '��: Popup status                          

'Dim lgSelectList                                            '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
'Dim lgSelectListDT                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 

'Dim lgTypeCD                                                '��: 'G' is for group , 'S' is for Sort    
'Dim lgFieldCD                                               '��: �ʵ� �ڵ尪                           
'Dim lgFieldNM                                               '��: �ʵ� ����                           
'Dim lgFieldLen                                              '��: �ʵ� ��(Spreadsheet����)              
'Dim lgFieldType                                             '��: �ʵ� ����                           
'Dim lgDefaultT                                              '��: �ʵ� �⺻��                           
'Dim lgNextSeq                                               '��: �ʵ� Pair��                           
'Dim lgKeyTag                                                '��: Key  ����                             

'Dim lgSortFieldNm                                           '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
'Dim lgSortFieldCD                                          '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark                                                  '��: ��ũ                                  
<%

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
'  Call GetAdoFiledInf("A7115MA1","G", "A")						'��: spread sheet �ʵ����� query   -----
																' G is for Qroup , S is for Sort
																' A is spreadsheet No
'--------------- ������ coding part(�������,End)-------------------------------------------------------------
%>

<!-- #Include file="../../inc/lgvariables.inc" -->

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "a7115mb1_ko441.asp"
Const C_MaxKey          = 0

Dim lsPoNo                                                 '��: Jump�� Cookie�� ���� Grid value
Dim  arrReturn
Dim  arrParent
Dim  arrParam

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

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
    lgIntFlgMode = parent.OPMD_CMODE									'Indicates that current mode is Create mode

	lgBlnFlgChgValue  = False											'��: Indicates that no value changed
    lgStrPrevKey      = ""												'��: initializes Previous Key
    lgSortKey         = 1												'��: initializes sort direction	
End Sub
'==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'=========================================================================================================
Sub SetDefaultVal()

'--------------- ������ coding part(�������,Start)--------------------------------------------------
 	
	frm1.txtDeprYYYYMM.text  = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)	
	Call ggoOper.FormatDate(frm1.txtDeprYYYYMM, parent.gDateFormat, 2)
	
'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>
End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("A7115MA1_KO441", "G", "A", "V20021211", parent.C_GROUP_DBAGENT ,  frm1.vspdData , C_MaxKey,"X","X")
	Call SetSpreadLock("A")
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(ByVal pOpt)
	If pOpt = "A" Then
		With frm1
			.vspdData.ReDraw = False
			ggoSpread.SpreadLock 1 , -1
			.vspdData.ReDraw = True
		End With
	End If
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()

End Sub

'**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 
Function OpenPopUp(strCode, iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	Case 0
		arrParam(0) = "�����ڵ��˾�"			' �˾� ��Ī 
		arrParam(1) = "A_ASSET_ACCT A, A_ACCT B"						' TABLE ��Ī 
		arrParam(2) = strCode						' Code Condition
		arrParam(3) = ""							' Name Cindition
		arrParam(4) = "A.ACCT_CD = B.ACCT_CD"			' Where Condition
		arrParam(5) = "�����ڵ�"				' �����ʵ��� �� ��Ī 
	
	    arrField(0) = "A.ACCT_CD"						' Field��(0)
		arrField(1) = "B.ACCT_NM"						' Field��(1)
    
	    arrHeader(0) = "�����ڵ�"				' Header��(0)
		arrHeader(1) = "������"					' Header��(1)
	Case Else
		Exit Function
	End Select
	
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1
			Select Case iWhere
			Case 0
				.txtAcctCd.value = arrRet(0)
				.txtAcctNm.value = arrRet(1)
			Case Else
			End Select
		End With
	End If	
	
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

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

	IsOpenPop = False

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
	End Select
	
	lgBlnFlgChgValue = True
End Function

'===========================================================================
' Function Name : OpenGroupPopup
' Function Desc : OpenGroupPopup Reference Popup
'===========================================================================
Function OpenGroupPopup()

End Function


Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

Sub OpenOrderByPopup(ByVal pSpdNo)

	Dim arrRet
	On Error Resume Next
	
	If IsOpenPop = True Then Exit sub
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp", Array(ggoSpread.GetXMLData("A"),gMethodText), "dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False

	If arrRet(0) = "X" Then
		Exit sub
	Else
		Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
		Call InitVariables
		Call InitSpreadSheet()
	End If

End Sub


'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================

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
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

'    ReDim lgPopUpR(parent.C_MaxSelList - 1,1)

	Call InitVariables														'��: Initializes local global variables
	frm1.txtDeprYYYYMM.focus

	Call SetDefaultVal	

	Call InitSpreadSheet()
'	Call InitComboBox()
'--------------- ������ coding part(�������,Start)----------------------------------------------------
'	Call CookiePage(0)
'    Call Parent.MASetToolbar("11000000000011")
    Call SetToolbar("11000000000011")
'	Call FncSetToolBar("New")
	frm1.txtDeprYYYYMM.focus

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
Sub txtDeprYYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub

Sub txtDeprYYYYMM_DblClick(Button)
    If Button = 1 Then
       frm1.txtDeprYYYYMM.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtDeprYYYYMM.Focus       
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col 'Sort in ascending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col ,lgSortKey 'Sort in descending
			lgSortKey = 1
		End If
		
		Exit Sub
	End If

	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
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

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgStrPrevKey <> "" Then								
			Call DisableToolBar(parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End If
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
	Dim IntRetCd

    FncQuery = False                                                        '��: Processing is NG

    Err.Clear                                                               '��: Protect system from crashing
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
		Exit Function
    End If
	
	If frm1.txtBizAreaCd.value = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If frm1.txtBizAreaCd1.value = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If UCase(Trim(frm1.txtBizAreaCd.value)) > UCase(Trim(frm1.txtBizAreaCd1.value)) Then
	  		IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
	  		frm1.txtBizAreaCd.focus
	  		Exit Function
		End If
	End If

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
     ggoSpread.source=frm1.vspdData
    Call ggoSpread.ClearSpreadData()

    Call InitVariables 														'��: Initializes local global variables
	
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
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

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
	Dim strYear
	Dim strMonth
	Dim strDay
	
    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1
		'--------------- ������ coding part(�������,Start)----------------------------------------------
 		Call ExtractDateFrom(frm1.txtDeprYYYYMM.Text,frm1.txtDeprYYYYMM.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
 		
		strVal = BIZ_PGM_ID & "?txtDeprYYYYMM="	& strYear & strMonth
		strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
		strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
		strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(frm1.txtBizAreaCd.alt)
		strVal = strVal & "&txtBizAreaCd1_Alt="	& Trim(frm1.txtBizAreaCd1.alt)
		   
		If frm1.Rb_WK1.checked = True Then
			strVal = strVal & "&txtDurYrsFg=" & "C"
		Else
			strVal = strVal & "&txtDurYrsFg=" & "T"
		End If
		'--------------- ������ coding part(�������,End)------------------------------------------------

        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")		'lgSelectListDT
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

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
'    Call @@@ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call FncSetToolBar("Query")

	frm1.txtDeprYYYYMM.focus
	Set gActiveElement = document.activeElement 
	'SetGridFocus
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>

				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=100%>
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
									<TD CLASS="TD5" NOWRAP>�󰢳��</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDeprYYYYMM" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT=�󰢳�� id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;</TD>
<!--									<TD CLASS="TD5" NOWRAP>
									<TD CLASS="TD6" NOWRAP>									
									<TD CLASS="HIDDEN"><INPUT TYPE="RADIO" CLASS="Radio" checked NAME="Radio1" TAG="12" ID="Rb_WK1" ><INPUT TYPE="RADIO" CLASS="Radio" NAME="Radio1" TAG="12" ID="Rb_WK2"></TD>			���� -->																															
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd.value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~</TD>
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>������ ����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 Checked tag="1"><LABEL FOR=Rb_WK1>���ȸ�����</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2 tag="1"><LABEL FOR=Rb_WK2>��������</LABEL></TD>
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
							<TR>
								<TD HEIGHT="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100%>
									<FIELDSET>
										<LEGEND>�հ�</LEGEND>
										<TABLE <%=LR_SPACE_TYPE_40%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>���⸻�󰢴����</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtAmtSum1 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="���⸻�󰢴����" tag="24X2"></OBJECT>');</SCRIPT></TD>
												<TD CLASS="TD5" NOWRAP>���⸻�̻��ܾ�</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtAmtSum5 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="���⸻�̻��ܾ�" tag="24X2"></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>����ʻ󰢴����</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtAmtSum2 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="����ʻ󰢴����" tag="24X2"></OBJECT>');</SCRIPT></TD>
												<TD CLASS="TD5" NOWRAP>����ʹ̻��ܾ�</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtAmtSum6 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="����ʹ̻��ܾ�" tag="24X2"></OBJECT>');</SCRIPT></TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>����󰢾�</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtAmtSum3 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="����󰢾�" tag="24X2"></OBJECT>');</SCRIPT></TD>
												<TD CLASS="TD5" NOWRAP> </TD>
												<TD CLASS="TD6" NOWRAP> </TD>
											</TR>
											<TR>
												<TD CLASS="TD5" NOWRAP>������󰢴����</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 name=txtAmtSum4 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="������󰢴����" tag="24X2"></OBJECT>');</SCRIPT></TD>
												<TD CLASS="TD5" NOWRAP>������̻��ܾ�</TD>
												<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 name=txtAmtSum7 style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH:150px" title="FPDOUBLESINGLE" ALT="������̻��ܾ�" tag="24X2"></OBJECT>');</SCRIPT></TD>
											</TR>
										</TABLE>
									</FIELDSET>
								</TD>
							<TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>


</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1" ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"     tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows"  tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hDeprYYYYMM" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hDurYrsFg"   tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hNextDataCount" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hNextDataFlag"  tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"  tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1" tag="24">
</FORM>
</BODY>
</HTML>


