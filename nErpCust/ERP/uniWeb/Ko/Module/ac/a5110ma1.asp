
<%Response.Expires = -1%>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Closing and Financial Statements
'*  3. Program ID           : a5110ma1
'*  4. Program Name         : ��(��)��ǥ ��ȸ 
'*  5. Program Desc         : Query of Daily/Monthly Summerization
'*  6. Comproxy List        : AG00411
'*  7. Modified date(First) : 1999/09/10
'*  8. Modified date(Last)  : 2001/02/14
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
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
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs">				  </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim lgBlnFlgChgValue                                        '��: Variable is for Dirty flag            
Dim lgStrPrevKey                                            '��: Next Key tag                          
Dim lgSortKey                                               '��: Sort���� ���庯��                      
Dim IsOpenPop                                               '��: Popup status                           
Dim lgIsOpenPop     

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

'Dim lgPopUpR                                                '��: Orderby default ��                    
Dim lgMark                                                  '��: ��ũ                                  


' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
'  Call GetAdoFiledInf("A5110MA1","S", "A")						'��: spread sheet �ʵ����� query   -----
																' G is for Qroup , S is for Sort
																' A is spreadsheet No
'--------------- ������ coding part(�������,End)-------------------------------------------------------------


'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "A5110MB1.asp"
Const BIZ_PGM_ID_SP 	= "a5110mb2.asp"

Const C_SHEETMAXROWS    = 30                                   '��: Spread sheet���� �������� row
Const C_SHEETMAXROWS_D  = 1000                                 '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
'Dim lsPoNo								                       '��: Jump�� Cookie�� ���� Grid value
Const C_MaxKey          = 0                                    '�١١١�: Max key value
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
	Dim strSvrDate, strDayCnt

	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	EndDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(EndDate, Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)


	frm1.txtDateFr.Text = StartDate 
	frm1.txtDateTo.Text = EndDate 

	
'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "COOKIE", "QA") %>
End Sub


'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("A5110MA1","S","A","V20021220",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock    
End Sub


'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
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
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	Select Case iWhere
	Case 0
		arrParam(0) = "����� �˾�"						' �˾� ��Ī 
		arrParam(1) = "B_Biz_AREA"							' TABLE ��Ī 
		arrParam(2) = strCode								' Code Condition
		arrParam(3) = ""									' Name Cindition

		' ���Ѱ��� �߰� 
		If lgAuthBizAreaCd <>  "" Then
			arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
		Else
			arrParam(4) = ""
		End If

		arrParam(5) = "������ڵ�"			
	
	    arrField(0) = "BIZ_AREA_CD"								' Field��(0)
		arrField(1) = "BIZ_AREA_NM"								' Field��(1)
    
	    arrHeader(0) = "������ڵ�"							' Header��(0)
		arrHeader(1) = "������"							' Header��(1)
    
	Case 1
		arrParam(0) = "�ϰ�ǥ���� �˾�"					' �˾� ��Ī 
		arrParam(1) = "A_ACCT_CLASS_TYPE"						' TABLE ��Ī 
		arrParam(2) = strCode									' Code Condition
		arrParam(3) = ""										' Name Cindition
		arrParam(4) = "CLASS_TYPE LIKE " & FilterVar("DMS%", "''", "S") & " "										' Where Condition
		arrParam(5) = "�ϰ�ǥ����"			
	
	    arrField(0) = "CLASS_TYPE"								' Field��(0)
		arrField(1) = "CLASS_TYPE_NM"							' Field��(1)
    
	    arrHeader(0) = "�ϰ�ǥ����"						' Header��(0)
		arrHeader(1) = "�ϰ�ǥ������"							' Header��(1)
    
	Case Else
		Exit Function
	End Select
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	With frm1
		Select Case iWhere
		Case 0
			.txtBizAreaCd.focus
		Case 1
			.txtClassType.focus
		End Select
	End With
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet, iWhere)
	End If	

End Function


 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
 '------------------------------------------  SetItemInfo()  --------------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
		Case 0
			.txtBizAreaCd.value = arrRet(0)
			.txtBizAreaNm.value = arrRet(1)
		Case 1
			.txtClassType.value   = arrRet(0)
			.txtClassTypeNm.value = arrRet(1)
		End Select
	End With

End Function

'===========================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "2")	   
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

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						 'Cookie Split String : CookiePage Function Use

	If Kubun = 1 Then								 'Jump�� ȭ���� �̵��� ��� 

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

		WriteCookie "PoNo" , lsPoNo					 'Jump�� ȭ���� �̵��Ҷ� �ʿ��� Cookie �������� 
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							 'Jump�� ȭ���� �̵��� ������� 

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		Dim iniSep

'--------------- ������ coding part(�������,Start)---------------------------------------------------
		 '�ڵ���ȸ�Ǵ� ���ǰ��� �˻����Ǻ� Name�� Match 
		For iniSep = 0 To UBound(arrVal) -1
			Select Case UCase(Trim(arrVal(iniSep)))
			Case UCase("��������")
				frm1.txtPoType.value =  arrVal(iniSep + 1)
			Case UCase("�������¸�")
				frm1.txtPoTypeNm.value =  arrVal(iniSep + 1)
			Case UCase("����ó")
				frm1.txtSpplCd.value =  arrVal(iniSep + 1)
			Case UCase("����ó��")
				frm1.txtSpplNm.value =  arrVal(iniSep + 1)
			Case UCase("���ű׷�")
				frm1.txtPurGrpCd.value =  arrVal(iniSep + 1)
			Case UCase("���ű׷��")
				frm1.txtPurGrpNm.value =  arrVal(iniSep + 1)
			Case UCase("ǰ��")
				frm1.txtItemCd.value =  arrVal(iniSep + 1)
			Case UCase("ǰ���")
				frm1.txtItemNm.value =  arrVal(iniSep + 1)
			Case UCase("Tracking No.")
				frm1.txtTrackNo.value =  arrVal(iniSep + 1)
			End Select
		Next
'--------------- ������ coding part(�������,End)---------------------------------------------------

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call FncQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

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

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

'    ReDim lgPopUpR(Parent.C_MaxSelList - 1,1)
 
	Call InitVariables													'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call InitComboBox()
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Call FncSetToolBar("New")
'	Call CookiePage(0)
    
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

    frm1.txtDateFr.focus
    frm1.txtYAmt.allownull = False 
    frm1.txtTAmt.allownull = False 
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

Sub txtDateFr_DblClick(Button)
	if Button = 1 then
		frm1.txtDateFr.Action = 7
	End if
End Sub

Sub txtDateTo_DblClick(Button)
	if Button = 1 then
		frm1.txtDateTo.Action = 7
	End if
End Sub

Sub txtDateFr_Keypress(Key)
    If Key = 13 Then
		frm1.txtDateTo.focus
        FncQuery()
    End If
End Sub

Sub txtDateTo_Keypress(Key)
    If Key = 13 Then
		frm1.txtDateFr.focus
        FncQuery()
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
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

	frm1.vspdData.Row = Row
'	lsPoNo=frm1.vspdData.Text
'--------------- ������ coding part(�������,End)------------------------------------------------------
	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row) 
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
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


'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
     '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	'��: ������ üũ'
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

   
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then										'��: This function check indispensable field
       Exit Function
    End If
	
	If CompareDateByFormat(frm1.txtDateFr.Text, frm1.txtDateTo.Text, frm1.txtDateFr.Alt, frm1.txtDateTo.Alt, _
						"970025", frm1.txtDateFr.UserDefinedFormat, Parent.gComDateType, true) = False Then
			frm1.txtDateFr.focus											'��: GL Date Compare Common Function
			Exit Function
	End if  
	
	 '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData


    '-----------------------
    'Query function call area
    '-----------------------
    IF  DbQuery	= False Then														'��: Query db data
		Exit Function
	END IF
	
    FncQuery = True		
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call Parent.FncExport(Parent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call Parent.FncFind(Parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    frm1.txtOUT.value = "2"
	Call DbQuery2
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
	Dim strValSp, strZeroFg
	
    Err.Clear                                                       
    DbQuery = False
    
'    Call GetQueryDate()
	Call LayerShowHide(1)

	if frm1.ZeroFg1.checked = True Then
		strZeroFg = "Y"
	Else
		strZeroFg = "N"
	End IF
	With frm1

			'sp�� ȣ���Ѵ�.        				
			strValSp = BIZ_PGM_ID_SP & "?txtStartDt="     & Trim(.txtDateFr.Text)
			strValSp = strValSp & "&txtEndDt="       & Trim(.txtDateTo.Text)
        	strValSp = strValSp & "&txtClassType=" & Trim(.txtClassType.value)
        	strValSp = strValSp & "&txtBizArea="	& Trim(.txtBizAreaCd.value)
			strValSp = strValSp & "&strZeroFg="		& strZeroFg
        	strValSp = strValSp & "&strUserId="		& Parent.gUsrID
        	strValSp = strValSp & "&strSpid="		& Trim(.txtSpid.value)

			' ���Ѱ��� �߰� 
			strValSp = strValSp & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
			strValSp = strValSp & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
			strValSp = strValSp & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
			strValSp = strValSp & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

        	Call RunMyBizASP(MyBizASP, strValSp)
           
    End With
    
    DbQuery = True

End Function


Function DbQuery2() 
	Dim strVal

    IF frm1.txtOUT.value = "" THEN 
       frm1.txtOUT.value = "1"
       DbQuery2 = False
    
       Err.Clear                                                               '��: Protect system from crashing
	   Call LayerShowHide(1)
	   Call FncSetToolBar("Query")
	 END IF

		
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------
		strVal = BIZ_PGM_ID & "?txtDateFr=" & Trim(.txtDateFr.Text)
		strVal = strVal & "&txtDateTo=" & Trim(.txtDateTo.Text)
		strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.Value)
		strVal = strVal & "&txtClassType=" & Trim(.txtClassType.Value)		
		strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtClassType_Alt=" & Trim(.txtClassType.Alt)		
		strVal = strVal & "&txtSPID=" & Trim(.txtSPID.value)		
		strVal = strVal & "&txtOUT=" & Trim(.txtOUT.value)		
		
'--------------- ������ coding part(�������,End)------------------------------------------------

		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag
        strVal = strVal & "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D)            '��: �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSqlGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ����       
              	
    End With
    
    DbQuery2 = True


End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()	
	Call DbQuery2()
End Function

Function DbQuery2Ok()														'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
'    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

	IF Trim(frm1.txtBizAreaCd.value) = "" then
		frm1.txtBizAreaNm.value = ""
	end if	

	'SetGridFocus
		
	'frm1.txtBankCd.focus
		Call FncSetToolBar("New")

	Set gActiveElement = document.activeElement 
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################


'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Function SetPrintCond(StrEbrFile, VarBizArea, VarDateFr, VarDateTo, VarBalYAmt, VarBalTAmt, VarSpid)
	StrEbrFile = "a5110ma1"
	
	With frm1

		' ���Ѱ��� �߰� 
		Dim IntRetCD
	
		varBizArea = UCASE(Trim(.txtBizAreaCd.value))

		If varBizArea = "" Then
			If lgAuthBizAreaCd <> "" Then			
				varBizArea  = lgAuthBizAreaCd
			Else
				varBizArea = "*"
			End If			
		Else
			If lgAuthBizAreaCd <> "" Then			
				If UCASE(lgAuthBizAreaCd) <> varBizArea Then
					IntRetCD = DisplayMsgBox("124200","x","x","x")
					SetPrintCond =  False
					Exit Function
				End If			
			End If			
		End If

		VarDateFr	= UniConvDateToYYYYMMDD(frm1.txtDateFr.Text,Parent.gDateFormat,"")
		VarDateTo	= UniConvDateToYYYYMMDD(frm1.txtDateTo.Text,Parent.gDateFormat,"")
		VarBalYAmt = .txtYAmt.text
		VarBalTAmt = .txtTAmt.text
		VarSpid = UCase(Trim(.txtSpid.value))
	End With

	SetPrintCond =  True
	
End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile, VarBizArea, VarClassTypeFr, VarClassTypeTo, VarDateFr, VarDateTo, VarBalYAmt, VarBalTAmt, VarSpid
	Dim strGlDtYr, strGlDtMnth, strGlDtDt
	Dim Fiscyyyy,Fiscmm,Fiscdd,VarFiscDt
	Dim IntRetCD	

    On Error Resume Next                                                    '��: Protect system from crashing
    
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	If UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat, "") Then
		Call DisplayMsgBox("970025", "X", frm1.txtDateFr.Alt, frm1.txtDateTo.Alt)
		frm1.txtDateFr.focus
		Exit Function
	End If

	IntRetCD =  SetPrintCond(StrEbrFile, VarBizArea, VarDateFr, VarDateTo, VarBalYAmt, VarBalTAmt, VarSpid)
	If IntRetCD = False Then
	    Exit Function
 	End If

	ObjName = AskEBDocumentName(StrEbrFile, "ebr")
    
    lngPos = 0
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

    Call ExtractDateFrom(UNIConvDate(frm1.txtDateFr.text),Parent.gServerDateFormat,Parent.gServerDateType,strGlDtYr,strGlDtMnth,strGlDtDt)	
    Call ExtractDateFrom(UNIConvDate(frm1.txtDateTo.text),Parent.gServerDateFormat,Parent.gServerDateType,strGlDtYr,strGlDtMnth,strGlDtDt)	
    
'============================
	StrUrl = StrUrl & "BizArea|"	& VarBizArea
	StrUrl = StrUrl & "|DateFr|"	& VarDateFr
	StrUrl = StrUrl & "|DateTo|"	& VarDateTo
	StrUrl = StrUrl & "|BalYAmt|"	& VarBalYAmt
	StrUrl = StrUrl & "|BalTAmt|"	& VarBalTAmt
	StrUrl = StrUrl & "|Spid|"	& VarSpid

	Call FncEBRPrint(EBAction,ObjName,StrUrl)
		
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	On Error Resume Next                                                    '��: Protect system from crashing
    
	Dim StrUrl
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile, VarBizArea, VarClassTypeFr, VarClassTypeTo, VarDateFr, VarDateTo, VarBalYAmt, VarBalTAmt, VarSpid
	Dim strGlDtYr, strGlDtMnth, strGlDtDt
	Dim Fiscyyyy,Fiscmm,Fiscdd,VarFiscDt
	Dim IntRetCD
    
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	
	If UniConvDateToYYYYMMDD(frm1.txtDateFr.Text, Parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtDateTo.Text, Parent.gDateFormat, "") Then
		Call DisplayMsgBox("970025", "X", frm1.txtDateFr.Alt, frm1.txtDateTo.Alt)
		frm1.txtDateFr.focus
		Exit Function
	End If

	IntRetCD =  SetPrintCond(StrEbrFile, VarBizArea, VarDateFr, VarDateTo, VarBalYAmt, VarBalTAmt, VarSpid)
	If IntRetCD = False Then
	    Exit Function
 	End If

	ObjName = AskEBDocumentName(StrEbrFile, "ebr")

    Call ExtractDateFrom(UNIConvDate(frm1.txtDateFr.text),Parent.gServerDateFormat,Parent.gServerDateType,strGlDtYr,strGlDtMnth,strGlDtDt)	
    Call ExtractDateFrom(UNIConvDate(frm1.txtDateTo.text),Parent.gServerDateFormat,Parent.gServerDateType,strGlDtYr,strGlDtMnth,strGlDtDt)	

'============================
	StrUrl = StrUrl & "BizArea|"	& VarBizArea
	StrUrl = StrUrl & "|DateFr|"	& VarDateFr
	StrUrl = StrUrl & "|DateTo|"	& VarDateTo
	StrUrl = StrUrl & "|BalYAmt|"	& VarBalYAmt
	StrUrl = StrUrl & "|BalTAmt|"	& VarBalTAmt
	StrUrl = StrUrl & "|Spid|"	& " " & FilterVar(VarSpid, "''", "S") & ""

	Call FncEBRPreview(ObjName,StrUrl)
		
End Function


'==========================================================
'���ٹ�ư ���� 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolBar("1100000000001111")
	Case "QUERY"
		Call SetToolBar("1000000000011111")
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>ȸ����</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateFr name=txtDateFr CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="��������"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTo name=txtDateTo CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12X1" ALT="��������"></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>������ڵ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtBizAreaCd.value,0)">&nbsp;
														   <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 tag="24X" ALT="������" STYLE="TEXT-ALIGN: Left">
									</TD>										  
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ϰ�ǥ����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtClassType" SIZE=11 MAXLENGTH=4 tag="12XXXU" ALT="�ϰ�ǥ����" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnClassType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(txtClassType.value,1)">&nbsp;
														   <INPUT TYPE=TEXT NAME="txtClassTypeNm" SIZE=20 tag="24X" ALT="�ϰ�ǥ������" STYLE="TEXT-ALIGN: Left">
									</TD>
									<TD CLASS="TD5" NOWRAP>��ȸ����</TD>
									<TD CLASS="TD6" NOWRAP>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ZeroFg" ID="ZeroFg1" VALUE="Y" tag="15"><LABEL FOR="ZeroFg1">��ü</LABEL></SPAN>
										<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="ZeroFg" CHECKED ID="ZeroFg2" VALUE="N" tag="15"><LABEL FOR="ZeroFg2">�߻��ݾ�</LABEL></SPAN></TD>
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
								<TD HEIGHT="100%" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>���������ܾ�</TD>
								<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtYAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="���������ܾ�" tag="24X2" id=fpDoubleSingle1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>���������ܾ�</TD>
								<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="���������ܾ�" tag="24X2" id=fpDoubleSingle4></OBJECT>');</SCRIPT></TD>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>&nbsp;
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBizAreaCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hClassType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hClassCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDateFr" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hDateTo" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hCommand" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSPID" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtOUT" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

