
<%@ LANGUAGE="VBSCRIPT" %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : PrePayment management
'*  3. Program ID           : a3115ma1.asp
'*  4. Program Name         : ä�ǻ���ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002.12.18
'*  8. Modified date(Last)  : 2004/02/04
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : U&I(Kim Chang Jin)
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2001.01.13
							  2004/02/04	��������� �߰� 
'**********************************************************************************************
 -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************ -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================ -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'=========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>
<Script Language="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID        = "a3115MB1.asp"                         '��: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "a3115MB2.asp"                         '��: Biz logic spread sheet for #2
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================


Const C_MaxKey            = 5                                    '�١١١�: Max key value
Const C_MaxKey_B            = 3                                   '�١١١�: Max key value
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  lgIsOpenPop                                            '��: Popup status                           
Dim  lgKeyPosVal
Dim  IsOpenPop												'��: Popup status   
Dim  lgPageNo_A                                              '��: Next Key tag                          
Dim  lgSortKey_A                                             '��: Sort���� ���庯��                     
Dim  lgPageNo_B                                              '��: Next Key tag                          
Dim  lgSortKey_B                                             '��: Sort���� ���庯��                     

ReDim  lgKeyPosVal(C_MaxKey)
Dim strYear, strMonth, strDay,  EndDate, StartDate


' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


<%
	Dim dtToday
	dtToday = GetSvrDate
%>	
	Call ExtractDateFrom("<%=dtToday%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateAdd("M", -1, EndDate, parent.gDateFormat)


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
Sub  InitVariables()
    lgBlnFlgChgValue = False                               'Indicates that no value changed
    lgIntFlgMode     = parent.OPMD_CMODE                          'Indicates that current mode is Create mode

    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B		 = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1
End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub  SetDefaultVal()
	frm1.txtFromDt.text	= StartDate
	frm1.txtToDt.text	= EndDate
End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub  LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>
End Sub
'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("A3115MA01","S","A","V20021211",Parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetZAdoSpreadSheet("A3115MA02","S","B","V20021211",Parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey_B, "X","X")
	Call SetSpreadLock("A")
	Call SetSpreadLock("B")																		
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadLock(Byval iOpt )
	If iOpt = "A" Then                                   ' �ʱ�ȭ Spreadsheet #1 
		With frm1.vspdData
			.ReDraw = False
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadLockWithOddEvenRowColor()	
			.ReDraw = True
		End With 
    Else                                                ' �ʱ�ȭ Spreadsheet #2 
		With frm1.vspdData2
			.ReDraw = False       
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SpreadLockWithOddEvenRowColor()	
			.ReDraw = True
		End With 
    End If   
End Sub

 '**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""									'FrDt
	arrParam(3) = ""								'ToDt 
	arrParam(4) = "B"							'B :���� S: ���� T: ��ü 
	Select Case iWhere
		Case 1
			arrParam(5) = "SOL"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
		Case 2
			arrParam(5) = "PAYER"									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	End Select
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.txtDealBpCd.focus
			Case 2
				frm1.txtPayBpCd.focus
		End Select
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)

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


'------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenConRouting()
'	Description : Routing PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcctPopUp(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim Field_fg

	If lgIsOpenPop = True Then Exit Function	
	
	lgIsOpenPop = True
	
	Field_fg = 3
	
	arrParam(0) = "�����ڵ��˾�"								' �˾� ��Ī 
	arrParam(1) = "A_Acct, A_ACCT_GP" 											' TABLE ��Ī 
	arrParam(2) = Trim(strCode)											' Code Condition
	arrParam(3) = ""												' Name Cindition
	arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD"					' Where Condition
	arrParam(5) = "�����ڵ�"									' �����ʵ��� �� ��Ī 

	arrField(0) = "A_ACCT.Acct_CD"									' Field��(0)
	arrField(1) = "A_ACCT.Acct_NM"									' Field��(1)
	arrField(2) = "A_ACCT_GP.GP_CD"									' Field��(2)
	arrField(3) = "A_ACCT_GP.GP_NM"									' Field��(3)
			
	arrHeader(0) = "�����ڵ�"									' Header��(0)
	arrHeader(1) = "�����ڵ��"									' Header��(1)
	arrHeader(2) = "�׷��ڵ�"									' Header��(2)
	arrHeader(3) = "�׷��"										' Header��(3)

	lgIsOpenPop = True
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then	    
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,Field_fg)
	End If
End Function

Function OpenARPopUp()
	Dim arrRet
	Dim Field_fg
	Dim arrParam
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a3101ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3101ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	ReDim arrParam(8)

	If lgIsOpenPop = True Then Exit Function	
	
	lgIsOpenPop = True
	
	Field_fg = 4
			
	' ���Ѱ��� �߰� 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
	     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			

	lgIsOpenPop = False
	
	If arrRet(0) = "" Then	    
		frm1.txtArNo.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,Field_fg)
	End If	 
End Function			
			
'=======================================================================================================
'	Name : SetBankAcct()
'	Description : Bank Account No Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		Case 1
			frm1.txtDealBpCd.value = arrRet(0)
			frm1.txtDealBpNm.value = arrRet(1)
			frm1.txtDealBpCd.focus				
		Case 2
			frm1.txtPayBpCd.value = arrRet(0)
			frm1.txtPayBpNm.value = arrRet(1)				
			frm1.txtPayBpCd.focus
		Case 3
			frm1.txtAcctCd.value = arrRet(0)
			frm1.txtAcctNm.value = arrRet(1)
			frm1.txtArNo.focus
		Case 4
			frm1.txtArNo.value = arrRet(0)
			frm1.txtArNo.focus
		case 5
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 6
			frm1.txtBizAreaCd1.Value	= arrRet(0)
			frm1.txtBizAreaNm1.Value	= arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : 
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	Dim iGridPos
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			iGridPos = "A"
		Case "VSPDDATA2"			
			iGridPos = "B"
	End Select			
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(iGridPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(iGridPos,arrRet(0),arrRet(1))
       Call InitVariables()
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
Sub  Form_Load()
	Call LoadInfTB19029()			
	
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec) 
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.LockField(Document, "N")										'��: Lock  Suitable  Field
    
	Call InitVariables()														'��: Initializes local global variables
	Call SetDefaultVal()
	Call InitSpreadSheet()
    Call SetToolbar("1100000000000111")											'��: ��ư ���� ���� 

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
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
Sub txtFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then
		frm1.txtToDt.focus
		Call FncQuery
	ENd if
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromDt.focus
		Call FncQuery
	End if		
End Sub

'========================================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'=========================================================================================================
Sub  txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus 	
	End If
End Sub

'========================================================================================================
'   Event Name : txtPoToDt
'   Event Desc :
'========================================================================================================
Sub  txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus 		
	End If
End Sub



'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub  vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")
	
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
	Set gActiveSpdSheet = frm1.vspdData        
    
    If Row <= 0 Then
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
	
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)	    

	Call DbQuery("2")
    
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
    lgPageNo_B       = ""                                  'initializes Previous Key
    lgSortKey_B      = 1
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub  vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    gMouseClickStatus = "SPC"	'Split �����ڵ�    

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
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
	    
		Call SetSpreadColumnValue("A",frm1.vspdData,Col,NewRow)	        
    
		Call DbQuery("2")
     
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	
		lgPageNo_B       = ""                                  'initializes Previous Key
		lgSortKey_B      = 1
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub  vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")
	
    gMouseClickStatus = "SP2C"	'Split �����ڵ� 
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

	Call SetSpreadColumnValue("B",frm1.vspdData2,Col,Row)	
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ'
		If lgPageNo_A <> "" Then													'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
           Call DbQuery("1")
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'��: ������ üũ'
		If lgPageNo_B <> "" Then													'��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
           Call DbQuery("2")
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
Function  FncQuery() 
    FncQuery = False                                                        '��: Processing is NG
    Err.Clear     
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
		Exit Function
    End If
    
	If CompareDateByFormat(frm1.txtFromDt.text,frm1.txtToDt.text,frm1.txtFromDt.Alt,frm1.txtToDt.Alt, _
        	               "970025",frm1.txtFromDt.UserDefinedFormat,parent.gComDateType, true) = False Then
		frm1.txtFromDt.focus
		Exit Function
	End If
	
	If Trim(frm1.txtBizAreaCd.value) <> "" and   Trim(frm1.txtBizAreaCd1.value) <> "" Then				
		If Trim(frm1.txtBizAreaCd.value) > Trim(frm1.txtBizAreaCd1.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
			frm1.txtBizAreaCd.focus
			Exit Function
		End If
	End If
	
	if frm1.txtBizAreaCd.value <> "" then
	  If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  	Call DisplayMsgBox("970000","X",frm1.txtBizAreaCd.alt,"X")            '�� : No data is found. 
	  	frm1.txtBizAreaNm.value = ""
	  	frm1.txtBizAreaCd.focus
 	  	Exit Function
	  End If
	End If
	  
	if frm1.txtBizAreaCd1.value <> "" then
	  If CommonQueryRs(" A.BIZ_AREA_NM ","B_BIZ_AREA A","A.BIZ_AREA_CD = " & FilterVar(frm1.txtBizAreaCd1.value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	  	Call DisplayMsgBox("970000","X",frm1.txtBizAreaCd1.alt,"X")            '�� : No data is found.
	  	frm1.txtBizAreaNm1.value = ""
	  	frm1.txtBizAreaCd1.focus
 	  	Exit Function
	  End If
	End If
	
	If Trim(frm1.txtDealBpCd.value) = "" Then
		frm1.txtDealBpNm.value = ""
	End If	
	
	If Trim(frm1.txtPayBpCd.value) = "" Then
		frm1.txtPayBpNm.value = ""
	End If	
	
	If Trim(frm1.txtAcctCd.value) = "" Then
		frm1.txtAcctnm.value = ""
	End If	
	
	If Trim(frm1.txtBizAreaCd.value) = "" Then
		frm1.txtBizAreaNm.value = ""
	End If
	
	If Trim(frm1.txtBizAreaCd1.value) = "" Then
		frm1.txtBizAreaNm1.value = ""
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData    
	    
    Call InitVariables() 														'��: Initializes local global variables
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery("1")															'��: Query db data

    FncQuery = True		
	
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function  FncPrint() 
    Call parent.FncPrint()
    	
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
		
	Set gActiveElement = document.activeElement    
	
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
    	
	Set gActiveElement = document.activeElement    
	
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
    Dim iColumnLimit2
    
    If gMouseClickStatus = "SPCRP" Then
		iColumnLimit = 3
       
		ACol = Frm1.vspdData.ActiveCol
		ARow = Frm1.vspdData.ActiveRow

		If ACol > iColumnLimit Then
		   iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
		   Exit Function  
		End If   
    
		Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
    
		ggoSpread.Source = Frm1.vspdData
    
		ggoSpread.SSSetSplit(ACol)    
    
		Frm1.vspdData.Col = ACol
		Frm1.vspdData.Row = ARow
    
		Frm1.vspdData.Action = 0    
    
		Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If   
	'----------------------------------------
	' Spread�� �ΰ��� ��� 2��° Spread
	'----------------------------------------
    If gMouseClickStatus = "SP2CRP" Then
		iColumnLimit2 = 4
       
       ACol = Frm1.vspdData2.ActiveCol
       ARow = Frm1.vspdData2.ActiveRow

       If ACol > iColumnLimit2 Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit2 , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData2
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData2.Col = ACol
       Frm1.vspdData2.Row = ARow
    
       Frm1.vspdData2.Action = 0    
    
       Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If   
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function  FncExit()
    FncExit = True
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 
'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function  DbQuery(ByVal iOpt) 
	Dim strVal
	
    Err.Clear																						'��: Protect system from crashing
	On Error Resume Next
	
    DbQuery = False
    Call DisableToolBar(parent.TBC_QUERY)															'��: Disable Query Button Of ToolBar
	Call LayerShowHide(1)
    
    With frm1
		Select Case iOpt 
			Case "1" 
'--------------- ������ coding part(�������,Start)----------------------------------------------
				strVal = BIZ_PGM_ID & "?txtFromDt="		& Trim(.txtFromDt.Text)
				strVal = strVal & "&txtToDt="			& Trim(.txtToDt.Text)
				strVal = strVal & "&txtDealBpCd="		& Trim(.txtDealBpCd.value)
				strVal = strVal & "&txtPayBpCd="		& Trim(.txtPayBpCd.value)
				strVal = strVal & "&txtAcctCd="			& Trim(.txtAcctCd.value)
				strVal = strVal & "&txtDesc="			& Trim(.txtDesc.value)
				strVal = strVal & "&txtArNo="			& Trim(.txtArNo.value)
				strVal = strVal & "&txtRefNo="			& Trim(.txtRefNo.value)
				strVal = strVal & "&txtInvDocNo="		& Trim(.txtInvDocNo.value)
				strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
				strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
				strVal = strVal & "&txtDealBpCd_ALT="	& .txtDealBpCd.alt
				strVal = strVal & "&txtPayBpCd_ALT="	& .txtPayBpCd.alt
				strVal = strVal & "&txtAcctCd_ALT="		& .txtAcctCd.alt
				strVal = strVal & "&txtDesc_ALT="		& .txtDesc.alt
				strVal = strVal & "&txtArNo_ALT="		& .txtArNo.alt
				strVal = strVal & "&txtRefNo_ALT="		& .txtRefNo.alt
				strVal = strVal & "&txtInvDocNo_ALT="	& .txtInvDocNo.alt
				strVal = strVal & "&txtBizAreaCd_ALT="	& .txtBizAreaCd.alt
				strVal = strVal & "&txtBizAreaCd_ALT1="	& .txtBizAreaCd1.alt
				strVal = strVal & "&txtProject="		& Trim(.txtProject.value)
				
    '--------- Developer Coding Part (End) ----------------------------------------------------------
				strVal = strVal & "&lgPageNo="			& lgPageNo_A									'��: Next key tag
				strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))

				' ���Ѱ��� �߰� 
				strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
				strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
				strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
				strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

			Case "2"
	'--------------- ������ coding part(�������,Start)----------------------------------------------				
				strVal = BIZ_PGM_ID1 & "?txtArNo="		& GetKeyPosVal("A",1)

				strVal = strVal & "&lgPageNo="			& lgPageNo_B									'��: Next key tag
				strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("B")
				strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("B")
				strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("B"))
		End Select 
      
		Call RunMyBizASP(MyBizASP, strVal)															'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk(byval iOpt)																		'��: ��ȸ ������ ������� 
    lgIntFlgMode = parent.OPMD_UMODE																'��: Indicates that current mode is Update mode
    
	If iOpt = 1 Then
       Call vspdData_Click(1,1)
       frm1.vspdData.focus
	End If																							'��: This function lock the suitable field

	Call ggoOper.LockField(Document, "Q")															'��: This function lock the suitable field 
End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
			 					<TR>
									<TD CLASS="TD5" NOWRAP>�߻��Ⱓ</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="12" VIEWASTEXT > </OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="12" VIEWASTEXT > </OBJECT>');</SCRIPT>					
									</TD>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBizAreaCd(frm1.txtBizAreaCd.Value, 5)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=30 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ֹ�ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDealBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="1XXXXU" ALT="�ֹ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtDealBpCd.Value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtDealBpNm" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBizAreaCd(frm1.txtBizAreaCd1.Value, 6)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=30 tag="14"></TD>
								</TR>			 					
			 					</TR>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPayBpCd" SIZE=10 MAXLENGTH=10  STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="����ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtPayBpCd.Value, 2)">&nbsp;<INPUT TYPE=TEXT NAME="txtPayBpNm" SIZE=30 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>�����ڵ�</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="�����ڵ�" MAXLENGTH="20" SIZE=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ><IMG align=top name=btnCalType onclick="vbscript:CALL OpenAcctPopUp(frm1.txtAcctCd.value)" src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON"> 
														<INPUT NAME="txtAcctnm" ALT="�����ڵ��" MAXLENGTH="20" SIZE=25 tag  ="14"></TD>										
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ä�ǹ�ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtArNo" MAXLENGTH=18 STYLE="TEXT-ALIGN: left" tag="11XXXU" ALT="ä�ǹ�ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenARPopUp()"></TD>
									<TD CLASS=TD5 NOWRAP>�����ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtInvDocNo" ALT="�����ȣ" MAXLENGTH="50" STYLE="TEXT-ALIGN: Left" tag="11XXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDesc" ALT="���" MAXLENGTH="128" SIZE="30" tag="11XXXU" ></TD>
									<TD CLASS=TD5 NOWRAP>������ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRefNo" ALT="������ȣ" MAXLENGTH="30" STYLE="TEXT-ALIGN: Left" tag="11XXXU" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>������Ʈ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="������Ʈ" MAXLENGTH=25 SIZE=25 tag="1X"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="60%">
								<TD WIDTH="100%" colspan=6>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TDT NOWRAP>ä�Ǿ�(�ڱ�)</TD>
								<TD CLASS=TDT NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="ä�Ǿ�(�ڱ�)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TDT NOWRAP>�����ݾ�(�ڱ�)</TD>
								<TD CLASS=TDT NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotClsLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�(�ڱ�)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>													
								<TD CLASS=TDT NOWRAP>�ܾ�(�ڱ�)</TD>
								<TD CLASS=TDT NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotBalLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�ܾ�(�ڱ�)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" colspan=6>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>

