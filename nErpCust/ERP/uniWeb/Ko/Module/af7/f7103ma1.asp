<%@ LANGUAGE="VBSCRIPT" %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : PRERECEIPT
'*  3. Program ID           : f7103ma1
'*  4. Program Name         : �����ݳ�����ȸ 
'*  5. Program Desc         : �����ݳ�����ȸ 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/09/25
'*  8. Modified date(Last)  : 2001/03/05
'*  9. Modifier (First)     : Hee Jung, Kim
'* 10. Modifier (Last)      : 
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
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">				</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID        = "F7103MB1.asp"                         '��: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "F7103MB2.asp"                         '��: Biz logic spread sheet for #2
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey            = 5                                    '�١١١�: Max key value
Const C_MaxKey_B            = 2                                    '�١١١�: Max key value

Const C_SHEETMAXROWS_A	  = 30
Const C_SHEETMAXROWS_B    = 10
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Dim lgIsOpenPop                                             '��: Popup status                           

'��:--------Spreadsheet #1-----------------------------------------------------------------------------   

Dim lgPageNo_A                                              '��: Next Key tag                          
Dim lgSortKey_A                                             '��: Sort���� ���庯��                      

'��:--------Spreadsheet #2-----------------------------------------------------------------------------   

Dim lgPageNo_B                                              '��: Next Key tag                          
Dim lgSortKey_B                                             '��: Sort���� ���庯��                      

'��:--------Spreadsheet temp---------------------------------------------------------------------------   

Dim lgKeyPos                                                '��: Key��ġ                               
Dim lgKeyPosVal                                             '��: Key��ġ Value                         

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
'	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
'	EndDate = GetSvrDate
'	Call ExtractDateFrom(EndDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)
'	StartDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, "01")
'	EndDate = UNIConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)
    
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
    lgIntFlgMode     = parent.OPMD_CMODE                          'Indicates that current mode is Create mode

    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1

End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	dtToday = "<%=GetSvrDate%>"
	Call parent.ExtractDateFrom(dtToday, Parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")                                                     

	frm1.txtFromDt.text	= StartDate
	frm1.txtToDt.text	= EndDate
'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "MA") %>
End Sub
'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("A7103MA1","S","A","V20021211",parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey,"X","X")
    Call SetZAdoSpreadSheet("A7103MA2","S","B","V20021211",parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey_B,"X","X")

    Call SetSpreadLock ("1")
    Call SetSpreadLock ("2")

    
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(Byval iOpt )
    If iOpt = "1" Then
       With frm1
          .vspdData.ReDraw = False
			ggoSpread.Source = frm1.vspdData
			ggoSpread.SpreadLockWithOddEvenRowColor()	
          .vspdData.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SpreadLockWithOddEvenRowColor()	
            .vspdData2.ReDraw = True
       End With
    End If   
End Sub



 '**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 

'------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If lgIsOpenPop = True Then Exit Function
		
	lgIsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' ä�ǰ� ����(�ŷ�ó ����)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Call EscPopUp(iwhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
		lgBlnFlgChgValue = True
	End If

End Function
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
	Select Case iWhere
		Case 2
			arrParam(0) = "����������"									' �˾� ��Ī 
			arrParam(1) = "a_jnl_item"	 									' TABLE ��Ī 
			arrParam(2) = strCode											' Code Condition
			arrParam(3) = ""												' Name Cindition
			arrParam(4) = "jnl_type = " & FilterVar("pr", "''", "S") & "  "								' Where Condition
			arrParam(5) = "����������"									' �����ʵ��� �� ��Ī 

			arrField(0) = "JNL_CD"											' Field��(0)
			arrField(1) = "JNL_NM"											' Field��(1)
    
			arrHeader(0) = "����������"									' Header��(0)
			arrHeader(1) = "������������"								' Header��(1)
	End Select
	

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		Call EscPopUp(iWhere)
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function EscPopUp(Byval iWhere)

	With frm1
		Select Case iWhere
			Case 1
				.txtBpCd.focus
			Case 2	
				.txtPrrcptType.focus
		End Select
	End With
	
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)

	With frm1
		Select Case iWhere
			Case 1
				.txtBpCd.value = arrRet(0)
				.txtBpNm.value = arrRet(1)	
				.txtBpCd.focus
			Case 2	
				.txtPrrcptType.value = arrRet(0)
				.txtPrrcptTypeNm.value = arrRet(1)		
				.txtPrrcptType.focus
		End Select
	End With
	
End Function

Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = frm1.txtFromDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToDt.Text
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' �������� ���� Condition  
	
	' ���Ѱ��� �߰� 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
End Function

'------------------------------------------  SetDept()  --------------------------------------------------
'	Name : SetDept()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetDept(Byval arrRet)
		frm1.hOrgChangeId.value=arrRet(2)
		
		frm1.txtDeptCd.value = arrRet(0)
		frm1.txtDeptNm.value = arrRet(1)		
		frm1.txtFromDt.text = arrRet(4)
		frm1.txtToDt.text = arrRet(5)
		frm1.txtDeptCd.focus
End Function


'==========================================================================================
'   Event Name : txtDeptCd_Onchange
'   Event Desc : 
'==========================================================================================
Sub txtDeptCD_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromDt.Text, gDateFormat,""), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtToDt.Text, gDateFormat,""), "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------

End Sub
'==========================================================================================
'   Event Name : OpenOrderBy
'   Event Desc : 
'==========================================================================================
	
Function OpenOrderBy()
	Dim arrRet, lgIsOpenPop
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
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
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
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
	End Select
	
	lgBlnFlgChgValue = True
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
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000111")							'��: ��ư ���� ���� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------   	
	frm1.txtFromDt.focus
	Set gActiveElement = document.activeElement

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

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'========================================================================================================
'   Event Name : DblClick
'   Event Desc :
'=========================================================================================================
Sub txtFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromDt.Action = 7
	    Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus  
	End if
End Sub

Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus  
	End if
End Sub

'========================================================================================================
'   Event Name : KeyPress
'   Event Desc :
'========================================================================================================
Sub txtFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToDt.focus
		Call FncQuery
	ENd if
		
End Sub

Sub txtToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtFromDt.focus
		Call FncQuery
	ENd if
End Sub


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
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")
    Dim ii
    
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    
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
     
'--------------- ������ coding part(�������,Start)----------------------------------------------------
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
End Sub


'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")
    Dim ii
    
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
	If Button  <> "1"  And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button <> "1"And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
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
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ'
		If lgPageNo_A <> "" Then                            '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(parent.TBC_QUERY)
            If DbQuery("1") = False Then
              Call RestoreToolBar()
              Exit Sub
			End IF
		End If
	End if
      
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'��: ������ üũ'
		If lgPageNo_B <> "" Then                            '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery("2") = False Then
              Call RestoreToolBar()
              Exit Sub
          End if
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
	Dim IntRetCD 

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear     

    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables 														'��: Initializes local global variables
    
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
	
	If Trim(frm1.txtPrrcptType.value)="" then
		frm1.txtPrrcptTypeNm.value="" 
	End if
	
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

Function FncPrint() 
    Call parent.FncPrint()
	    		
	Set gActiveElement = document.activeElement    

End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call parent.FncExport(C_MULTI)
	    		
	Set gActiveElement = document.activeElement    

End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
	    		
	Set gActiveElement = document.activeElement    

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
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 
'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function DbQuery(ByVal iOpt) 
	Dim strVal
	DbQuery = False

	Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)

	With frm1
		If iOpt = "1" Then
		'--------------- ������ coding part(�������,Start)----------------------------------------------

		        
			strVal = BIZ_PGM_ID & "?txtFromDt="		& Trim(.txtFromDt.Text)
			strVal = strVal & "&txtToDt="			& Trim(.txtToDt.Text)
			strVal = strVal & "&txtDeptCd="			& Trim(.txtDeptCd.value) 
			strVal = strVal & "&txtBpCd="			& Trim(.txtBpCd.value) 
			strVal = strVal & "&txtPrrcptType="		& Trim(.txtPrrcptType.value)   
			strVal = strVal & "&txtDeptCd_Alt="		& .txtDeptCd.Alt
			strVal = strVal & "&txtBpCd_Alt="		& .txtBpCd.Alt
			strVal = strVal & "&txtBizAreaCd="		& Trim(.txtBizAreaCd.value)
			strVal = strVal & "&txtBizAreaCd1="		& Trim(.txtBizAreaCd1.value)
			strVal = strVal & "&txtBizAreaCd_Alt="	& Trim(frm1.txtBizAreaCd.alt)
			strVal = strVal & "&txtBizAreaCd1_Alt="	& Trim(frm1.txtBizAreaCd1.alt)
			strVal = strVal & "&txtPrrcptType_Alt="	& Trim(frm1.txtPrrcptType.alt)		   

			strVal = strVal & "&OrgChangeId="		& Trim(.hOrgChangeId.Value)     
			strVal = strVal & "&lgPageNo="			& lgPageNo_A                          '��: Next key tag
			strVal = strVal & "&lgMaxCount="		& Cstr(C_SHEETMAXROWS_A)
			strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")		'lgSelectListDT
			strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")		'lgMaxFieldCount,lgPopUpR,parent.gFieldCD,parent.gNextSeq,parent.gTypeCD(0),parent.C_MaxSelList)
			strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))

			' ���Ѱ��� �߰� 
			strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
			strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
			strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
			strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
		Else          
			strVal = BIZ_PGM_ID1 & "?txtPrrcptNo="	& GetKeyPosVal("A",1)
			strVal = strVal & "&lgPageNo="			& lgPageNo_B                          '��: Next key tag
			strVal = strVal & "&lgMaxCount="		& Cstr(C_SHEETMAXROWS_B)
			strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("B")		'lgSelectListDT
			strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("B")		'lgMaxFieldCount,lgPopUpR,parent.gFieldCD,parent.gNextSeq,parent.gTypeCD(0),parent.C_MaxSelList)
			strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("B"))
		End If   

		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	End With
	    
	DbQuery = True
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk(byval iOpt)														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11000000000111")							'��: ��ư ���� ���� 
	
	If iOpt = 1 Then
       Call vspdData_Click(1,1)
       frm1.vspdData.focus
	End If							                                     '��: This function lock the suitable field

	Call ggoOper.LockField(Document, "Q")	
	
End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
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

					<TD WIDTH="*">&nbsp;</TD>
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
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="12X1" VIEWASTEXT id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="������" tag="12X1" VIEWASTEXT id=fpDateTime2></OBJECT>');</SCRIPT></TD>
			 						<TD CLASS="TD5" NOWRAP>�μ�</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�μ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">&nbsp;
										<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=30 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 1)">&nbsp;
										<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ۻ����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBizAreaCd(frm1.txtBizAreaCd.value, 0)">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=25 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����������</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPrrcptType" SIZE=10 MAXLENGTH=10  tag="11XXXU" ALT="����������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtPrrcptType.value, 2)">&nbsp;
										<INPUT TYPE=TEXT NAME="txtPrrcptTypeNm" SIZE=30 tag="14XXXU" ALT="������������"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="60%" WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData width="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD HEIGHT="40%" WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=  <%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows"     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd"       tag="24" TABINDEX="-1"> 
<INPUT TYPE=HIDDEN NAME="hItemCd"        tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hRoutNo"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"	 tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd"  tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizAreaCd1" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
