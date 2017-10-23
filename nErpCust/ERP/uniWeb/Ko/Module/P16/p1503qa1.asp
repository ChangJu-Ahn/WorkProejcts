<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1503qa1
'*  4. Program Name         : �ڿ���Shift��ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/12/13
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Ryu Sung Won
'* 11. Comment              :
'**********************************************************************************************-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'############################################################################################################-->
<!--'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'************************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ===========================================
'============================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'=====================================  1.1.2 ���� Include   =============================================
'===========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '��: indicates that All variables must be declared in advance

'****************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                             '��: Popup status                           

'��:--------Spreadsheet #1-----------------------------------------------------------------------------   
Dim lgTypeCD_A                                              '��: 'G' is for group , 'S' is for Sort    
Dim lgFieldCD_A                                             '��: �ʵ� �ڵ尪                           
Dim lgFieldNM_A                                             '��: �ʵ� ����                           
Dim lgFieldLen_A                                            '��: �ʵ� ��(Spreadsheet����)              
Dim lgFieldType_A                                           '��: �ʵ� ����                           
Dim lgDefaultT_A                                            '��: �ʵ� �⺻��                           
Dim lgNextSeq_A                                             '��: �ʵ� Pair��                           
Dim lgKeyTag_A                                              '��: Key ����                              

Dim lgSelectList_A                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT_A                                        '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgPopUpR_A                                              '��: Orderby,Groupby default ��            

Dim lgSortFieldNm_A                                         '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD_A                                         '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgStrPrevKey_A                                          '��: Next Key tag                          
Dim lgSortKey_A                                             '��: Sort���� ���庯��                      

'��:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim lgTypeCD_B                                              '��: 'G' is for group , 'S' is for Sort    
Dim lgFieldCD_B                                             '��: �ʵ� �ڵ尪                           
Dim lgFieldNM_B                                             '��: �ʵ� ����                           
Dim lgFieldLen_B                                            '��: �ʵ� ��(Spreadsheet����)              
Dim lgFieldType_B                                           '��: �ʵ� ����                           
Dim lgDefaultT_B                                            '��: �ʵ� �⺻��                           
Dim lgNextSeq_B                                             '��: �ʵ� Pair��                           
Dim lgKeyTag_B                                              '��: Key ����                              

Dim lgSelectList_B                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT_B                                        '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgPopUpR_B                                              '��: Orderby,Groupby default ��            

Dim lgSortFieldNm_B                                         '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD_B                                         '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgStrPrevKey_B                                          '��: Next Key tag                          
Dim lgSortKey_B                                             '��: Sort���� ���庯��                      

'��:--------Spreadsheet temp---------------------------------------------------------------------------   
                                                               '��:--------Buffer for Spreadsheet -----   
Dim lgTypeCD_T                                              '��: 'G' is for group , 'S' is for Sort    
Dim lgFieldCD_T                                             '��: �ʵ� �ڵ尪                           
Dim lgFieldNM_T                                             '��: �ʵ� ����                           
Dim lgFieldLen_T                                            '��: �ʵ� ��(Spreadsheet����)              
Dim lgFieldType_T                                           '��: �ʵ� ����                           
Dim lgDefaultT_T                                            '��: �ʵ� �⺻��                           
Dim lgNextSeq_T                                             '��: �ʵ� Pair��                           
Dim lgKeyTag_T                                              '��: Key ����                              

Dim lgSelectList_T                                          '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgSelectListDT_T                                        '��: SpreadSheet�� �ʱ�  ��ġ�������� ���� 
Dim lgPopUpR_T                                              '��: Orderby,Groupby default ��            
Dim lgMark_T                                                '��: ��ũ                                  

Dim lgSortFieldNm_T                                         '��: Orderby popup�� ����Ÿ(�ʵ弳��)      
Dim lgSortFieldCD_T                                         '��: Orderby popup�� ����Ÿ(�ʵ��ڵ�)      

Dim lgKeyPos                                                '��: Key��ġ                               
Dim lgKeyPosVal                                             '��: Key��ġ Value                         

Dim StartDate, EndDate


StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
StartDate = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
EndDate   = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
    
'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "p1503qb1.asp"                         '��: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "p1503qb2.asp"                         '��: Biz logic spread sheet for #2
Const BIZ_PGM_JUMP_ID   = "p1504qa1.asp"				  	       '��: �����Ͻ� ���� ASP�� 
Const C_MaxKey            = 2                                    '�١١١�: Max key value

Dim lsPoNo                                                 '��: Jump�� Cookie�� ���� Grid value
Dim	lgTopLeft

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
	lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False                               'Indicates that no value changed

    lgStrPrevKey_A   = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgStrPrevKey_B   = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1
End Sub

'==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'                 lgSort...�� �����ϴ� ���� ������ sort��� ����� ���� 
'                 IsPopUpR ���������� sort ������ �⺻�� �Ǵ� �� ���� 
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFromDt.Text	= startdate
	frm1.txtToDt.Text	= UniConvYYYYMMDDToDate(parent.gDateFormat, "2999","12","31")
End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'====================DBQUERY=======================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "QA") %>
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet(Byval iOpt)
    Call AppendNumberPlace("6","2","0")
	Call SetZAdoSpreadSheet("P1503QA1","S","A","V20021210", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit2(1)
	Call SetZAdoSpreadSheet("P1503QA1","S","B","V20021210", Parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X" )
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.SSSetSplit2(1)
	Call SetSpreadLock("A") 
	Call SetSpreadLock("B") 
End Sub

'========================================= 2.7 SetSpreadLock() ===========================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub SetSpreadLock(Byval iOpt )
    If iOpt = "A" Then
       ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
    Else
       ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
    End If   
End Sub

'**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 

'------------------------------------------  OpenConItemCd()  -------------------------------------------------
'	Name : OpenConItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Item Code
	arrParam(1) = Trim(frm1.txtItemCd.value)
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
	
    arrField(0) = 1 							' Field��(0) : "ITEM_CD"
    arrField(1) = 2 							' Field��(1) : "ITEM_NM"
    
    iCalledAspName = AskPRAspName("b1b11pa1")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "b1b11pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenSppl()  -------------------------------------------------
'	Name : OpenConRouting()
'	Description : Routing PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConRouting()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "�����"												' �˾� ��Ī 
	arrParam(1) = "P_ROUTING_HEADER"										' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtRoutNo.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "	' Where Condition
	arrParam(5) = "�����"												' TextBox ��Ī 
	
    arrField(0) = "ROUT_NO"												' Field��(0)
    arrField(1) = "DESCRIPTION"												' Field��(1)
    arrField(2) = "MAJOR_FLG"												' Field��(1)
    
    arrHeader(0) = "�����"												' Header��(0)
    arrHeader(1) = "����ø�"											' Header��(1)
    arrHeader(2) = "�ֶ����"										' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtRoutNo.Value = arrRet(0)
		frm1.txtRoutNm.Value = arrRet(1)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutNo.focus
	
End Function

'------------------------------------------  OpenConPlant()  -------------------------------------------------
'	Name : OpenConPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"							' �˾� ��Ī 
	arrParam(1) = "B_PLANT"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Condition
	arrParam(5) = "����"							' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"							' Field��(0)
    arrField(1) = "PLANT_NM"							' Field��(1)
        
    arrHeader(0) = "����"						' Header��(0)
    arrHeader(1) = "�����"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")		

	lgIsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenResource()  -------------------------------------------------
'	Name : OpenResource()
'	Description : Resource PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResource()

	Dim arrRet
	Dim arrParam(5), arrField(6),arrHeader(6)


	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
			
	lgIsOpenPop = True
	arrParam(0) = "�ڿ��˾�"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = Trim(frm1.txtResourceCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " "			
	arrParam(5) = "�ڿ�"
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ�"		
    arrHeader(1) = "�ڿ���"
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False

	If arrRet(0) <> "" Then
		frm1.txtResourceCd.Value = arrRet(0)
		frm1.txtResourceNm.Value = arrRet(1)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtResourceCd.focus
		
End Function

'------------------------------------------  OpenResourceGroup()  -------------------------------------------------
'	Name : OpenResourceGroup()
'	Description : ResourceGroup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenResourceGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then
		lgIsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtResourceGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	lgIsOpenPop = True

	arrParam(0) = "�ڿ��׷��˾�"	
	arrParam(1) = "P_RESOURCE_GROUP"				
	arrParam(2) = Trim(frm1.txtResourceGroupCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " " 
				  			
	arrParam(5) = "�ڿ��׷�"			
	    
    arrField(0) = "RESOURCE_GROUP_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "�ڿ��׷�"		
    arrHeader(1) = "�ڿ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtResourceGroupCd.Value = arrRet(0)
		frm1.txtResourceGroupNm.Value = arrRet(1)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtResourceGroupCd.focus
	
End Function


'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenOrderBy()
	Dim arrRet
	
	On Error Resume Next

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
  
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
		Call InitVariables
		Call InitSpreadSheet("A")
	End If
End Function

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
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    ReDim lgPopUpR_A(parent.C_MaxSelList - 1,1)
    ReDim lgPopUpR_B(parent.C_MaxSelList - 1,1)

	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	

	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")

    Call SetToolbar("11000000000011")							'��: ��ư ���� ���� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------

	If parent.gPlant <> "" then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtResourceCd.focus 
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
	End If
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

'========================================================================================================
'   Event Name : txtPoFrDt
'   Event Desc :
'=========================================================================================================
Sub txtFromDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtFromDt.Focus
	End if
End Sub
'========================================================================================================
'   Event Name : txtPoToDt
'   Event Desc :
'========================================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		Frm1.txtToDt.Focus
	End if
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtFromDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_KeyDown(keycode, shift)
'   Event Desc : Enter Event�� FncQuery�Ѵ�.
'=======================================================================================================
Sub txtToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim ii
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

	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
	
	Call DisableToolBar(parent.TBC_QUERY)   
	If DbQuery("B") = False Then
		Call RestoreToolBar()
		Exit Sub
	End If
     
    frm1.vspdData2.MaxRows = 0
    lgStrPrevKey_B   = ""                                  'initializes Previous Key
    lgSortKey_B      = 1
     
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    Dim ii
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
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
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

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	'��: ������ üũ'
		If lgStrPrevKey_A <> "" Then                            '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			lgTopLeft = "Y"
			Call DisableToolBar(TBC_QUERY)  
			If DbQuery("A") = False Then
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
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) Then	'��: ������ üũ'
		If lgStrPrevKey_B <> "" Then                        '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
			Call DisableToolBar(TBC_QUERY)  
			If DbQuery("B") = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
           
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
    
    Err.Clear     


    '-----------------------
    'Erase contents area
    '-----------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If	
	
	If frm1.txtResourceCd.value = "" Then
		frm1.txtResourceNm.value = ""
	End If
	
	If frm1.txtResourceGroupCd.value = "" Then
		frm1.txtResourceGroupNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables

    If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then
		Exit Function
	End If
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If

    '-----------------------
    'Query function call area
    '-----------------------
	If DbQuery("A") = False Then   
		Exit Function           
    End If     
    															'��: Query db data

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
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)
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
       iColumnLimit  = frm1.vspdData.MaxCols - 1
       
       ACol = Frm1.vspdData.ActiveCol
       ARow = Frm1.vspdData.ActiveRow

       If ACol > iColumnLimit Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData.Col = ACol
       Frm1.vspdData.Row = ARow
    
       Frm1.vspdData.Action = 0    
    
       Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
    End If   
	
	'----------------------------------------
	' Spread�� �ΰ��� ��� 2��° Spread
	'----------------------------------------
	
	
    If gMouseClickStatus = "SP2CRP" Then
		iColumnLimit2 = frm1.vspdData.MaxCols - 1
       
       ACol = Frm1.vspdData2.ActiveCol
       ARow = Frm1.vspdData2.ActiveRow

       If ACol > iColumnLimit2 Then
          iRet = DisplayMsgBox("900030", "X", iColumnLimit2 , "X")
          Exit Function  
       End If   
    
       Frm1.vspdData2.ScrollBars = SS_SCROLLBAR_NONE
    
       ggoSpread.Source = Frm1.vspdData2
    
       ggoSpread.SSSetSplit(ACol)    
    
       Frm1.vspdData2.Col = ACol
       Frm1.vspdData2.Row = ARow
    
       Frm1.vspdData2.Action = 0    
    
       Frm1.vspdData2.ScrollBars = SS_SCROLLBAR_BOTH
    End If   
    
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
'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function DbQuery(ByVal iOpt) 
	Dim strVal
	Dim ResourceGroupCd1, ResourceGroupCd2, ToDt1, ToDt2

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	LayerShowHide(1) 

    With frm1
		
		If .txtResourceGroupCd.value = "" Then
			ResourceGroupCd1 = ""
			ResourceGroupCd2 = "zzzzzzzzzz"
		Else
			ResourceGroupCd1 = .txtResourceGroupCd.value
			ResourceGroupCd2 = .txtResourceGroupCd.value
		End If
				
		If .txtFromDt.text = "" Then
			ToDt1 = "1900-01-01"
		Else
			ToDt1 = .txtFromDt.text
		End If
				
		If .txtToDt.text = "" Then
			ToDt2 = "2999-12-31"
		Else
			ToDt2 = .txtToDt.text
		End If

		If iOpt = "A" Then
'--------------- ������ coding part(�������,Start)----------------------------------------------
           strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
           strVal = strVal & "&txtResourceCd=" & Trim(.txtResourceCd.value)
           strVal = strVal & "&txtResourceGroupCd1=" & Trim(ResourceGroupCd1)
           strVal = strVal & "&txtResourceGroupCd2=" & Trim(ResourceGroupCd2)
           strVal = strVal & "&txtToDt1=" & Trim(ToDt1)
           strVal = strVal & "&txtToDt2=" & Trim(ToDt2)
           strVal = strVal & "&iOpt=" & iOpt
        Else   
           strVal = BIZ_PGM_ID1 & "?txtPlantCd=" & Trim(.txtPlantCd.value)
           strVal = strVal & "&txtResourceCd=" & GetKeyPosVal("A",1)
           strVal = strVal & "&iOpt=" & iOpt
          
        End If   

'--------------- ������ coding part(�������,End)------------------------------------------------
        If iOpt = "A" Then
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey_A                      '��: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A") 'lgSelectListDT_A
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A") 'MakeSql()
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A")) 'EnCoding(lgSelectList_A)
        Else   
           strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey_B                      '��: Next key tag
           strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B") 'lgSelectListDT_B
           strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B") 'MakeSql()
           strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B")) 'EnCoding(lgSelectList_B)
        End If
        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk(ByVal iOpt)														'��: ��ȸ ������ ������� 
	
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If	
	
	lgIntFlgMode = parent.OPMD_UMODE											'��: Indicates that current mode is Update mode 

	Call ggoOper.LockField(Document, "Q")								'��: This function lock the suitable field 
	Call SetToolbar("11000000000111")		
	lgBlnFlgChgValue = False
	
	If iOpt = "A" Then
		If lgTopLeft <> "Y" Then
			Call vspdData_Click(1,1)
		End If
		lgTopLeft = "N"
	End If
	
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'=========================================================================================================
' Function Name : CopyPopupInfABT
' Function Desc : set popup information according to iOpt
'===========================================================================================================
Sub CopyPopupInfABT(Byval iOpt)
    Dim ii
    Call CopyTBL(iOpt)    
    If iOpt = "1" Then
       For ii = 0 to  parent.C_MaxSelList - 1
           lgPopUpR_T(ii,0)   =   lgPopUpR_A(ii,0)  
           lgPopUpR_T(ii,1)   =   lgPopUpR_A(ii,1)  
       Next
       
       ReDim lgSortFieldCD_T(UBound(lgSortFieldCD_A))
       ReDim lgSortFieldNM_T(UBound(lgSortFieldNM_A))

       For ii = 0 to UBound(lgSortFieldCD_A)
           lgSortFieldCD_T(ii) = lgSortFieldCD_A(ii)
           lgSortFieldNM_T(ii) = lgSortFieldNM_A(ii)
       Next
    Else
       For ii = 0 to  parent.C_MaxSelList - 1
           lgPopUpR_T(ii,0)   =   lgPopUpR_B(ii,0)  
           lgPopUpR_T(ii,1)   =   lgPopUpR_B(ii,1)  
       Next

       ReDim lgSortFieldCD_T(UBound(lgSortFieldCD_B))
       ReDim lgSortFieldNM_T(UBound(lgSortFieldNM_B))

       For ii = 0 to UBound(lgSortFieldCD_B)
           lgSortFieldCD_T(ii) = lgSortFieldCD_B(ii)
           lgSortFieldNM_T(ii) = lgSortFieldNM_B(ii)
       Next
    End If       
End Sub

'=========================================================================================================
' Function Name : CopyPopupInfTAB
' Function Desc : set popup information according to iOpt
'===========================================================================================================
Sub CopyPopupInfTAB(Byval iOpt)
    Dim ii
    If iOpt = "1" Then
          
       For ii = 0 to  parent.C_MaxSelList - 1
           lgPopUpR_A(ii,0)   =   lgPopUpR_T(ii,0)      
           lgPopUpR_A(ii,1)   =   lgPopUpR_T(ii,1)      
       Next
       
       lgSelectList_A        =   lgSelectList_T  
       lgSelectListDT_A      =   lgSelectListDT_T
    Else

       For ii = 0 to  parent.C_MaxSelList - 1
           lgPopUpR_B(ii,0)   =   lgPopUpR_T(ii,0)      
           lgPopUpR_B(ii,1)   =   lgPopUpR_T(ii,1)      
       Next
       lgSelectList_B        =   lgSelectList_T  
       lgSelectListDT_B      =   lgSelectListDT_T
    End If       
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ڿ���Shift��ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*" align=right><button name="btnAutoSel" class="clsmbtn" ONCLICK="OpenOrderBy()">���ļ���</button></td>
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
			 						<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP>
										<script language =javascript src='./js/p1503qa1_I799951517_txtFromDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p1503qa1_I375921899_txtToDt.js'></script>					
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ڿ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>�ڿ��׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtResourceGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="�ڿ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResourceGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceGroupNm" SIZE=25 tag="14"></TD>
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
							<TR HEIGHT="100%">
								<TD WIDTH="50%" colspan=4>
								<script language =javascript src='./js/p1503qa1_I755646252_vspdData.js'></script></TD>
								<TD WIDTH="50%" colspan=4>
								<script language =javascript src='./js/p1503qa1_vaSpread1_vspdData2.js'></script></TD>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>
