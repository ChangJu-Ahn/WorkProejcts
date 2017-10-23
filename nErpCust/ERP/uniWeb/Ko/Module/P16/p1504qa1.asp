<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1504qa1
'*  4. Program Name         : Shift������ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/12/16
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

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "p1504qb1.asp"                         '��: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "p1504qb2.asp"                         '��: Biz logic spread sheet for #2

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
	
    lgBlnFlgChgValue = False                               'Indicates that no value changed
	lgIntFlgMode = parent.OPMD_CMODE
	
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
'--------------- ������ coding part(�������,Start)--------------------------------------------------
	
'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "P", "NOCOOKIE", "QA") %>
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("P1504QA1","S","A","V20021210", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetSplit2(1)
	Call SetZAdoSpreadSheet("P1504QA1","S","B","V20021210", Parent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X" )
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
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
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
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & " "			
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

'------------------------------------------  OpenShift()  -------------------------------------------------
'	Name : OpenShift()
'	Description : OpenShift PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenShift()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then
		lgIsOpenPop = False
		Exit Function
	End If
	
	If UCase(frm1.txtShiftCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	lgIsOpenPop = True

	arrParam(0) = "Shift�˾�"	
	arrParam(1) = "P_SHIFT_HEADER"				
	arrParam(2) = Trim(frm1.txtShiftCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & " " 
				  			
	arrParam(5) = "Shift"			
	    
    arrField(0) = "SHIFT_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "Shift"		
    arrHeader(1) = "Shift��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtShiftCd.Value = arrRet(0)
		frm1.txtShiftNm.Value = arrRet(1)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtShiftCd.focus
	
End Function

'===========================================================================
' Sub Name : PopZAdoConfigGrid
' Sub Desc : 
'===========================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderBy(gActiveSpdSheet.Id)
End Sub

'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
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
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

    ReDim lgPopUpR_A(parent.C_MaxSelList - 1,1)
    ReDim lgPopUpR_B(parent.C_MaxSelList - 1,1)

	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
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
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    Call SetPopupMenuItemInf("00000000001")
    
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
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
    gMouseClickStatus = "SP2C"
    Set gActiveSpdSheet = frm1.vspdData2
	Call SetPopupMenuItemInf("00000000001")
	    
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
			Call DisableToolBar(parent.TBC_QUERY)  
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
			Call DisableToolBar(parent.TBC_QUERY)  
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
	
	If frm1.txtShiftCd.value = "" Then
		frm1.txtShiftNm.value = ""
	End If
	
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
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
    Call parent.FncFind(parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
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
	LayerShowHide(1) 
		
	with frm1
		
        If iOpt = "A" Then
'--------------- ������ coding part(�������,Start)----------------------------------------------
           strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value)
           strVal = strVal & "&txtResourceCd=" & Trim(.txtResourceCd.value)
           strVal = strVal & "&txtShiftCd=" & Trim(.txtShiftCd.value)
           strVal = strVal & "&iOpt=" & iOpt
        Else   
           strVal = BIZ_PGM_ID1 & "?txtResourcePlantCd=" & Trim(.txtPlantCd.value)
           strVal = strVal & "&txtShiftPlantCd=" & Trim(.txtPlantCd.value)
           strVal = strVal & "&txtShiftCd=" & GetKeyPosVal("A",1)
           strVal = strVal & "&txtResourceCd=" & GetKeyPosVal("A",2)
           strVal = strVal & "&iOpt=" & iOpt
        End If   
        
'--------------- ������ coding part(�������,End)------------------------------------------------
        If iOpt = "A" Then
           strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey_A                      '��: Next key tag
           strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
           strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
           strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
        Else   
           strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey_B                      '��: Next key tag
           strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
           strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
           strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))
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
	
	lgIntFlgMode = parent.OPMD_UMODE
	
    Call ggoOper.LockField(Document, "Q")								'��: This function lock the suitable field 
	Call SetToolbar("11000000000111")		
	lgBlnFlgChgValue = False
	
	If iOpt = "A" Then
		If lgTopLeft <> "Y" Then
			Call vspdData_Click(1,1)
		End If
		lgTopLeft = "N"
	End If						'��: This function lock the suitable field
	

End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################


'========================================================================================
' Function Name : MakeSql()
' Function Desc : Order by ���� group by ���� �����.
'========================================================================================
Function MakeSql()
    Dim iStr,jStr
    Dim ii,jj,kk      '<  ----------------------���� ���� 
    Dim iFirst

    Dim tmpPopUpR     '<  ----------------------���� ���� 
    
    '2001��3��30�� �߰� �ڵ�(����)    
    Redim tmpPopUpR(C_MaxSelList - 1)    
    For kk = 0 to C_MaxSelList - 1
        tmpPopUpR(kk) = lgPopUpR_T(kk,0)
    Next

    '2001��3��30�� �߰� �ڵ�(����)

    
    iFirst = "N"
    iStr   = ""  
    jStr   = ""      

    Redim  lgMark_T(UBound(lgFieldNM_T)) 
    
    For ii = 0 to C_MaxSelList - 1
        If tmpPopUpR(ii) <> "" Then     ' '2001��3��30�� �ڵ� ����  lgPopUpR -> tmpPopUpR
           If lgTypeCD_T(0) = "G" Then
              For jj = 0 To UBound(lgFieldNM_T) - 1                                            'Sort ��󸮽�Ʈ   ���� 
                  If lgMark_T(jj) <> "X" Then
                     If lgPopUpR_T(ii,0) = lgFieldCD_T(jj) Then
                        If iFirst = "Y" Then
                           iStr = iStr & " , "
                           jStr = jStr & " , " 
                        End If   
                        If CInt(Trim(lgNextSeq_T(jj))) >= 1 And CInt(Trim(lgNextSeq_T(jj))) <= UBound(lgFieldNM_T) Then
                           iStr = iStr & " " & lgPopUpR_T(ii,0) & " " & lgPopUpR_T(ii,1) & " , " & lgFieldCD_T(CInt(lgNextSeq_T(jj)) - 1)
                           jStr = jStr & " " & lgPopUpR_T(ii,0) & " " &                    " , " & lgFieldCD_T(CInt(lgNextSeq_T(jj)) - 1)
                           lgMark_T(CInt(lgNextSeq_T(jj)) - 1) = "X"

                           '2001��3��30�� �߰� �ڵ�(����)
                           If (ii + 1) <  C_MaxSelList   Then
                              For kk = ii + 1 to C_MaxSelList - 1
                                  If  lgPopUpR_T(kk,0) = lgFieldCD_T(CInt(lgNextSeq_T(jj)) - 1)  Then
                                      iStr = iStr & " " & lgPopUpR_T(kk,1) 
                                      tmpPopUpR(kk) = ""
                                  End If    
                              Next    
                           End If   
                           '2001��3��30�� �߰� �ڵ�(��)

                        Else
                          iStr = iStr & " " & lgPopUpR_T(ii,0) & " " & lgPopUpR_T(ii,1)
                          jStr = jStr & " " & lgPopUpR_T(ii,0) 
                        End If
                        iFirst = "Y"
                        lgMark_T(jj) = "X"
                     End If
                     
                  End If
              Next
           Else
              If iFirst = "Y" Then
                 iStr = iStr & " , "
                 jStr = jStr & " , " 
              End If   
              iStr = iStr & " " & lgPopUpR_T(ii,0) & " " & lgPopUpR_T(ii,1)
              iFirst = "Y"
           End If
              
        End If
    Next     
    
    If Trim(iStr) <> "" Then
      iStr =  " Order By " & iStr 
    End If
    
    If lgTypeCD_T(0) = "G" Then
       MakeSql = "Group By " & jStr & iStr
    Else
       MakeSql = iStr
    End If   
End Function

'=========================================================================================================
' Function Name : CopyPopupInfABT
' Function Desc : set popup information according to iOpt
'===========================================================================================================
Sub CopyPopupInfABT(Byval iOpt)
    Dim ii
    Call CopyTBL(iOpt)    
    If iOpt = "1" Then
       For ii = 0 to  C_MaxSelList - 1
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
       For ii = 0 to  C_MaxSelList - 1
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
          
       For ii = 0 to  C_MaxSelList - 1
           lgPopUpR_A(ii,0)   =   lgPopUpR_T(ii,0)      
           lgPopUpR_A(ii,1)   =   lgPopUpR_T(ii,1)      
       Next
       
       lgSelectList_A        =   lgSelectList_T  
       lgSelectListDT_A      =   lgSelectListDT_T
    Else

       For ii = 0 to  C_MaxSelList - 1
           lgPopUpR_B(ii,0)   =   lgPopUpR_T(ii,0)      
           lgPopUpR_B(ii,1)   =   lgPopUpR_T(ii,1)      
       Next
       lgSelectList_B        =   lgSelectList_T  
       lgSelectListDT_B      =   lgSelectListDT_T
    End If       
End Sub



'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'#########################################################################################################-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Shift������ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH="*" align=right><!--<button name="btnAutoSel" class="clsmbtn" ONCLICK="OpenOrderBy()">���ļ���</button>--></td>
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
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ڿ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtResourceCd" SIZE=15 MAXLENGTH=10 tag="11XXXU" ALT="�ڿ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenResource()">&nbsp;<INPUT TYPE=TEXT NAME="txtResourceNm" SIZE=30 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>Shift</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtShiftCd" SIZE=5 MAXLENGTH=2 tag="11XXXU" ALT="�ڿ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenShift()">&nbsp;<INPUT TYPE=TEXT NAME="txtShiftNm" SIZE=30 tag="14"></TD>
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
								<TD WIDTH="40%" colspan=4>
								<script language =javascript src='./js/p1504qa1_A_vspdData.js'></script></TD>
								<TD WIDTH="60%" colspan=4>
								<script language =javascript src='./js/p1504qa1_B_vspdData2.js'></script></TD>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hRoutNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
</HTML>
