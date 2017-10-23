<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : PrePayment management
'*  3. Program ID           : a2111ma1.asp
'*  4. Program Name         : ��ǥ�����׸���ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003.08.28
'*  8. Modified date(Last)  : 2003.08.28
'*  9. Modifier (First)     : Jeong Yong Kyun
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2001.01.13
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
Const BIZ_PGM_ID        = "a2111mb1.asp"							'��: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "a2111mb2.asp"							'��: Biz logic spread sheet for #2
Const BIZ_PGM_SAVE_ID   = "a2111mb3.asp"							'��: Biz logic For Update Row Data

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey            = 0										'�١١١�: Max key value
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim  lgIsOpenPop													'��: Popup status                           
Dim  lgKeyPosVal
Dim  IsOpenPop														'��: Popup status   
Dim  lgPageNo_A														'��: Next Key tag                          
Dim  lgSortKey_A													'��: Sort���� ���庯��                     
Dim  lgPageNo_B														'��: Next Key tag                          
Dim  lgSortKey_B													'��: Sort���� ���庯�� 
Dim  lgFncQuery

Dim  C_GL_CTRL_FLD 
Dim  C_GL_CTRL_NM  

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
    lgIntFlgMode     = parent.OPMD_CMODE                   'Indicates that current mode is Create mode

    lgPageNo_A       = ""                                  'initializes Previous Key for spreadsheet #1
    lgSortKey_A      = 1

    lgPageNo_B		 = ""                                  'initializes Previous Key for spreadsheet #2
    lgSortKey_B      = 1
End Sub

 '==========================================  2.2 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ������ ���� �ʵ� ���� ������ ������ 
'========================================================================================================= 
Sub  SetDefaultVal()

End Sub
'========================================  2.3 LoadInfTB19029()  =========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===========================================================================================================
Sub  LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("I", "A", "NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("I", "A", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	C_GL_CTRL_FLD = 1
	C_GL_CTRL_NM  = 2
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()

    With frm1.vspdData
		.MaxCols	= C_GL_CTRL_NM + 1
		.Col		= .MaxCols
		.ColHidden	= True

		.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		.ReDraw = False
		ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread

		Call GetSpreadColumnPos()

		ggoSpread.SSSetEdit	C_GL_CTRL_FLD ,"��ǥ�����׸�"  , 30,,,30,2
		ggoSpread.SSSetEdit	C_GL_CTRL_NM  ,"��ǥ�����׸��", 50

		Call ggoSpread.MakePairsColumn(C_GL_CTRL_FLD,C_GL_CTRL_NM,"1")
		
		.ReDraw = True
		Call SetSpreadLock_A()
    End With

    Call SetZAdoSpreadSheet("A2111MA1","S","A","V20021211",Parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X","X")
	Call SetSpreadLock_B()																		
End Sub

'=========================================================================================================
' Function Name : SetSpreadLock_A
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadLock_A()
	ggoSpread.SpreadLock C_GL_CTRL_FLD	, -1, C_GL_CTRL_FLD
	ggoSpread.SpreadLock C_GL_CTRL_NM	, -1, C_GL_CTRL_NM
End Sub

'=========================================================================================================
' Function Name : SetSpreadLock_B
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadLock_B()
	With frm1.vspdData2
		.ReDraw = False       
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()	
		.ReDraw = True
	End With 
End Sub

'=========================================================================================================
' Function Name : SetSpreadColor_A
' Function Desc : This method set color and protect in spread sheet celles
'=========================================================================================================
Sub  SetSpreadColor_A()
	Dim ii 
	
	With frm1.vspddata
		ggoSpread.Source = frm1.vspddata
		For ii = 1 To .MaxRows
			.row = ii
			.col = C_GL_CTRL_FLD
			If UCase(Left(Trim(.value),7)) = "USER_DF" Then
				ggoSpread.SpreadUnLock C_GL_CTRL_NM	, ii, C_GL_CTRL_NM ,ii
			Else	
				ggoSpread.SpreadLock   C_GL_CTRL_NM	, ii, C_GL_CTRL_NM ,ii
			End If
		Next
	End With					
End Sub

'=========================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'=========================================================================================================
Sub GetSpreadColumnPos()
    Dim iCurColumnPos

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_GL_CTRL_FLD = iCurColumnPos(1)
	C_GL_CTRL_NM  = iCurColumnPos(2)
End Sub


'**********************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'************************************************************************************** 

'------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenGlCtrlPopUp()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenGlCtrlPopUp(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function	
	
	lgIsOpenPop = True
	
	arrParam(0) = "��ǥ�����׸��˾�"								' �˾� ��Ī 
	arrParam(1) = "A_SUBLEDGER_CTRL " 									' TABLE ��Ī 
	arrParam(2) = Trim(strCode)											' Code Condition
	arrParam(3) = ""													' Name Condition
	arrParam(4) = ""													' Where Condition
	arrParam(5) = "��ǥ�����׸�"									' �����ʵ��� �� ��Ī 

	arrField(0) = "GL_CTRL_FLD"											' Field��(0)
	arrField(1) = "ISNULL(GL_CTRL_NM,'')"								' Field��(1)
	arrField(2) = ""													' Field��(2)
	arrField(3) = ""													' Field��(3)
			
	arrHeader(0) = "��ǥ�����׸�"									' Header��(0)
	arrHeader(1) = "��ǥ�����׸��"									' Header��(1)
	arrHeader(2) = ""													' Header��(2)
	arrHeader(3) = ""													' Header��(3)

	lgIsOpenPop = True
		
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then	    
		frm1.txtGlCtrlFld.Focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,1)
	End If
End Function
			
'=======================================================================================================
'	Name : SetBankAcct()
'	Description : Bank Account No Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		Case 1
			frm1.txtGlCtrlFld.value = arrRet(0)
			frm1.txtGlCtrlNm.value = arrRet(1)
			frm1.txtGlCtrlFld.focus				
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
	
	iGridPos = "B"
	
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
	Call InitVariables()																	'��: Initializes local global variables
	Call SetDefaultVal()
	Call InitSpreadSheet()
    Call SetToolbar("1100100000011111")														'��: ��ư ���� ���� 
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
	
	Call DbQuery("2",Row)
    
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
	
    lgPageNo_B       = ""                                  'initializes Previous Key
    lgSortKey_B      = 1
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    gMouseClickStatus = "SPC"	'Split �����ڵ�    
	Set gActiveSpdSheet = frm1.vspdData        

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
		    Exit Sub
	    End If
	    
		Call DbQuery("2",NewRow)
     
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
           Call DbQuery("1","")
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
           Call DbQuery("2","")
		End If
   End if
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub  vspdData_Change(ByVal Col, ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

' #########################################################################################################
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
    FncQuery = False															'��: Processing is NG
    Err.Clear     

	lgFncQuery = True
	
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then											'��: This function check indispensable field
		Exit Function
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
    Call DbQuery("1","")														'��: Query db data

    FncQuery = True		
	
	Set gActiveElement = document.activeElement
	lgFncQuery = False
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
Function  FncExit()
    FncExit = True
End Function

'=======================================================================================================
' Function Name : `
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1,var2
 
    FncSave = False                                                         
    
    On Error Resume Next
    Err.Clear                                                               
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange

    If var1 = False  Then											'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")				'��: Display Message(There is no changed data.)
		Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()													'��: Save db data
    FncSave = True  

    Set gActiveElement = document.activeElement
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function  DbSave() 
    Dim lngRows 
    Dim lGrpcnt
    DIM strVal 

    DbSave = False                                                          
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   
	Err.Clear 
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data ���� ��Ģ 
    ' 0: Sheet��, 1: Flag , 2: Row��ġ, 3~N: �� ����Ÿ 
    lGrpCnt = 1
    strVal = ""
    
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			If .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & parent.gColSep & lngRows & parent.gColSep
			    .Col = C_GL_CTRL_FLD '1
			    strVal = strVal & Trim(.Text) & parent.gColSep
			    .Col = C_GL_CTRL_NM  '2
			    strVal = strVal & Trim(.Text) & parent.gRowSep
			          
			    lGrpCnt = lGrpCnt + 1          
			End If
		Next
	End With
 
	frm1.txtSpread.value =  strVal								'Spread Sheet ������ ���� 
	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)					'���� �����Ͻ� ASP �� ���� 
    
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function  DbSaveOk()											'��: ���� ������ ���� ���� 
	ggoSpread.Source = frm1.vspdData        				
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2        				
								
	Call DBquery(1,"")
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 
'========================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'==========================================================================================================
Function  DbQuery(ByVal iOpt,ByVal Row) 
	Dim strVal
	Dim strCode
	Dim iRow
	
    Err.Clear																						'��: Protect system from crashing
	On Error Resume Next
	
	If Row = "" Then 
		iRow = frm1.vspddata.ActiveRow
	Else
		iRow = Row		
	End If	
	
    DbQuery = False
    Call DisableToolBar(parent.TBC_QUERY)															'��: Disable Query Button Of ToolBar
	Call LayerShowHide(1)
    
    With frm1
		Select Case iOpt 
			Case "1" 
				strVal = BIZ_PGM_ID & "?txtGlCtrlFld=" & Trim(.txtGlCtrlFld.value)
				strVal = strVal & "&txtGlCtrlFld_ALT=" & .txtGlCtrlFld.alt
			Case "2"
				.vspddata.row = iRow
				.vspddata.col = C_GL_CTRL_FLD
				strCode = .vspddata.value

				strVal = BIZ_PGM_ID1 & "?txtGlCtrlFld=" & strCode
				strVal = strVal & "&lgPageNo="        & lgPageNo									'��: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList="      & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="    & EnCoding(GetSQLSelectList("A"))
		End Select 
      
		Call RunMyBizASP(MyBizASP, strVal)															'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
End Function

'==================================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'==================================================================================================================
Function DbQueryOk(ByVal iOpt)																		'��: ��ȸ ������ ������� 
    lgIntFlgMode = parent.OPMD_UMODE																'��: Indicates that current mode is Update mode
    
	If iOpt = 1 Then

       Call vspdData_Click(1,1)
       frm1.vspdData.focus
	End If																							'��: This function lock the suitable field

	Call ggoOper.LockField(Document, "Q")															'��: This function lock the suitable field 
	Call SetSpreadColor_A()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǥ�����׸����</font></td>
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
									<TD CLASS="TD5" NOWRAP>��ǥ�����׸�</TD>
									<TD CLASS="TD6" COLSPAN=3 NOWRAP><INPUT TYPE=TEXT NAME="txtGlCtrlFld" SIZE=20 MAXLENGTH=20  STYLE="TEXT-ALIGN: Left" Tag="11XXXU" ALT="��ǥ�����׸�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGlCtrlFld" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenGlCtrlPopUp(frm1.txtGlCtrlFld.Value)">&nbsp;<INPUT TYPE=TEXT NAME="txtGlCtrlNm" SIZE=30 tag="14"></TD>
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
								<script language =javascript src='./js/a2111ma1_I548417818_vspdData.js'></script></TD>
							</TR>
							<TR HEIGHT="40%">
								<TD WIDTH="100%" colspan=6>
								<script language =javascript src='./js/a2111ma1_I369390183_vspdData2.js'></script></TD>
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
<TEXTAREA	CLASS=HIDDEN NAME=txtSpread	tag="24" TABINDEX="-1"></TEXTAREA>
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

