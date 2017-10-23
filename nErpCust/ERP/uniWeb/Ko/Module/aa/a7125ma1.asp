
<%@ LANGUAGE="VBSCRIPT"%>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7102ma1
'*  4. Program Name         : �����ڻ����󼼳������ 
'*  5. Program Desc         : �����ڻ꺰 ��� �� ������ ���,����,����,��ȸ 
'*  6. Comproxy List        : +As0021
'                             +As0029
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2001/05/19
'*  9. Modifier (First)     : ������ 
'* 10. Modifier (Last)      : ������ 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/003/30 : ..........
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################

'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* 
 -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
'==============================================  1.1.1 Style Sheet  ======================================
'=========================================================================================================
 -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js">			</SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit									'��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_ID = "a7125mb1.asp"			'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID_Q1 = "a7125mb2.asp"			'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID_Q2 = "a7125mb3.asp"			'��: �����Ͻ� ���� ASP�� 


'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================

'''�ڻ�master
Dim C_Deptcd
Dim C_DeptNm
Dim C_AcctCd
Dim C_AcctNm
Dim C_AsstNo
Dim C_AsstNm
Dim C_AcqAmt
Dim C_AcqLocAmt
Dim C_AcqQty
Dim C_ResAmt
Dim C_RefNo
Dim C_Desc

Const C_SHEETMAXROWS = 30

''���󼼳��� 
Dim C_Seq_2
Dim C_Desc_2
Dim C_Amt_2
Dim C_AsstNo_2
Dim C_LocAmt_2


Const C_SHEETMAXROWS_2  = 30	


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
Dim lgStrPrevKey_m

'Dim lgLngCurRows
'Dim lgKeyStream
Dim lgKeyStream_m

'Dim lgSortKey

'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  --------------------------------------------------------------
Dim IsOpenPop        
'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
 
 
 
'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'#########################################################################################################

'======================================================================================================
' Name : initSpreadPosVariables()
' Description : �׸���(��������) �÷� ���� ���� �ʱ�ȭ 
'=======================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)
	Select Case UCase(Trim(pvSpdNo))
		Case "A"	
			C_Deptcd		= 1
			C_DeptNm		= 2
			C_AcctCd		= 3
			C_AcctNm		= 4
			C_AsstNo		= 5
			C_AsstNm		= 6
			C_AcqAmt		= 7
			C_AcqLocAmt	= 8
			C_AcqQty		= 9
			C_ResAmt		= 10
			C_RefNo		= 11
			C_Desc		= 12
		Case "B"			
			C_Seq_2			= 1
			C_Desc_2		= 2
			C_Amt_2			= 3
			C_AsstNo_2		= 4
			C_LocAmt_2		= 5
	End Select
End Sub


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
	
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
frm1.txtAcqNo.focus
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey_m = 0                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    
End Sub

'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'*********************************************************************************************************
'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>

End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
   Select Case pOpt
       Case "Q"
                  lgKeyStream = Frm1.txtAcqNo.Value  & Parent.gColSep       'You Must append one character(Parent.gColSep)
       Case "M"
                  lgKeyStream = Frm1.htxtAcqNo.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
   End Select 
                   
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub  InitSpreadSheet(ByVal pvSpdNo)
    Call initSpreadPosVariables(pvSpdNo)
    
    Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

				.ReDraw = false
				.MaxCols = C_Desc +1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
				.Col = .MaxCols								    '��: ������Ʈ�� ��� Hidden Column
				.ColHidden = True
				.MaxRows = 0
	
				Call GetSpreadColumnPos(pvSpdNo)

   				'Col, Header, ColWidth, HAlign, FloatMax, FloatMin, FloatSeparator, FloatSepChar, FloatDecimalPlaces, FloatDeciamlChar
				ggoSpread.SSSetEdit		C_DeptCd,  "�μ��ڵ�", 8, , , 10
				ggoSpread.SSSetEdit		C_DeptNm,  "�μ���",   10

				ggoSpread.SSSetEdit		C_AcctCd,  "�����ڵ�", 10, , , 20
				ggoSpread.SSSetEdit		C_AcctNm,  "������",   20
				ggoSpread.SSSetEdit		C_AsstNo, "�ڻ��ȣ", 15, , , 18
			    ggoSpread.SSSetEdit		C_AsstNm, "�ڻ��",   20, , , 40
			    
				ggoSpread.SSSetFloat    C_AcqAmt,   "���ݾ�",      15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetFloat    C_AcqLocAmt,"���ݾ�(�ڱ�)",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				Call AppendNumberPlace("6","3","0")

			    ggoSpread.SSSetFloat    C_AcqQty,   "������",      15,"6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetFloat    C_ResAmt,"��������(�ڱ�)",15, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetEdit		C_RefNo, "������ȣ", 30, , , 30
				ggoSpread.SSSetEdit		C_Desc,  "����",     30, , , 128

				.ReDraw = true

			End With
		Case "B"
		
			With frm1.vspdData2
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 
				
				.ReDraw = false
				.MaxCols = C_LocAmt_2 +1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
				.Col = .MaxCols								    '��: ������Ʈ�� ��� Hidden Column
				.ColHidden = True
				.MaxRows = 0

				Call GetSpreadColumnPos(pvSpdNo)

				'Col, Header, ColWidth, HAlign, FloatMax, FloatMin, FloatSeparator, FloatSepChar, FloatDecimalPlaces, FloatDeciamlChar
				ggoSpread.SSSetFloat	C_Seq_2,     "����", 14, "6", ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,1,True  ,, "1","999"
				ggoSpread.SSSetEdit		C_Desc_2,  "����"		,53, , , 40
				ggoSpread.SSSetFloat    C_Amt_2,   "�ݾ�"		,25, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
				ggoSpread.SSSetEdit		C_AsstNo_2, "�ڻ��ȣ", 15, , , 18			'Asset_no�� Hidden���� ������ ����.
				ggoSpread.SSSetFloat    C_LocAmt_2,"�ݾ�(�ڱ�)"	,25, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		        Call ggoSpread.SSSetColHidden(C_AsstNo_2,C_AsstNo_2,True)

				.ReDraw = true
					
			End With

	End Select
	
    Call SetSpreadLock(pvSpdNo)	
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

	Select Case UCase(Trim(pvSpdNo))
		Case "A"
			With frm1.vspdData
	
				ggoSpread.Source = frm1.vspdData

				.ReDraw = False		

					ggoSpread.SpreadLock C_DeptCd,   -1
					ggoSpread.SpreadLock C_DeptNm,   -1
					ggoSpread.SpreadLock C_AcctCd,   -1
					ggoSpread.SpreadLock C_AcctNm,   -1
					ggoSpread.SpreadLock C_AcqAmt,   -1
					ggoSpread.SpreadLock C_AcqLocAmt,   -1
					ggoSpread.SpreadLock C_AcqQty,   -1
					ggoSpread.SpreadLock C_ResAmt,   -1
					ggoSpread.SpreadLock C_RefNo,   -1
					ggoSpread.SpreadLock C_Desc,   -1
						
				.ReDraw = True

			End With    
		Case "B"	
			With frm1.vspdData2
	
				ggoSpread.Source = frm1.vspdData2
						
				.ReDraw = False		
				
					ggoSpread.SpreadLock C_Seq_2,   -1
					ggoSpread.SpreadUnLock C_Desc_2,   -1
					ggoSpread.SpreadUnLock C_Amt_2,   -1
					ggoSpread.SpreadUnLock C_LocAmt_2,   -1
					
					ggoSpread.SSSetProtected C_LocAmt_2 +1, -1,C_LocAmt_2 +1
				
				.ReDraw = True

			End With    
		End Select

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow , ByVal pvEndRow)
	With frm1.vspdData2
		.Redraw = False
		ggoSpread.Source = frm1.vspdData2			
		ggoSpread.SSSetRequired C_Seq_2, pvStartRow, pvEndRow
'		.Col = 2											'�÷��� ���� ��ġ�� �̵� 
'		.Row = .ActiveRow
'		.Action = 0                         
'		.EditMode = True		
		.Redraw = True		
    End With		
End Sub


'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : This method call saved columnorder
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)		

			C_Deptcd	= iCurColumnPos(1)
			C_DeptNm	= iCurColumnPos(2)
			C_AcctCd	= iCurColumnPos(3)
			C_AcctNm	= iCurColumnPos(4)
			C_AsstNo	= iCurColumnPos(5)
			C_AsstNm	= iCurColumnPos(6)
			C_AcqAmt	= iCurColumnPos(7)
			C_AcqLocAmt	= iCurColumnPos(8)
			C_AcqQty	= iCurColumnPos(9)
			C_ResAmt	= iCurColumnPos(10)
			C_RefNo		= iCurColumnPos(11)
			C_Desc		= iCurColumnPos(12)

		Case "B"
			ggoSpread.Source = frm1.vspdData2

			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)							
			
			C_Seq_2		= iCurColumnPos(1)
			C_Desc_2	= iCurColumnPos(2)
			C_Amt_2		= iCurColumnPos(3)
			C_AsstNo_2	= iCurColumnPos(4)
			C_LocAmt_2	= iCurColumnPos(5)

	End select
End Sub



'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

 '========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
 '++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++++
'======================================================================================================
'   Function Name : OpenAcqNoInfo()
'   Function Desc : 
'=======================================================================================================
Function OpenAcqNoInfo()
	Dim arrRet
	Dim arrParam(3)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a7102ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7102ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True	
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent,arrParam), _
		     "dialogWidth=800px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	

	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetAcqNoInfo(arrRet)
	End If	

End Function

'======================================================================================================
'   Function Name : SetAcqNoInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetAcqNoInfo(Byval arrRet)

	With frm1
		.txtAcqNo.value  = arrRet(0)
		
		.txtAcqNo.focus
	End With

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

    Call LoadInfTB19029()                                                         'Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field

    Call InitSpreadSheet("A")                                                     'Setup the Spread sheet

    Call InitSpreadSheet("B")                                                     'Setup the Spread sheet
    
    Call InitVariables                                                      '��: Initializes local global variables
    
    Call SetDefaultVal    
	Call SetToolbar("1100000000000111")        
	
	frm1.txtAcqNo.focus
	
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

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
	
	Dim IntRetCD 
    Dim var_i, var_m
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData2
    var_i = ggoSpread.SSCheckChange

    ggoSpread.Source = frm1.vspdData
    var_m = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var_i = True or var_m = True    Then    
		IntRetCD = DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X") '�� �ٲ�κ� 
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    ggoSpread.Source = frm1.vspdData2
	ggospread.ClearSpreadData		'Buffer Clear
        
    Call InitVariables															'��: Initializes local global variables
'    Call InitSpreadSheet("A")                                                     'Setup the Spread sheet
'    Call InitSpreadSheet("B")                                                     'Setup the Spread sheet
    
    '-----------------------
    'Check condition area
    '-----------------------

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery("Q") = False Then                                                       '��: Query db data
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                               '��: Processing is OK
	   
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    Dim var_i
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    On Error Resume Next                                                    '��: Protect system from crashing
    
'	if frm1.vspdData2.MaxRows < 1 then'
'		IntRetCD = DisplayMsgBox("900001","X","X","X")  ''�ڻ꼼�γ����� �Է��Ͻʽÿ�.
'		Exit Function
'	end if

		
    ggoSpread.Source = frm1.vspdData2
    var_i = ggoSpread.SSCheckChange
	
    If lgBlnFlgChgValue = False and var_i = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")  '�� �ٲ�κ� 
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------

    ggoSpread.Source = frm1.vspdData2
    If Not ggoSpread.SSDefaultCheck Then	
		Exit Function
    End if
	    
    Call DbSave()				                                                
    
    FncSave = True                                                          
    
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
    Dim IntRetCD
    
	frm1.vspdData2.ReDraw = False

	if frm1.vspdData2.MaxRows < 1 then Exit Function
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.CopyRow

	SetSpreadColor frm1.vspdData2.ActiveRow , frm1.vspdData2.ActiveRow
    
	frm1.vspdData2.Col  = C_Seq_2
	frm1.vspdData2.Text = ""
		
	frm1.vspdData2.ReDraw = True
	
End Function


'========================================================================================
' Function Name : FncPaste
' Function Desc : This function is related to Paste Button of Main ToolBar
'========================================================================================
Function FncPaste() 
     ggoSpread.SpreadPaste
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    if frm1.vspdData2.MaxRows < 1 then	 Exit Function

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.EditUndo                                                  '��: Protect system from crashing

End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(Byval pvRowCnt) 
	Dim varMaxRow
	Dim strDoc
	Dim varXrate
	Dim imRow
	
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

		If ImRow="" then
			Exit Function
		End If
	End If
		
	with frm1
		varMaxRow = .vspdData2.MaxRows 

		.vspdData2.focus
		
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False
		
		ggoSpread.InsertRow ,imRow

		
		frm1.vspdData.row = frm1.vspdData.activeRow
		frm1.vspdData.Col = C_asstNo

		.vspdData2.row = .vspdData2.ActiveRow
		.vspdData2.Col = C_asstNo_2
		.vspdData2.value = frm1.vspdData.value
		
		.vspdData2.Col = C_Amt_2
		.vspdData2.value = 0
		.vspdData2.Col = C_LocAmt_2
		.vspdData2.value = 0
		
		.vspdData2.ReDraw = True

		SetSpreadColor .vspdData2.ActiveRow , frm1.vspdData2.ActiveRow

	end with
	
'	Call SetToolbar("1100111100111111")

End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows 
    Dim lTempRows 

	frm1.vspdData2.focus
   	ggoSpread.Source = frm1.vspdData2

	if frm1.vspdData2.MaxRows < 1 then Exit Function
	
	lDelRows = ggoSpread.DeleteRow    

End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Parent.fncPrint()    
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                               
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
	Dim IntRetCD
	FncExit = False	
		
	If lgBlnFlgChgValue = True then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")   '�� �ٲ�κ� 

		If IntRetCD = vbNo Then		
			Exit Function
		End If

    End If
    
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()												'��: ���� ������ ���� ���� 
	Call Detail_Sum
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery(pDirect)

	Dim strVal
	
    Err.Clear                                                                    '��: Clear err status
    On Error Resume Next
    
    frm1.txtpDirect.value = pDirect
    
    DbQuery = False                                                              '��: Processing is NG

'    Call DisableToolBar(TBC_QUERY)                                               '��: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '��: Show Processing Message

    Call MakeKeyStream(pDirect)

    Select Case pDirect
       Case "M" 
          With Frm1
                 strVal = BIZ_PGM_ID_Q1  & "?txtMode="         & parent.UID_M0001						         
                 strVal = strVal      & "&txtKeyStream="    & lgKeyStream           '��: Query Key
                 strVal = strVal      & "&txtMaxRows="      & .vspdData.MaxRows
                 strVal = strVal      & "&lgStrPrevKey="    & lgStrPrevKey          '��: Next key tag
          End With
       Case "Q"
                 strVal = BIZ_PGM_ID_Q1 & "?txtMode="          & parent.UID_M0001            '��: Query
                 strVal = strVal      & "&txtKeyStream="     & lgKeyStream          '��: Query Key
    End Select    
    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                  '��:  Run biz logic

    DbQuery = True                                                      '��: Processing is OK

    Set gActiveElement = document.ActiveElement   
        
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()													'��: ��ȸ ������ �������	
		
    lgIntFlgMode =  parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    
   ' Call ggoOper.LockField(Document, "Q")										'��: This function lock the suitable field    	
	Call SetToolbar("1100111100111111")	
	
	lgBlnFlgChgValue = False

	IF frm1.txtpDirect.value = "M" Then
		Exit Function
	End IF
	
	Call dbquery2(1,1,"Q")
	
End Function

 '========================================================================================
'    Function Name : InitData()
'    Function Desc : 
'   ======================================================================================== 
Sub InitData()

End Sub

'========================================================================================
' Function Name : DbQuery2
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery2(pRow, pCol,pDirect) 

	Dim strVal
	Dim IntRetCD
	
    Err.Clear                                                                    '��: Clear err status
    On Error Resume Next

    DbQuery2 = False                                                         '��: Processing is NG

	If pDirect = "Q" Then 

		ggoSpread.Source = frm1.vspdData2
    
		If ggoSpread.SSCheckChange = True Then    
			IntRetCD = DisplayMsgBox("990027", "X","X","X") '�� �ٲ�κ� 
			frm1.vspdData.row = frm1.txtActiveRows.value
			frm1.vspdData.Col = frm1.txtActiveCols.value
			frm1.vspdData.action = 0
			Exit Function
		End If    
		frm1.vspdData2.MaxRows = 0
	End IF
	        
    Call DisableToolBar(Parent.TBC_QUERY)                                               '��: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '��: Show Processing Message

	With frm1.vspdData
		.Row = pRow
		.Col = C_AsstNo
		frm1.txtKeyStream_m.value = .value
        lgKeyStream_m = .Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
    End with
    
    
    Select Case pDirect
       Case "M" 
          With Frm1
                 strVal = BIZ_PGM_ID_Q2  & "?txtMode="         & Parent.UID_M0001						         
                 strVal = strVal      & "&txtKeyStream_m="    & lgKeyStream_m           '��: Query Key
                 strVal = strVal      & "&txtMaxRows_2="      & .vspdData2.MaxRows
                 strVal = strVal      & "&lgStrPrevKey_m="    & lgStrPrevKey_m          '��: Next key tag
          End With
       Case "Q"
                 strVal = BIZ_PGM_ID_Q2 & "?txtMode="          & Parent.UID_M0001            '��: Query
                 strVal = strVal      & "&txtKeyStream_m="     & lgKeyStream_m          '��: Query Key
    End Select    

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
    Call RunMyBizASP(MyBizASP, strVal)                                  '��:  Run biz logic

    DbQuery2 = True                                                      '��: Processing is OK

    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk2()													'��: ��ȸ ������ ������� 

    lgIntFlgMode = Parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
    
	Call SetToolbar("1100111100111111")	
	
	Call Detail_Sum
	
	With frm1.vspdData
		frm1.txtActiveRows.value = .activerow
		.Row = frm1.txtActiveRows.value
'		.Col = C_Next
'		.value = lgStrPrevKey_m
	End with
    
End Function


Function Detail_Sum()
	Dim i
	Dim Sum
	Dim LocSum

	Sum = 0 
	LocSum = 0
	
	With frm1.vspdData2
		for i = 1 to .Maxrows
			.row = i
			.col = C_Amt_2
			
			Sum = UNICDbl(Sum) + UNICDbl(.text)
			.Col = C_LocAmt_2
			
			LocSum = UNICDbl(LocSum) + UNICDbl(.text)
		Next
	End With
	frm1.txtSum.text  = UNIFormatNumber(Sum, Parent.ggAmtOfMoney.DecPoint, -2, 0, Parent.ggAmtOfMoney.RndPolicy, Parent.ggAmtOfMoney.RndUnit)
	frm1.txtLocSum.text  = UNIFormatNumber(LocSum, Parent.ggAmtOfMoney.DecPoint, -2, 0, Parent.ggAmtOfMoney.RndPolicy, Parent.ggAmtOfMoney.RndUnit)
'	frm1.txtSum.text  = Sum
'	frm1.txtLocSum.text = LocSum

End Function


'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim IntRows 
    Dim lGrpcnt 
	Dim strVal, strDel
	Dim IntLocAmt
	
    DbSave = False                                                          '��: Processing is NG    

	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value    = Parent.UID_M0002										'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ����			
	End With
	
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data ���� ��Ģ 
    ' 0: Sheet��, 1: Flag , 2: Row��ġ, 3~N: �� ����Ÿ 
	frm1.txtSum.text  = "0"
	frm1.txtLocSum.text  = "0"

    lGrpCnt = 1    
	strVal = ""
	strDel = ""
    
    '-----------------------------
    '   Acq item Part
    '-----------------------------
    With frm1.vspdData2
	    
    For IntRows = 1 To .MaxRows
    	
		.Row = IntRows
		.Col = 0		
		
		Select Case .Text		    
		        
		    Case ggoSpread.DeleteFlag

		        strDel = strDel & "D" & Parent.gColSep & IntRows & Parent.gColSep

				.Col = C_AsstNo_2
				strDel = strDel & Trim(.value) & Parent.gColSep
					
				.Col = C_Seq_2
				strDel = strDel & Trim(.Text) & Parent.gRowSep				    '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 
					
				lGrpcnt = lGrpcnt + 1            
		    
		    Case ggoSpread.UpdateFlag

				strVal = strVal & "U" & Parent.gColSep & IntRows & Parent.gColSep

				.Col = C_AsstNo_2
				strVal = strVal & Trim(.value) & Parent.gColSep
					
				.Col = C_Seq_2
				strVal = strVal & Trim(.Text) & Parent.gColSep
					
				.Col = C_Desc_2
				strVal = strVal & Trim(.value) & Parent.gColSep
					
				.Col = C_Amt_2
				strVal = strVal & UNIConvNum(Trim(.Text),0)  & Parent.gColSep
				
				.Col = C_LocAmt_2
				strVal = strVal & UNIConvNum(Trim(.Text),0)  & Parent.gRowSep				    '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 

		        lGrpCnt = lGrpCnt + 1
		        
		    Case ggoSpread.InsertFlag

				strVal = strVal & "C" & Parent.gColSep & IntRows & Parent.gColSep

				.Col = C_AsstNo_2
				strVal = strVal & Trim(.value) & Parent.gColSep
					
				.Col = C_Seq_2
				strVal = strVal & Trim(.Text) & Parent.gColSep
					
				.Col = C_Desc_2
				strVal = strVal & Trim(.value) & Parent.gColSep
					
				.Col = C_Amt_2
				strVal = strVal & UNIConvNum(Trim(.Text),0) & Parent.gColSep
					
				.Col = C_LocAmt_2
				strVal = strVal & UNIConvNum(Trim(.Text),0) & Parent.gRowSep				    '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 

		        lGrpCnt = lGrpCnt + 1

		End Select

    Next

	End With
	
	frm1.txtMaxRows_2.value  = lGrpCnt-1										'��: Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread_m.value = strDel & strVal									'��: Spread Sheet ������ ���� 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'��: ���� �����Ͻ� ASP �� ���� 

    DbSave = True                                                           ' ��: Processing is OK
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
   ' Call InitVariables	
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    'lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey_m = 0                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    '-----------------------
    'Erase contents area
    '-----------------------
    frm1.vspdData2.MaxRows = 0
	call dbquery2(frm1.txtActiveRows.value,1,"Q")
	
End Function


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'===================================== PopSaveSpreadColumnInf()  ======================================
' Name : PopSaveSpreadColumnInf()
' Description : �̵��� �÷��� ������ ���� 
'====================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'===================================== PopRestoreSpreadColumnInf()  ======================================
' Name : PopRestoreSpreadColumnInf()
' Description : �÷��� ���������� ������ 
'====================================================================================================
Sub  PopRestoreSpreadColumnInf()
	Dim indx

	Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA" 
			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()

		Case "VSPDDATA2"
			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")
			Call ggoSpread.ReOrderingSpreadData()
	End Select
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================

Sub vspdData_Change(Col , Row)


End Sub


'========================================================================================================
'   Event Name : vspdData2_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_Change(Col , Row)

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row
	
	If Col = C_Amt_2 or Col = C_LocAmt_2 Then
		Call Detail_Sum
	End If
    
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(Col, Row)
	If  lgIntFlgMode =  parent.OPMD_UMODE Then
		Call SetPopUpMenuItemInf("1111111111")
    Else
		Call SetPopUpMenuItemInf("0000111111")    
    End if
    
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey
            lgSortKey = 1
        End If
    End If

    
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData2_Click(Col, Row)
	If  lgIntFlgMode =  parent.OPMD_UMODE Then
		Call SetPopUpMenuItemInf("1111111111")
    Else
		Call SetPopUpMenuItemInf("0000111111")    
    End if

    gMouseClickStatus = "SP2C"
    Set gActiveSpdSheet = frm1.vspdData2
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col ,lgSortKey
            lgSortKey = 1
        End If
    End If

    
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) �÷� width ���� �̺�Ʈ �ڵ鷯 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

Sub vspdData2_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) �÷� width ���� �̺�Ʈ �ڵ鷯 
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Ư�� column�� click�Ҷ� 
'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'======================================================================================================
'   Event Name : vspdData2_MouseDown
'   Event Desc : Ư�� column�� click�Ҷ� 
'======================================================================================================
Sub vspdData2_MouseDown(Button , Shift , x , y)

	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub


Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData2_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("B")
End Sub



'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================


Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    With frm1.vspdData
	
		If col < 1 or Row < 1 or NewCol < 1 or NewRow < 1 Then
			Exit Sub
		End IF
				
		If Row = NewRow Then
		    Exit Sub
		End If
		 
			frm1.txtActiveRows.value = NewCol
			frm1.txtActiveCols.value = NewRow

		Call Dbquery2(NewRow, NewCol, "Q")
		
    End With
    
End Sub

Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
End Sub

Sub  vspdData2_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    With frm1.vspdData2
	
		If Newcol < 1 or NewRow < 1 Then
			frm1.txtActiveRows_m.value = NewRow
			Exit Sub
		End If
				
		If Row = NewRow Then
		    Exit Sub
		End If
		
    End With
    
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If



    if frm1.vspdData.MaxRows < (NewTop + VisibleRowCnt(frm1.vspdData,NewTop)) Then	
    	If lgStrPrevKey <> "" Then  
           If DbQuery("M") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    if frm1.vspdData2.MaxRows < (NewTop + VisibleRowCnt(frm1.vspdData2,NewTop)) Then	 
    	If lgStrPrevKey_m <> 0 Then   
           If DbQuery2(frm1.txtActiveRows.value,frm1.txtActiveCols.value, "M") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!--'#########################################################################################################
'       					6. Tag�� 
'######################################################################################################### -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����ڻ����泻�����</font></td>
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
									<TD CLASS="TD5" NOWRAP>����ȣ</TD>
									<TD CLASS="TD6"><INPUT NAME="txtAcqNo" TYPE="Text" MAXLENGTH=18 tag="12XXXU" ALT="����ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo1" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenAcqNoInfo"></TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%></TD>
				</TR>
				<TR HEIGHT=100%>
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" TYPE="Text" MAXLENGTH=3 SIZE=10 tag="24XXXU" ></TD>
							    <TD CLASS="TD5" NOWRAP>�������</TD>																							    
								<TD CLASS="TD6" NOWRAP>
									<script language =javascript src='./js/a7125ma1_fpDateTime1_txtAcqDt.js'></script>											    
								</TD>
							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBpCd" ALT="�ŷ�ó" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="24XXXU">
													<INPUT NAME="txtBpNm" TYPE="Text" SIZE = 22 tag="24">
								</TD>
								<TD CLASS=TD5 NOWRAP>ȯ��</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/a7125ma1_fpDoubleSingle1_txtXchRate.js'></script>
	                            </TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=4></TD>
							</TR>
							
							<TR>
								<TD WIDTH="100%" HEIGHT=45% COLSPAN=4>
									<script language =javascript src='./js/a7125ma1_vspdData_vspdData.js'></script>
								</TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=4></TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT=40% COLSPAN=4>
									<script language =javascript src='./js/a7125ma1_vspdData2_vspdData2.js'></script>
								</TD>
							</TR>
							<TR>
								<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=4></TD>
							</TR>
							<TR>
								<TD HEIGHT=20 WIDTH=100% COLSPAN=4>
									<FIELDSET CLASS="CLSFLD">
										<TABLE <%=LR_SPACE_TYPE_20%>>
											<TR>
												<TD CLASS="TD5" NOWRAP>�󼼳����հ�</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/a7125ma1_fpDoubleSingle2_txtSum.js'></script>&nbsp;
												</TD>
												<TD CLASS="TD5" NOWRAP>�󼼳����հ�(�ڱ�)</TD>
												<TD CLASS=TD6 NOWRAP>
													<script language =javascript src='./js/a7125ma1_fpDoubleSingle3_txtlocSum.js'></script>&nbsp;
												</TD>
											</TR>
										</TABLE>
									</FIELDSET>
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
	<TR HEIGHT=10>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="htxtAcqNo"    tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMode"      tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtActiveRows" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtActiveCols" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtActiveRows_m" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows_2" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows_3" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream_m"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtSpread_m"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"   tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtpDirect"   tag="24" TABINDEX = "-1" >

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

