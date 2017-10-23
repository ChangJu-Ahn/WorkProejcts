<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7105ma1
'*  4. Program Name         : �����ڻ� �μ�������� ��� 
'*  5. Program Desc         : �����ڻ��� �μ���������� ���/����/����/��ȸ�Ѵ�.
'*  6. Comproxy List        : +As0061ManageSvr
'*                            +As0068ListSvr
'*                            +B1a028ListMinorCode
'*  7. Modified date(First) : 2000/09/18
'*  8. Modified date(Last)  : 2001/05/31
'*  9. Modifier (First)     : hersheys
'* 10. Modifier (Last)      : Kim Hee Jung
'* 11. Comment              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit									'��: indicates that All variables must be declared in advance


'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_ID = "a7105mb2.asp"			'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 = "a7105mb3.asp"			'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================


Dim C_ASSETNO
Dim C_ASSETNOPopUp
Dim C_ASSETNONM

Dim C_DeptCd
Dim C_DeptCdPopUp
Dim C_DeptNm
Dim	C_OrgChangeId
Dim C_CostCd
Dim C_CostNm
Dim C_CostType
Dim C_CostTypeNm
Dim C_InvQty
Dim C_AssnRate




'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
'Dim lgLngCurRows


'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

'-------------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim lgBlnStartFlag				' �޼��� �����Ͽ� ���α׷� ���۽��� Check Flag
Dim IsOpenPop        
Dim lgRetFlag

'+++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

' ���Ѱ��� �߰�
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' �����
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ����
Dim lgStrPrevToKey 
Dim lgMasterFg 

'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

<!-- #Include file="../../inc/lgvariables.inc" -->	


Sub initSpreadPosVariables()


	C_ASSETNO      = 1 
	C_ASSETNOPopUp = 2 
	C_ASSETNONM    = 3 
	C_DeptCd       = 4 
	C_DeptCdPopUp  = 5 
	C_DeptNm       = 6 
	C_OrgChangeId  = 7 
	C_CostCd	   = 8 
	C_CostNm	   = 9 
	C_CostType     = 10
	C_CostTypeNm   = 11
	C_InvQty	   = 12
	C_AssnRate     = 13
	
	
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    lgStrPrevToKey = 1
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0       'initializes Deleted Rows Count
    lgSortKey = 1                     
    lgPageNo     = "0"
    
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

<%
    Dim svrDate
    svrDate = GetSvrDate
%>
	frm1.htxtCurrentDt.value = UniConvDateAToB("<%=svrDate%>", parent.gServerDateFormat,parent.gDateFormat)

	frm1.txtYyyymm.Text = UNIConvDateAToB("<%=svrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)'Convert DB date type to Company

    ' �μ������Master �� �μ������History �� � ���� ��������� ���� ���� ����
	Dim IntRet, iWhere

    iWhere = " MAJOR_CD = 'A1001' AND MINOR_CD = 'DP' AND SEQ_NO = 2 "
	IntRet = CommonQueryRs(" REFERENCE ","B_CONFIGURATION", iWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If intRet <> False Then
		If Replace(lgF0,Chr(11),"") = "M" Then
			lgMasterFg = "M"
		Else
			lgMasterFg = "H"
		End If
	End If
   
	call Set1()
	
End Sub

Sub Set1()
	If lgMasterFg = "H" Then    
	    Call SetToolbar("11101111001111")										'��: ��ư ���� ���� 
	     frm1.button1(0).disabled = false
	     frm1.button1(1).disabled = false
	Else
	    Call SetToolbar("11000000000011")										'��: ��ư ���� ���� 
	    frm1.button1(0).disabled = true
	    frm1.button1(1).disabled = true
    End If
end Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
    	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'====================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================

Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()

	With frm1.vspdData

	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.Spreadinit "V20071103",,parent.gAllowDragDropSpread  
	    .ReDraw = false
	    
		.MaxCols = C_AssnRate+1							'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	    .Col = .MaxCols								    '��: ������Ʈ�� ��� Hidden Column
	    .ColHidden = True
	    
	'   	.Col = C_CostCd								'��: ����� �� Hidden Column
	'    .ColHidden = True
	'   	.Col = C_CostType								'��: ����� �� Hidden Column
	'    .ColHidden = True

	    '.MaxRows = 0
	    ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit     C_ASSETNO,     "�ڻ��ȣ",		  10, , , 18,2 '3
	    ggoSpread.SSSetButton   C_ASSETNOPopUp									'4
	    ggoSpread.SSSetEdit		C_ASSETNONM,     "�ڻ��",           20			'5

	    ggoSpread.SSSetEdit     C_DeptCd,     "�����μ�",		  10, , , 10,2 '3
	    ggoSpread.SSSetButton   C_DeptCdPopUp									'4
	    ggoSpread.SSSetEdit		C_DeptNm,     "�μ���",           16			'5
	    ggoSpread.SSSetEdit		C_OrgChangeId,"���������ڵ�",     10			'5
	    ggoSpread.SSSetEdit     C_CostCd,	  "�ڽ�Ʈ��Ÿ",		  10			'6
	    ggoSpread.SSSetEdit     C_CostNm,	  "�ڽ�Ʈ��Ÿ��",     21			'7
	    ggoSpread.SSSetEdit    C_CostType,   "",                 10			'8
	    ggoSpread.SSSetEdit    C_CostTypeNm, "����������",        15			'9
	    ggoSpread.SSSetFloat	C_InvQty,       "������",       10, parent.ggQtyNo,       ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		'Col, Header, ColWidth, HAlign, FloatMax, FloatMin, FloatSeparator, FloatSepChar, FloatDecimalPlaces, FloatDeciamlChar
	   	ggoSpread.SSSetFloat	C_AssnRate,   "��к���(%)", 15, parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z","0","100"
		
		ggoSpread.SpreadLock C_AssnRate+1, -1, C_AssnRate+1
	   	
	   	Call ggoSpread.MakePairsColumn(C_DeptCd,C_DeptCdPopUp,"1")
	   	
	    	 Call ggoSpread.SSSetColHidden(C_CostCd,C_CostCd,True)
	     Call ggoSpread.SSSetColHidden(C_OrgChangeId,C_OrgChangeId,True)
	     Call ggoSpread.SSSetColHidden(C_CostType,C_CostType,True)
	     Call ggoSpread.SSSetColHidden(C_InvQty,C_InvQty,True)	   	
	    
		.ReDraw = true
				
	    Call SetSpreadLock 
    
    End With
    
End Sub


'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'===========================================================================================================

Sub SetSpreadLock()

    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock C_DeptCd, -1, C_CostTypeNm , C_InvQty
		ggoSpread.SpreadLock C_AssnRate+1, -1, C_AssnRate+1

		.vspdData.ReDraw = True
    End With
    
End Sub


'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'============================================================================================================

Sub SetSpreadColor(ByVal query_fg, ByVal pvStarRow, ByVal pvEndRow)
	
    With frm1
		.vspdData.ReDraw = False
		if UCase(query_fg) = "Q" then
			ggoSpread.SSSetProtected  C_ASSETNO,     pvStarRow, pvEndRow
			ggoSpread.SSSetProtected  C_ASSETNOPopUp,     pvStarRow, pvEndRow
			ggoSpread.SSSetProtected  C_ASSETNONM,     pvStarRow, pvEndRow
			
			ggoSpread.SSSetProtected  C_DeptCd,     pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_DeptNm,     pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_CostType,   pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_CostNm,  pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_CostTypeNm, pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_InvQty, pvStarRow, pvEndRow
			ggoSpread.SSSetRequired  C_AssnRate,   pvStarRow, pvEndRow					
		else	
			ggoSpread.SSSetRequired  C_ASSETNO,     pvStarRow, pvEndRow
			'ggoSpread.SSSetProtected C_ASSETNOPopUp,     pvStarRow, pvEndRow
			ggoSpread.SSSetProtected  C_ASSETNONM,     pvStarRow, pvEndRow
				
			ggoSpread.SSSetRequired  C_DeptCd,     pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_DeptNm,     pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_CostType,   pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_CostNm,  pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_CostTypeNm, pvStarRow, pvEndRow
			ggoSpread.SSSetProtected C_InvQty, pvStarRow, pvEndRow
			ggoSpread.SSSetRequired  C_AssnRate,   pvStarRow, pvEndRow		
		end if	
		.vspdData.ReDraw = True
    End With

End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_ASSETNO      = iCurColumnPos(1 )
				C_ASSETNOPopUp = iCurColumnPos(2 )
				C_ASSETNONM    = iCurColumnPos(3 )
				C_DeptCd       = iCurColumnPos(4 )
				C_DeptCdPopUp  = iCurColumnPos(5 )
				C_DeptNm       = iCurColumnPos(6 )
				C_OrgChangeId  = iCurColumnPos(7 )
				C_CostCd	   = iCurColumnPos(8 )
				C_CostNm	   = iCurColumnPos(9 )
				C_CostType     = iCurColumnPos(10)
				C_CostTypeNm   = iCurColumnPos(11)
				C_InvQty	   = iCurColumnPos(12)
				C_AssnRate     = iCurColumnPos(13)
	End Select
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 

Sub InitComboBox()

	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Dim IntRetCD1

	IntRetCD1= CommonQueryRs("MINOR_CD,MINOR_NM","B_MINOR","(MAJOR_CD = " & FilterVar("B9013", "''", "S") & " )",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
	If intRetCD1 <> False Then
		ggoSpread.Source = frm1.vspddata
		ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_CostType		
		ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_CostTypeNm
	End If		

	'------ Developer Coding part (End )   --------------------------------------------------------------

end sub

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
 '------------------------------------------  OpenMasterRef()  -------------------------------------------------
'	Name : OpenMasterRef()
'	Description : Asset Master Condition PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMasterRef(pVal,Row)

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	' ���Ѱ��� �߰�
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True
	
	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	If arrRet(0) = "" Then
		IsOpenPop = False
		Exit Function
	Else
		if pVal="H" then 'header
			Call SetPoRef(arrRet)
			
		else 'detail
			
			Call SetPoRefd(arrRet,Row)
			
		end if
	End If	

	IsOpenPop = False

	frm1.txtCondAsstNo.focus
	
End Function

 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(strRet)
       
	frm1.txtCondAsstNo.value     = strRet(0)
	frm1.txtcondAsstNm.value	 = strRet(1)
		
End Sub

 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRefD()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 

Sub SetPoRefD(strRet,Row)

    with frm1.vspdData
		.Row = Row
        .Col = C_Assetno : .Value =  strRet(0)
        .Col = C_Assetno+2 : .Value =  strRet(1)

		
	end With
		
End Sub


 '------------------------------------------  OpenAcct()  -------------------------------------------------
'	Name : OpenAcct()
'	Description : Account PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenAcqNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg
    
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�ڻ�����ȣ �˾�"			' �˾� ��Ī 
	arrParam(1) = "a_asset_master"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtCondAsstNo.value)		    ' Code Condition
	arrParam(3) = ""							    ' Name Condition
	arrParam(4) = ""							    ' Where Condition
	arrParam(5) = "�ڻ�����ȣ"				' �����ʵ��� �� ��Ī 
	
    arrField(0) = "acq_no"						    ' Field��(0)
	arrField(1) = "F2" & parent.gColSep & "convert(varchar(03),acq_seq)"	' Field��(1)
	arrField(2) = "asst_nm"

    arrHeader(0) = "�ڻ�����ȣ"				' Header��(0)
	arrHeader(1) = "������"					' Header��(1)
	arrHeader(2) = "�ڻ��"  					' Header��(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=600px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		field_fg = 1
		Call SetReturnVal(arrRet, field_fg)
	End If	
	
End Function


'===========================================================================
' Function Name : OpenDept
' Function Desc : OpenDeptCode Reference Popup
'===========================================================================

Function OpenAcctDeptPopUp(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(8)

	' ���Ѱ��� �߰�
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(strCode) 'strCode		            '  Code Condition
   	arrParam(1) = frm1.htxtCurrentDt.value
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAcctDept(arrRet, iWhere)
	End If	
End Function

 '------------------------------------------  SetReturnVal()  --------------------------------------------------
'	Name : SetReturnVal()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetAcctDept(byval arrRet, byval iWhere)

	With frm1.vspdData
		.Row = Cint(iWhere)
		.Col = C_DeptCd
		.text = arrRet(0)						
		frm1.htxtCurrentDt.value	= arrRet(3)
		
		Call vspdData_DeptCd_OnChange(iWhere)  
		lgBlnFlgChgValue = True

	End With
	
End Function

'==========================================================================================
'   Event Name : txtDeptCd_Change
'   Event Desc : 
'==========================================================================================

Sub vspdData_DeptCd_OnChange(byval iWhere)
	Dim IntRetCD
	Dim StrDeptCd
	Dim ArrDeptCd
	Dim ArrDeptNm
	Dim ArrOrgChangeId
	Dim	ArrCostNm
	Dim	ArrCostTypeNm
	
	If Trim(frm1.htxtCurrentDt.value = "") Then    
		Exit sub
    End If

    lgBlnFlgChgValue = True

	With frm1.vspdData
		.Row = iWhere
		.Col = C_DeptCd
		StrDeptCd = UCase(Trim(.text))
	End with

	If CommonQueryRs("A.DEPT_CD, A.DEPT_NM, A.ORG_CHANGE_ID, B.COST_NM, C.MINOR_NM ", _
					 "B_ACCT_DEPT A(NOLOCK)	JOIN B_COST_CENTER B(NOLOCK) ON A.COST_CD = B.COST_CD " _
					 & "JOIN B_MINOR C(NOLOCK) ON C.MINOR_CD = B.DI_FG AND C.MAJOR_CD = " & FilterVar("B9013", "''", "S") & "  ", _
					 "A.DEPT_CD =  " & FilterVar(StrDeptCd , "''", "S") & "" & " AND ORG_CHANGE_DT = (SELECT MAX(ORG_CHANGE_DT) FROM B_ACCT_DEPT(NOLOCK) " _
					 & "WHERE ORG_CHANGE_DT <=  " & FilterVar(UniConvDateAToB(frm1.htxtCurrentDt.value ,gDateFormat, parent.gServerDateFormat), "''", "S") & ")" ,_
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  

		With frm1.vspdData
			.Row = iWhere
			.Col = C_DeptCd
			.text = ""
			.Col = C_DeptNm
			.text = ""
			.Col = C_OrgChangeId
			.text = ""
			.Col = C_CostNm
			.text = ""
			.Col = C_CostTypeNm
			.text = ""
			.Col = C_DeptCd
		End with
	Else

		ArrDeptCd = Split(lgF0, parent.gColSep)
		ArrDeptNm = Split(lgF1, parent.gColSep)
		ArrOrgChangeId = Split(lgF2, parent.gColSep)
		ArrCostNm = Split(lgF3, parent.gColSep)
		ArrCostTypeNm = Split(lgF4, parent.gColSep)
			
		With frm1.vspdData
			.Row = iWhere
			.Col = C_DeptCd
			.text = ArrDeptCd(0)
			.Col = C_DeptNm
			.text = ArrDeptNm(0)
			.Col = C_OrgChangeId
			.text = ArrOrgChangeId(0)
			.Col = C_CostNm
			.text = ArrCostNm(0)
			.Col = C_CostTypeNm
			.text = ArrCostTypeNm(0)
		End with
			
	End If	
	
	'----------------------------------------------------------------------------------------

End Sub

 '==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
 '------------------------------------------  SetAcct()  --------------------------------------------------
'	Name : SetAcct()
'	Description : Account Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetReturnVal(ByVal arrRet, ByVal field_fg)
	
	Select case field_fg

		case 3	'OpenAcctCd
			frm1.txtAcctCd.Value		= arrRet(0)
			frm1.txtAcctNm.Value		= arrRet(1)
		case 4	'OpenMgmtId
			frm1.txtMgmtUserId.Value	= arrRet(0)
			frm1.txtMgmtUserNm.Value	= arrRet(1)
			lgBlnFlgChgValue = True
		case 5	'OpenCurrency
			frm1.txtDocCur.Value		= arrRet(0)
	End select	

End Function


Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	Call SetSpreadColor ("Q",-1, -1)

End Sub

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
                                                        '��: Load Common DLL
    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.FormatDate(frm1.txtYyyymm, parent.gDateFormat, 2)
    Call ggoOper.LockField(Document, "N")
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables

    '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    'Call InitComboBox
    
    frm1.txtCondAsstNo.focus 

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
'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	Dim StrDeptCd, StrOrgChangeId
	Dim	ArrCostNm
	Dim	ArrCostTypeNm

	if Col = C_DeptCd then
		Call vspdData_DeptCd_OnChange(Row)
	end if 

    Call CheckMinNumSpread(frm1.vspdData, Col, Row)  
   
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row	
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)
Dim strTemp
Dim intPos1

	With frm1.vspdData 

	If Row > 0 And Col = C_DeptCdPopUp Then
	    .Col = C_DeptCd
	    .Row = Row
	    strTemp = UCase(Trim(.text))
	        
	    Call OpenAcctDeptPopUp(strTemp, Row)
	End If
	
	If Row > 0 And Col = C_assetnoPopUp Then
	    .Col = C_assetnoPopUp-1
	    .Row = Row
	    strTemp = UCase(Trim(.text))
	        
	    OpenMasterRef "D", Row
	End If
	
	    
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1111111111")
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData

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
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1, ByVal Col2)		' 7) �÷� width ���� �̺�Ʈ �ڵ鷯
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(Col1,Col2)

End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell�� ����� �����ǹ߻� �̺�Ʈ 
'==========================================================================================

Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

Sub vspdData_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				'8) �÷� title ����
    Dim iColumnName

	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if

'    If Row <= 0 Then
'       frm1.vspdData.Row=Row
'       frm1.vspdData.Col=Col
'       iColumnName = frm1.vspdData.Text

'       iColumnName = AskSpdSheetColumnName(iColumnName)
        
'       If iColumnName <> "" Then
'          ggoSpread.Source = frm1.vspdData
'          Call ggoSpread.SSSetReNameHeader(Col,iColumnName)

          'Call SetSortFieldNM("A", frm1.vspdData,Col)
'       End If

        
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
'    End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData
		.Row = Row

		.Col = Col
		index = .Value
			
		.Col = 4
		.Value = index
	End With
End Sub
'========================================================================================
' Function Name : vspdData_TopLeftChange
' Function Desc : 
'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    if  lgIntFlgMode <> parent.OPMD_UMODE then exit sub
    ' If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
	' And Not( lgStrPrevToKey = "") Then
	
	 If Not( lgStrPrevToKey = "") Then
		Call DisableToolBar(parent.TBC_QUERY)					'�� : Query ��ư�� disable ��Ŵ.
		If DBQuery = False Then 
		   Call RestoreToolBar()
		   Exit Sub 
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

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                    '��: Processing is NG
    
    Err.Clear                                                           '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")								'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables
    																	'��: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							'��: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                    '��: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    
   
    
    FncNew = True                                                           '��: Processing is OK
    
    'SetGridFocus

End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                    '��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
		Call DisplayMsgBox("900002", "X", "X", "X")                                  '��:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")  '�� �ٲ�κ� 
    If IntRetCD = vbNo Then
        Exit Function
    End If

    If DbDelete = False Then                                                '��: Delete db data
       Exit Function                                                        '��:
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
    
    FncDelete = True                                                        '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    Dim var_m
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    On Error Resume Next                                                    '��: Protect system from crashing    
 
    ggoSpread.Source = frm1.vspdData
    var_m = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False and var_m = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")  '�� �ٲ�κ� 
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") then                                   '��: Check contents area
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then	
	Exit Function
    End if
   	
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave				                                                  '��: Save db data
    
'	frm1.vspdData.ReDraw = False
'	ggoSpread.SSDeleteFlag 1 , frm1.vspdData.MaxRows
'   Call SetSpreadLock
'	frm1.vspdData.ReDraw = True

	FncSave = True                                                          '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData
		.ReDraw = False
		
		ggoSpread.Source = frm1.vspdData 
	    ggoSpread.CopyRow
		SetSpreadColor "i", frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
		''''Key field clear
		.Col  = C_DeptCd
		.Text = ""
		
		.Col  = C_DeptNm
		.Text = ""
		
		.Col = C_CostNm
		.Text = ""
		
		.Col = C_CostType		
		.Text = ""
		
		.Col = C_CostTypeNm		
		.Text = ""
		
		.Col = C_InvQty		
		.Text = ""
						
		.ReDraw = True
    End With
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel()

	If frm1.vspdData.MaxRows < 1 Then Exit Function

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
    
    lgBlnFlgChgValue = False
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow(Byval pvRowCnt) 
		Dim IntRetCd
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
'		imRow = AskSpdSheetAddRowCount()
'		If imRow = "" then
'			Exit Function
'		End If
	
	With frm1	
	   ' If .txtCondAsstNo.value = "" Then
	'		IntRetCD = DisplayMsgBox("117326", "X", "X", "X") '''Please Insert Asset No.			
	'		Exit Function
	 '   End If    
	    
		.vspdData.focus
		ggoSpread.Source = .vspdData
		'.vspdData.EditMode = True
		.vspdData.ReDraw = False
		ggoSpread.InsertRow ,imRow
		.vspdData.Col  = C_InvQty
		.vspdData.Text = 0
		.vspdData.ReDraw = True
		    
		SetSpreadColor "i",.vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
	    lgBlnFlgChgValue = True
    End With
    
    Call SetToolbar("11101111001111")										'��: ��ư ���� ���� 
    
    Set gActiveElement = document.ActiveElement  
    
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    
	If frm1.vspdData.MaxRows < 1 Then Exit Function
    
    With frm1.vspdData 
	    .focus
		ggoSpread.Source = frm1.vspdData 
    
		lDelRows = ggoSpread.DeleteRow

		lgBlnFlgChgValue = True
    End With
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    On Error Resume Next                                                    '��: Protect system from crashing
    
    Call parent.FncPrint()
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
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
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
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
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
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 

	Dim strYear
	Dim strMonth
	Dim strDay
	Dim strYyyymm

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing

	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1

		strVal = BIZ_PGM_ID & "?txtMode="   & parent.UID_M0001						'��: 
				    
	    If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = strVal & "&txtCondAsstNo=" & Trim(.hAsstNo.value)				'�Ѱ��� ��� hidden�� �ʿ� ���� 
		Else
			strVal = strVal & "&txtCondAsstNo=" & Trim(.txtCondAsstNo.value)			    '��: ��ȸ ���� ����Ÿ 
		End If    

		Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

		strYyyymm = strYear & strMonth
		
		strVal = strVal & "&txtYyyymm=" & Trim(strYyyymm)			    '��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows="    & .vspdData.MaxRows
	    strVal = strVal & "&lgStrPrevToKey=" & lgStrPrevToKey
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
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	call Set1()
	InitData()
	'SetDefaultVal()
	SetSpreadColor "Q",-1, -1
	
End Function

Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
			
			.Col = C_CostType
			intIndex = .value
			.col = C_CostTypeNm
			.value = intindex
					
		Next	
	End With
	
	'SetGridFocus
	
End Sub


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          '��: Processing is NG
    
    On Error Resume Next                                                   '��: Protect system from crashing

	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value = parent.UID_M0002
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		strDel = ""
    
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag							'��: �ű� 
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep '��: C=Create, Row��ġ ���� 
		            .vspdData.Col = C_OrgChangeId
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_DeptCd
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_AssnRate
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
		             .vspdData.Col = C_ASSETNO
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep

		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.UpdateFlag							'��: ���� 
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep '��: U=Update
		            .vspdData.Col = C_OrgChangeId
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_DeptCd
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_AssnRate
		            strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & parent.gColSep
					.vspdData.Col = C_ASSETNO
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.DeleteFlag							'��: ���� 
					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep'��: U=Update
		            .vspdData.Col = C_OrgChangeId
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		            .vspdData.Col = C_DeptCd
		            strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
		           
		            strDel = strDel & "" & parent.gColSep
		           .vspdData.Col = C_ASSETNO
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            
		            lGrpCnt = lGrpCnt + 1
		    End Select

		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
'		.txtSpread.value = strVal

		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 

	End With
	
    DbSave = True                                                           '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
	Call InitVariables
	
	Call fncquery()
	
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
End Function

' �ϰ�����
Function fnButtonExec(Byval Jobtype)

	Dim IntRetCD
	Dim strYyyymm, strYear, strMonth, strDay

	IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO,"X","X")   '�� �ٲ�κ� 
	'RetFlag = Msgbox("�۾��� ���� �Ͻðڽ��ϱ�?", vbOKOnly + vbInformation, "����")
	If IntRetCD = VBNO Then
		Exit Function
	End IF  

	Call ExtractDateFrom(frm1.txtYyyymm.Text,frm1.txtYyyymm.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

	strYyyymm = strYear & strMonth

	Call LayerShowHide(1)
	
		Call CommonQueryRs(" COUNT(*) "," A_ASSET_INFORM_OF_DEPT_HISTORY "," YYYYMM = " & FilterVar(strYyyymm, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If Cint(Replace(lgF0,Chr(11),parent.gColSep)) > 0 Then
			if Jobtype="1" then 
			
				IntRetCD = DisplayMsgBox("800397", parent.VB_YES_NO,"X","X")   '�� �ٲ�κ� 
				If IntRetCD = VBNO Then
					Exit Function
				End IF 
		    end if 
			
	
		End If
		
	if Jobtype="1" then 
		frm1.txtMode.value = parent.UID_M0002
	else
		frm1.txtMode.value = parent.UID_M0003
	end if
	
	
	'���Ѱ����߰� start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'���Ѱ����߰� end

	Call ExecMyBizASP(frm1, BIZ_PGM_ID2)										'��: �����Ͻ� ASP �� ���� 

End Function

Function fnButtonExecOk()
    'Dim IntRetCD 

  
    call fncQuery()
    '  call DisplayMsgBox("990000","X","X","X")   '�� �ٲ�κ�    
	   
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'------------  Coding part  -------------------------------------------------------------
'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtYyyymm_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub


Sub txtYyyymm_DblClick(Button)
    If Button = 1 Then
       frm1.txtYyyymm.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtYyyymm.Focus       
    End If
End Sub

Sub txtYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub


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
									<TD CLASS="TD5" NOWRAP>��г��</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYyyymm" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT=��г�� id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>		
									<TD CLASS="TD5" NOWRAP>�ڻ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtCondAsstNo"  SIZE=18 MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="�ڻ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcqNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMasterRef 'H','' "> <INPUT TYPE="Text" NAME="txtCondAsstNm" SIZE=25 MAXLENGTH=30 tag="14" ALT="�ڻ��"></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>">
							<TR>
								<TD WIDTH="100%" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT=20>
  		<TD WIDTH="100%">
  			<TABLE <%=LR_SPACE_TYPE_30%>>
   				<TR>
   					<TD WIDTH=10>&nbsp;</TD>
   					<TD><BUTTON NAME="button1" CLASS="CLSMBTN" ONCLICK="vbscript:Call FnButtonExec(1)" Flag=0>�ϰ�����</BUTTON>&nbsp;
						<BUTTON NAME="button1" CLASS="CLSMBTN" ONCLICK="vbscript:Call FnButtonExec(0)" Flag=0>�ϰ�����</BUTTON>&nbsp;
   					</TD>
   				</TR>
   			</TABLE> 
  		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hAsstNo" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtCurrentDt" tag="24">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


