<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Account
*  2. Function Name        : 
*  3. Program ID           : A3103ma1
*  4. Program Name         : ä�ǽ��� 
*  5. Program Desc         : ä�ǽ��� �� �̽��� ��ȸ �� ���� ó�� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2002/12/06
*  9. Modifier (First)     : 
* 10. Modifier (Last)      : Jeong Yong Kyun
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit																	'��: indicates that All variables must be declared in advance

'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_ID1 = "a3103mb2.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "a3103mb1.asp"												'��: �����Ͻ� ���� ASP�� 

Dim lgIsOpenPop                                          
Dim lgQueryFlag																	' �ű���ȸ �� �߰���ȸ ���� Flag

Dim IsOpenPop          

<% 
Dim dtToday
dtToday = GetSvrDate                                                 
%>

Dim C_CHECK_FG
Dim C_AR_DT 
Dim C_GL_DT 
Dim C_AR_NO 
Dim C_BP_NM 
Dim C_DOC_CUR 
Dim C_AR_AMT 
Dim C_AR_LOC_AMT 
Dim C_DEPT_CD 
Dim C_TEMP_GL_NO 
Dim C_GL_NO 
Dim C_CONF_FG

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.1 Common Group -1
' Description : This part declares 1st common function group
'=======================================================================================================
'*******************************************************************************************************




'======================================================================================================
' Name : initSpreadPosVariables()
' Description : �׸���(��������) �÷� ���� ���� �ʱ�ȭ 
'=======================================================================================================
Sub initSpreadPosVariables()
	C_CHECK_FG   = 1
	C_AR_DT      = 2
	C_GL_DT      = 3
	C_AR_NO      = 4 
	C_BP_NM      = 5 
	C_DOC_CUR    = 6
	C_AR_AMT     = 7
	C_AR_LOC_AMT = 8
	C_DEPT_CD    = 9
	C_TEMP_GL_NO = 10
	C_GL_NO      = 11
	C_CONF_FG    = 12
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE							'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgStrPrevKey     = ""											'initializes Previous Key
    
    lgPageNo  = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	frm1.txtFromReqDt.text	=  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtToReqDt.text	=  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.GIDate.text		=  UniConvDateAToB("<%=dtToday%>", parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtFromReqDt.focus	
    frm1.cboConfFg.value	=	"U" 
    frm1.hOrgChangeId.value = parent.gChangeOrgId   	
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE" , "QA") %>  
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>                              
End Sub

'========================================================================================
'	Name : InitComboBox_cond()
'	Description :                        
' ========================================================================================  
Sub InitComboBox()
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1007", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboConfFg ,lgF0  ,lgF1  ,Chr(11))
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet()
    Call initSpreadPosVariables()
    
    With frm1.vspdData
    
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadInit "V20021128",,parent.gAllowDragDropSpread 

		.Redraw = False

		.MaxCols = C_CONF_FG + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.Col = .MaxCols															'������Ʈ�� ��� Hidden Column
		.ColHidden = True    
		.MaxRows = 0
		    
		Call AppendNumberPlace("6","3","0")
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck  C_CHECK_FG   , ""			       ,5 , ,"", True,-1
		ggoSpread.SSSetDate   C_AR_DT      , "ä������"    ,10, 2, parent.gDateFormat  
		ggoSpread.SSSetDate   C_GL_DT      , "��ǥ����"    ,10, 2, parent.gDateFormat  
		ggoSpread.SSSetEdit   C_AR_NO      , "ä�ǹ�ȣ"    ,20,3
		ggoSpread.SSSetEdit   C_BP_NM      , "�ŷ�ó"      ,12,3        
		ggoSpread.SSSetEdit   C_DOC_CUR    , "��ȭ"        ,12,3        
		ggoSpread.SSSetFloat  C_AR_AMT     , "ä�Ǿ�"      ,15, "A"  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_AR_LOC_AMT , "ä�Ǿ�(�ڱ�)",15, parent.ggAmtOfMoneyNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_DEPT_CD    , "�μ��ڵ�"    ,10, ,,10,2
		ggoSpread.SSSetEdit   C_TEMP_GL_NO , "������ǥ��ȣ",15,3        
		ggoSpread.SSSetEdit   C_GL_NO      , "ȸ����ǥ��ȣ",15,3        
		ggoSpread.SSSetCheck  C_CONF_FG    , ""			       ,3 , ,"", True,-1

		Call ggoSpread.MakePairsColumn(C_AR_AMT,C_AR_LOC_AMT)
		
		Call ggoSpread.SSSetColHidden(C_CONF_FG,C_CONF_FG,True)

		.Redraw = True 
    End With

	Call SetSpreadLock()
End Sub

'=======================================================================================================
'	Name : Openpopupgl()
'	Description : 
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim arrField
	Dim intFieldCount
	Dim i	
	
	dim TmpPopUpR
	Dim TInf(5)
	Dim ii,jj
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .maxRows > 0 then
				
			.Row = .ActiveRow
			.Col = C_GL_NO			

			arrParam(0) = Trim(.Text)	'ȸ����ǥ��ȣ 
			arrParam(1) = ""			'Reference��ȣ 
		End If
	End With
	
	IsOpenPop = True

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function
'=======================================================================================================
'	Name : OpenPopuptempGL()
'	Description : 
'=======================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim arrField
	Dim intFieldCount
	Dim i	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	With frm1.vspdData
		If .MaxRows > 0 then 
			
			.Row = .ActiveRow
			.Col =  C_TEMP_GL_NO
			arrParam(0) = Trim(.Text)	'Temp_gl_no
			arrParam(1) = ""			'Reference��ȣ 
		End If
	End With

	IsOpenPop = True

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function
'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtFromReqDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToReqDt.Text
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' �������� ���� Condition  

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
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
		frm1.txtFromReqDt.text = arrRet(4)
		frm1.txtToReqDt.text = arrRet(5)
		frm1.txtDeptCd.focus
End Function
'------------------------------------------  OpenBp()  ---------------------------------------
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
	arrParam(3) = ""									'ToDt
	arrParam(4) = "B"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.txtBpCd.focus
		End Select
		Exit Function
	Else
		Select Case iWhere
			Case 1
				frm1.txtBpCd.value=arrRet(0)
				frm1.txtBpNm.value= arrRet(1)
				frm1.txtBpCd.focus
		End Select

	End If	
End Function
 '========================================== 2.4.2 Open???()  =============================================
'	Name : OpenPopUp()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 1
			arrParam(0) = "�ŷ�ó �˾�"  				' �˾� ��Ī 
			arrParam(1) = "B_BIZ_PARTNER"	 				' TABLE ��Ī 
			arrParam(2) = strCode							' Code Condition
			arrParam(3) = ""								' Name Cindition
			arrParam(4) = ""								' Where Condition
			arrParam(5) = "�ŷ�ó"	    				' �����ʵ��� �� ��Ī 

			arrField(0) = "BP_CD"							' Field��(0)
			arrField(1) = "BP_NM"							' Field��(1)
    
			arrHeader(0) = "�ŷ�ó"	     				' Header��(0)
			arrHeader(1) = "�ŷ�ó��"					' Header��(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				frm1.txtBpCd.value  = arrRet(0)
				frm1.txtBpNm.value  = arrRet(1)			    
				frm1.txtBpCd.focus
		End Select

	End With
	
End Function

Sub txtBpCd_onBlur()
	
	if frm1.txtBpCd.value = "" then
		frm1.txtBpNm.value = ""
	end if
End Sub	

'==========================================================================================
'   Event Name : txtDeptCd_Onchange
'   Event Desc : 
'==========================================================================================
Function txtDeptCD_OnChange()
        
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
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromReqDt.Text, gDateFormat,""), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtToReqDt.Text, gDateFormat,""), "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

		' ���Ѱ��� �߰� 
		If lgInternalCd <> "" Then
			strWhere		= strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
		End If
	
		If lgSubInternalCd <> "" Then
			strWhere		= strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
		End If		
	
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
			txtDeptCD_OnChange = False
			'window.event.cancelBubble = True
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

End Function


'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	Dim arrField
	Dim intFieldCount
	Dim i,j	


    With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SpreadUnLock C_CHECK_FG	,-1 , C_CHECK_FG							'����üũ�ڽ� UnLoking
		ggoSpread.SpreadLock   C_AR_DT		,-1 ,C_AR_DT							'����üũ�ڽ����� ��ǥ�������� Locking
		ggoSpread.SpreadUnLock C_GL_DT		,-1 , C_GL_DT							'��ǥ�� UnLocking
		ggoSpread.SpreadLock   C_AR_NO		,-1									'��ǥ�ϴ������� ������ Locking
				
		.vspdData.ReDraw = True
    End With
End Sub

'================================== SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim arrField
	Dim intFieldCount
	Dim i,j
	
    With frm1
	    .vspdData.ReDraw = False

		.vspdData.Col = C_CONF_FG

		If .vspdData.Text = "C" Then
			ggoSpread.SSSetProtected	C_GL_DT,  pvStartRow, pvEndRow
		Else
			ggoSpread.SSSetRequired		C_GL_DT,  pvStartRow, pvEndRow
		End If	
		
		.vspdData.ReDraw = True
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
				C_CHECK_FG   = 1
				C_AR_DT      = 2
				C_GL_DT      = 3
				C_AR_NO      = 4 
				C_BP_NM      = 5 
				C_DOC_CUR    = 6
				C_AR_AMT     = 7
				C_AR_LOC_AMT = 8
				C_DEPT_CD    = 9
				C_TEMP_GL_NO = 10
				C_GL_NO      = 11
				C_CONF_FG    = 12
	End select
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029()														

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)    
    Call ggoOper.LockField(Document, "N")                                   

	Call InitVariables()														
	Call InitSpreadSheet()
	Call InitComboBox()	
    Call SetDefaultVal()	

    Call SetToolBar("1100000000000111")		

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
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 
Sub txtFromReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtToReqDt.focus
		Call FncQuery
	end if
End Sub

Sub txtToReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt.focus
		Call FncQuery
	End if
End Sub

Sub GIDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt.focus
		Call FncQuery
	end if
End Sub
'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If                                                  
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables() 											
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							
		Exit Function
    End If
    
	If CompareDateByFormat(frm1.txtFromReqDt.text,frm1.txtToReqDt.text,frm1.txtFromReqDt.Alt,frm1.txtToReqDt.Alt, _
        	               "970025",frm1.txtFromReqDt.UserDefinedFormat,parent.gComDateType, True) = False Then
		frm1.txtFromReqDt.focus
		Exit Function
	End If    
	
	IF NOT CheckOrgChangeId Then
		  IntRetCD = DisplayMsgBox("800600","X",frm1.txtFromReqDt.alt,"X")            '��: Display Message(There is no changed data.)
		Exit Function
	End if

	lgQueryFlag = "New"		' �ű���ȸ �� �߰���ȸ ���� Flag (����� �ű���)	
    '-----------------------
    'Query function call area
    '-----------------------

    Call DbQuery()																'��: Query db data

    FncQuery = True													
	Set gActiveElement = document.ActiveElement 
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X") '�� �ٲ�κ�    
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '��: Clear Contents  Field
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitVariables()                                                      '��: Initializes local global variables
    Call SetDefaultVal()
    
    FncNew = True       
    Set gActiveElement = document.ActiveElement                                                     '��: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
	DIm lRow
	Dim Fg
	dim GLDT
	
    FncSave = False																	'��: Processing is NG

    On Error Resume Next															'��: Protect system from crashing    
    Err.Clear																		'��: Clear err status    
   
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then			'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'��: Display Message(There is no changed data.)
        Exit Function
    End If

    If Not chkField(Document, "1") Then												'��: Check required field(Single area)
       Exit Function
    End If

	With frm1
		For lRow = 1 To .vspdData.MaxRows
				.vspdData.Row = lRow
				.vspdData.Col = C_CHECK_FG
			If Trim(.vspdData.Text) = "1" Then
				.vspdData.Col = C_CONF_FG
				Fg= .vspdData.text
				.vspdData.Col = C_GL_DT
				If Trim(.vspdData.text)= "" And Trim(Fg) = "U" Then
					Call DisplayMsgBox("117523","X","X","X") 
				    Exit Function				
				End if	
				GLDT=Trim(.vspdData.text)
				.vspdData.Col = C_AR_DT
				If CompareDateByFormat(.vspdData.text,GLDT,"ä������","��ǥ��", _
        	               "970023",.txtFromReqDt.UserDefinedFormat ,parent.gComDateType, True) = False Then
					Exit Function
				End If								
			End if
		Next
	End With
	'-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData

    If Not ggoSpread.SSDefaultCheck Then											'��: Check contents area
		Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																	'��: Save db data

    Set gActiveElement = document.ActiveElement      	

    FncSave = True 
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    If frm1.vspdData.MaxRows < 1 then Exit Function

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo															'��: Protect system from crashing    
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call parent.FncPrint()                                                       '��: Protect system from crashing
    FncPrint = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '��: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call parent.FncExport(parent.C_MULTI)

    FncExcel = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call parent.FncFind(parent.C_MULTI, True)

    FncFind = True                                                               '��: Processing is OK
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
    
    iColumnLimit  = Frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
    End If   
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = parent.SS_SCROLLBAR_BOTH
End Function

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '��: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    Call LayerShowHide(1)
    
    Err.Clear																				'��: Protect system from crashing
    
    With frm1
		strVal = BIZ_PGM_ID

		strVal = strVal & "?txtMode="      & parent.UID_M0001							'��:��ȸǥ�� 
		strVal = strVal & "&txtBpCd="     & Trim(.txtBpCd.value)	 			    '��: ��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtDeptCd="    & Trim(.txtDeptcd.value)
		strVal = strVal & "&cboConfFg="    & Trim(.cboConfFg.value)
		strVal = strVal & "&txtFromReqDt=" & UNIConvDate(Trim(.txtFromReqDt.Text))
		strVal = strVal & "&txtToReqDt="   & UNIConvDate(Trim(.txtToReqDt.Text))
		strVal = strVal & "&OrgChangeId=" & Trim(.hOrgChangeId.Value)
		strVal = strVal & "&txtProject="    & Trim(.txtProject.value)
		strVal = strVal & "&txtMaxRows="   & frm1.vspdData.MaxRows
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgPageNo="     & lgPageNo         		

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

		Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 
        
    End With

    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()		
	Dim arrField
	Dim intFieldCount
	Dim i	

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE													'��: Indicates that current mode is Update mode
   
    Call ggoOper.LockField(Document, "Q")													'��: This function lock the suitable field
    Call LayerShowHide(0)
    
    Call SetToolBar("110010000001111")    

	Call SetSpreadColor(1, frm1.vspddata.maxrows)
	
	If UCase(Trim(frm1.cboConfFg.value)) = "C" Then
	    frm1.vspdData.ReDraw = False
		ggoSpread.source = frm1.vspdData
		ggoSpread.SpreadLock C_GL_DT, 1, C_GL_DT, frm1.vspdData.MaxRows

		frm1.vspdData.ReDraw = True
	End If
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal

	Dim arrField_S
	Dim intFieldCount
'	Dim i,j,k
	
    DbSave = False																			'��: Processing is NG
    Call LayerShowHide(1)

    On Error Resume Next																	'��: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002

		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 1
			If Trim(.vspdData.Text) = "1" Then
				strVal = strVal & "U" & parent.gColSep										'��: U=Update
 				.vspdData.Col = C_AR_NO  
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep						'AR_NO

				.vspdData.Col = C_GL_DT  
				strVal = strVal & UNIConvDate(Trim(.vspdData.Text)) & parent.gColSep		'Journal Date
				

				.vspdData.Col = C_CONF_FG														'CONF_FG
				If .vspdData.Text = "U" Then
					strVal = strVal & "C" & parent.gRowSep
				Else
					strVal = strVal & "U" & parent.gRowSep
				End If	

				lGrpCnt = lGrpCnt + 1
			End If
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal	

		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID1)												'��: �����Ͻ� ASP �� ���� 
	End With

    DbSave = True																			'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()																			'��: ���� ������ ���� ���� 
    Call LayerShowHide(0)

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	ggoSpread.SSDeleteFlag 1 , frm1.vspdData.MaxRows
	ggoSpread.ClearSpreadData
    Call SetSpreadLock
	frm1.vspdData.ReDraw = True
	
	lgQueryFlag = "Save"																	' �ű���ȸ �� �߰���ȸ ���� Flag (����� �߰���ȸ(������)��)	
	
	IF frm1.cboConfFg.value = "C" Then
		frm1.cboConfFg.value = "U"
	Else
		frm1.cboConfFg.value = "C"
	End If
	
	Call InitVariables()	
	Call DBQuery()		
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 

End Function



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************




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
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("0001111111")
    gMouseClickStatus = "SPC"									'Split �����ڵ� 
    
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.Maxrows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col							'Ascending Sort
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col,lgSortKey				'Descending Sort
			lgSortKey = 1
		End If										
		Exit Sub
	End If		

End Sub	

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub    

'======================================================================================================
'   Event Name :vspddata_ScriptDragDropBlock
'   Event Desc :
'=======================================================================================================
Sub  vspddata_ScriptDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	ggoSpread.Source = frm1.vspdData 
	Call ggoSpread.SpreadDragDropBlock(Col,Row,Col2,Row2,NewCol,NewRow,NewCol2,NewRow2,Overwrite,Action,DataOnly,Cancel)
	Call GetSpreadColumnPos("A")
End Sub
'==========================================================================================
'   Event Name : vspdData_Change
'   Event Desc : 
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
	lgBlnFlgChgValue = true
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : This event is spread sheet data Button Clicked
'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
	    If Row >= NewRow Then
			Exit Sub
		End If
    End With
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================
'   Event Name :vspdData_KeyPress
'   Event Desc :
'==========================================================================================
Sub vspdData_KeyPress(index , KeyAscii )
     lgBinFlgChgValue = True                                                 '��: Indicates that value changed
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row <=0 Then
		Exit Sub				
	End If		

    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub
	
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If    
	
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then	
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub
'======================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : �󼼳��� �׸����� (��Ƽ)�÷��� �ʺ� �����ϴ� ��� 
'=======================================================================================================
Sub  vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.Source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'=======================================================================================================
'   Event Name : txtDueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFromReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtFromReqDt.Focus 
    End If
End Sub

Sub txtToReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtToReqDt.Focus 
    End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtFromEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtToEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'=======================================================================================================
'   Event Name : GIDate_Change()
'   Event Desc :  
'=======================================================================================================
Sub GIDate_DblClick(Button)
    If Button = 1 Then
        frm1.GIDate.Action = 7
		Call SetFocusToDocument("M")
		Frm1.GIDate.Focus 
    End If
    
End Sub

Sub GIDate_Change()
	Dim gDate
	Dim IRow
	
	gDate = frm1.GIDate.Text

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
	If UCase(Trim(frm1.cboConfFg.value)) = "C" Then
		Exit Sub
	End If

	frm1.vspdData.Col = C_GL_DT

	For IRow = 1 To frm1.vspdData.MaxRows
		frm1.vspdData.Row  = IRow
		frm1.vspdData.Text = gDate
	Next
	lgBlnFlgChgValue = True
End Sub
'==========================================================================================
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtFromReqDt.Text, gDateFormat,""), "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(.txtToReqDt.Text, gDateFormat,""), "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With
		'----------------------------------------------------------------------------------------

End Function
'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kCM.inc"  -->	
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
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A>&nbsp;</TD>
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
									<TD CLASS="TD5"NOWRAP>ä������</TD>
									<TD CLASS="TD6"NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFromReqDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="ä������" id=fpDateTime1></OBJECT>');</SCRIPT>~ 
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToReqDt" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="ä������" id=fpDateTime2></OBJECT>');</SCRIPT>										
									</TD>
									<TD CLASS="TD5"NOWRAP>�μ�</TD>
									<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtDeptCd" SIZE=10  MAXLENGTH=10 tag="11XXXU" ALT="�μ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDeptOrgPopup()">
										 <INPUT TYPE=TEXT ID="txtDeptNm" NAME="txtDeptNm" SIZE=20 tag="14X">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>���λ���</TD>																		
									<TD CLASS="TD6"NOWRAP><SELECT NAME="cboConfFg" tag="12" STYLE="WIDTH:82px:" Alt="���λ���"></SELECT> 
									<TD CLASS="TD5"NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6"NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON" ONCLICK="vbscript:CALL OpenBp(frm1.txtBpCd.Value, 1)">
										 <INPUT TYPE=TEXT ID="txtBpNm" NAME="txtBpNm" SIZE=20 tag="14X">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>��ǥ����</TD>
									<TD CLASS="TD6"NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=GIDate name=GIDate CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="��ǥ����"></OBJECT>');</SCRIPT></TD>
									<TD CLASS=TD5 NOWRAP>������Ʈ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME=txtProject ALT="������Ʈ" MAXLENGTH=25 SIZE=25 tag="11"></TD>
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
								<TD HEIGHT="100%" COLSPAN="4"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD class=TDT NOWRAP></TD>
								<TD class=TD6 NOWRAP></TD>
								<TD class=TD5 NOWRAP>ä�Ǿ�(�ڱ�)</TD>
								<TD class=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTotArLocAmt" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="ä�Ǿ�(�ڱ�)" tag="24X2" id=OBJECT22></OBJECT>');</SCRIPT></TD>
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
		<TD WIDTH=100% HEIGHT= <%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hDeptCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hBizCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hcboConfFg" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtWorkFg" tag="24">
<INPUT TYPE=HIDDEN NAME="htxtFocus" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromReqDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hToReqDt" tag="24">
<INPUT		TYPE=hidden	 NAME="hOrgChangeId"	tag="14" TABINDEX="-1">
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

