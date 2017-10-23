<%@ LANGUAGE="VBSCRIPT" %>

<!--
'=======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : A_RECEIPT
'*  3. Program ID           : A3119ma1
'*  4. Program Name         : �������ܾ����� 
'*  5. Program Desc         : �Ա�û�� 
'*  6. Modified date(First) : 2000/09/25
'*  7. Modified date(Last)  : 2000/12/20
'*  8. Modifier (First)     : �强�� 
'*  9. Modifier (Last)      : hersheys
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================

'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"     SRC="../ar/AcctCtrl3.vbs">				</SCRIPT>
<SCRIPT LANGUAGE=vbscript>

Option Explicit																	'��: indicates that All variables must be declared in advance

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'@PGM_ID
Const BIZ_PGM_ID         = "a3119mb1.asp"									' F_PrPaym_Sttl �� CRUD

Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"								'��: ȯ������ �����Ͻ� ���� ASP�� 

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

Dim C_ItemSeq   
Dim C_AdjustDt  
Dim C_AcctCd    
Dim C_AcctPopUp 
Dim C_AcctNm	
Dim C_AdjustAmt   
Dim C_AdjustLocAmt
Dim C_DocCur     
Dim C_DocCurPopUp
Dim C_AdjustDESC
Dim C_Temp_GlNo
Dim C_GlNo 
Dim C_RefNo


Dim  lgStrPrevKeyDtl
Dim  lgStrPrevKey2
Dim  lgStrPrevKey3
Dim  lgCurrRow
Dim  lgPrevNo
Dim  lgNextNo

Dim  IsOpenPop	                'Popup
Dim  gSelframeFlg
Dim  dtToday
dtToday = "<%=GetSvrDate%>"

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




'========================================================================================================= 
' Name : initSpreadPosVariables()
' Description : �׸���(��������) �÷� ���� ���� �ʱ�ȭ 
'========================================================================================================= 
Sub initSpreadPosVariables()
   
End Sub

'=======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub  InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
    lgStrPrevKey = 0                            'initializes Previous Key
    lgStrPrevKeyDtl = 0                         'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    frm1.hOrgChangeId.value = parent.gChangeOrgId
End Sub

'=======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub  SetDefaultVal()
	frm1.txtadjustDt.text = UniConvDateAToB(dtToday, parent.gServerDateFormat,gDateFormat)
	frm1.txtDocCur.value = parent.gcurrency	
	 Call txtDocCur_OnChange()   
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub  LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub  InitSpreadSheet()
    frm1.txtadjustDt.text = UniConvDateAToB(dtToday, parent.gServerDateFormat,gDateFormat)
End Sub

'=======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpreadLock()
    With frm1
	
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
	
    End With
End Sub

'======================================================================================================
' Function Name : SetSpread2ColorAR
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub  SetSpread2ColorAR()
	Dim i

    With frm1
		ggoSpread.Source = .vspdData2

		.vspdData2.ReDraw = False
		                
		for i = 1 to .vspddata2.maxrows
			ggoSpread.SSSetProtected C_DtlSeq   , i, i
			ggoSpread.SSSetProtected C_CtrlCd   , i, i
			ggoSpread.SSSetProtected C_CtrlNm   , i, i
			ggoSpread.SSSetProtected C_CtrlValNm, i, i

			.vspddata2.Col = C_DrFg		
		
			If (.vspddata2.text = "C" And .vspddata2.text <> "") _
                            Or .vspddata2.text = "Y" Or .vspddata2.text = "DC" Then
				ggoSpread.SSSetRequired C_CtrlVal, i, i	' 
			End If
		Next
		
		.vspdData2.ReDraw = True
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
   
End Sub
'=========================================================================================================
'	Name : OpenAdjustNo()
'	Description : Ref ȭ���� call�Ѵ�. : ä�ǹ߻����� 
'========================================================================================================= 
Function OpenAdjustNo()
	
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a3506ra2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3506ra2", "X")
		IsOpenPop = False
		Exit Function
	End If
   
	IsOpenPop = True

	arrParam(0) = ""				' �˻������� ������� �Ķ���� 
	arrParam(1) = ""				
	arrParam(2) = ""			
	arrParam(3) = "M"


	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) = "" Then		
		Exit Function
	Else		
		Call SetAdjustNo(arrRet)
	End If
End Function
'======================================================================================================
'   Function Name : SetAdjustNo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetAdjustNo(Byval arrRet)
	With frm1
		frm1.txtAdjustNo.value	= arrRet(0)
		frm1.txtAdJustNo.focus
	End With
End Function

'=======================================================================================================
'	Name : Openpopupgl()
'	Description : 
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim RetFlag
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(frm1.txtGlNo.value) = "" Then
		RetFlag = DisplayMsgBox("970000","X" , frm1.txtGlNo.Alt, "X") 	
		IsOpenPop = False
		Exit Function
	End If
	arrParam(0) = Trim(frm1.txtGlNo.value)
	arrParam(1) = ""			'Reference��ȣ 


	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent, arrParam), _
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
	Dim RetFlag
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(frm1.txtTempGlNo.value) = "" Then
		RetFlag = DisplayMsgBox("970000","X" , frm1.txtTempGlNo.Alt, "X") 	
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'������ǥ��ȣ 
	arrParam(1) = ""			'Reference��ȣ 

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'=======================================================================================================
'	Name : OpenRcptNo()
'	Description : Prepayment No PopUp
'=======================================================================================================
Function OpenRcptNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True  Then Exit Function
	
	iCalledAspName = AskPRAspName("a3119ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a3119ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	If Trim(frm1.txtAdJustNo.value) <> "" And lgIntFlgMode = parent.OPMD_UMODE Then Exit Function

	IsOpenPop = True


	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetRcptNo(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetRcptNo()
'	Description : PrpaymNo Popup���� Return�Ǵ� �� setting
'***	�Ʒ��� sql �� �ش� db ���� �����غ��� ÷�� �״�� ����Ѵ�.	�� ������ �������Ŀ� �����.
'***	select field_nm + '---'  + field_cd+ '---arrRet('  +  rtrim(convert(char(3), key_tag-1 )) + ')'  
'***	from z_ado_field_inf
'***	where pgm_id = 'a3119ra1'
'***	and lang_Cd = 'ko'
'***	order by seq_no

'=======================================================================================================
Function SetRcptNo(byval arrRet)
	With frm1
		.txtRcptDt.text  = arrRet(0)
		.txtBpCd.value = arrRet(1)
		.txtBpNM.value  = arrRet(2)
		.txtDeptCd.value  = arrRet(3)
		.txtDeptNm.value = arrRet(4)
		.txtRefNo.value = arrRet(5)
		.txtRcptNo.value = arrRet(6)
		.txtDocCur.value  = arrRet(7)
		.txtXchRate.text = arrRet(8)
		.txtRcptAmt.text  = arrRet(9)
		.txtRcptLocAmt.text = arrRet(10)
		.txtBalAmt.text  = arrRet(11)
		.txtBalLocAmt.text = arrRet(12)
		.txtRcptDesc.value = arrRet(14)
		

		Call txtDocCur_OnChange()
		If Trim(.txtDocCur.value) <> "" Then
			Call ggoOper.SetReqAttr(frm1.txtDocCur,   "Q")
		End If
		
	End With		
    lgBlnFlgChgValue = True
End Function


'======================================================================================================
'   Function Name : OpenPopup(Byval strCode, Byval iWhere)
'   Function Desc : 
'======================================================================================================
Function  OpenPopup(Byval strCode, iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	
	Select Case iWhere
		Case 1
			
			If frm1.txtAcctCd.className = "protected" Then Exit Function
			
			arrParam(0) = "�����ڵ��˾�"								' �˾� ��Ī 
			arrParam(1) = "A_Acct, A_ACCT_GP" 								' TABLE ��Ī 
			arrParam(2) = Trim(strCode)										' Code Condition
			arrParam(3) = ""												' Name Condition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "												' Where Condition
			arrParam(5) = "�����ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "A_ACCT.Acct_CD"									' Field��(0)
			arrField(1) = "A_ACCT.Acct_NM"									' Field��(1)
			arrField(2) = "A_ACCT_GP.GP_CD"									' Field��(2)
			arrField(3) = "A_ACCT_GP.GP_NM"									' Field��(3)
					
			arrHeader(0) = "�����ڵ�"									' Header��(0)
			arrHeader(1) = "�����ڵ��"									' Header��(1)
			arrHeader(2) = "�׷��ڵ�"									' Header��(2)
			arrHeader(3) = "�׷��"										' Header��(3)
    
		Case 2
			
		
		End Select
	IsOpenPop = True	
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtAcctCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'======================================================================================================
'   Function Name : SetPopUp(Byval arrRet,byval iWhere)
'   Function Desc : 
'======================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
					.txtAcctCd.value = arrRet(0)
					.txtAcctNm.value  = arrRet(1)
					Call txtAcctCd_OnChange()
					.txtAcctCd.focus
			Case 2
		End Select
		
	    lgBlnFlgChgValue = True
	End With
End Function



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.2 Common Group-2
' Description : This part declares 2nd common function group
'=======================================================================================================
'*******************************************************************************************************




'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'======================================================================================================
Sub  Form_Load()
   	
    Call LoadInfTB19029()																'Load table , B_numeric_format
    Call ggoOper.ClearField(Document, "1")										'��: Condition field clear
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, _
							parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
	Call InitCtrlSpread()
	Call InitVariables()																'Initializes local global variables
    
    Call SetDefaultVal()
    Call SetToolbar("1110100000001111")
	frm1.txtAdJustNo.focus
   

    lgBlnFlgChgValue = False            

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

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'======================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
    
End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'======================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    Dim  var2
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then													'This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Check previous data area
    '-----------------------
   
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True  Or var2 = True Then		
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")	    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables()																'Initializes local global variables
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()																		'��: Query db data
    
    FncQuery = True		
    	
	Set gActiveElement = document.activeElement    
															
End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function  FncNew() 
    Dim IntRetCD 
	Dim var1, var2
	    
    FncNew = False                                                          
    
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True  Or var2 = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")               
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
    Call ggoOper.ClearField(Document, "2")										'Clear Condition Field
    Call ggoOper.LockField(Document, "N")										'Lock  Suitable  Field
    Call InitVariables()																'Initializes local global variables
    
    Call SetDefaultVal()
    Call txtDocCur_OnChange()
	Call DisableRefPop()    

	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
    
    lgBlnFlgChgValue = False            
    
    FncNew = True    
    	
	Set gActiveElement = document.activeElement    
	                                                      
End Function

'=======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function  FncDelete() 
    Dim IntRetCD
    
    FncDelete = False                                                      
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then										'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")					'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete()																	'��: Delete db data
    
    FncDelete = True     
    	
	Set gActiveElement = document.activeElement    
	                                                   
End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function  FncSave() 
    Dim IntRetCD 
	Dim var1,var2
	
    FncSave = False                                                         
    
    Err.Clear                                                               
        
 

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    If lgBlnFlgChgValue = False  And var2 = False  Then				'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")								'��: Display Message(There is no changed data.)
		Exit Function
    End If

	If Not chkField(Document, "2") Then												'��: Check required field(Single area)
		Exit Function
    End If
   

	If Trim(frm1.txtRcptNo.value)  = "" Then
		IntRetCD = DisplayMsgBox("112700","X","X","X")									'�Ա�����check
        Exit Function
    End If
    
	
	
    If CheckSpread4 = False Then
	IntRetCD = DisplayMsgBox("110420","X","X","X")									'�ʼ��Է� check!!
        Exit Function
    End If
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()  																	'��: Save db data
    
    FncSave = True      
    	
	Set gActiveElement = document.activeElement    
	                                                 
End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function  FncCopy() 
	Dim  IntRetCD
	
	
End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function  FncCancel() 
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
     	
	Set gActiveElement = document.activeElement    
	
End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function  FncInsertRow(ByVal pvRowcnt) 
	Dim iCurRowPos
	Dim imRow
    Dim ii
    
	On Error Resume Next															'��: If process fails
    Err.Clear   
	
    FncInsertRow = False															'��: Processing is NG

    If IsNumeric(Trim(pvRowcnt)) Then 
		imRow  = Cint(pvRowcnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
    End If
        
    Call ggoOper.LockField(Document, "I")									'This function lock the suitable field
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function  FncDeleteRow() 
   	Dim lDelRows
    Dim DelItemSeq

End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function  FncPrint() 
    On Error Resume Next                                               
    FncPrint()
End Function

'=======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function  FncPrev() 
    On Error Resume Next                                               
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function  FncNext() 
    On Error Resume Next                                               
End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'========================================================================================================
Function  FncFind() 
    Call FncFind(parent.C_SINGLEMULTI , True) 
    	
	Set gActiveElement = document.activeElement    
	                         
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function  FncExcel() 
	Call FncExport(parent.C_SINGLEMULTI)
	
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
    
    iColumnLimit  = 5
    
   
    If gMouseClickStatus = "SP2CRP" Then
		ACol = Frm1.vspdData2.ActiveCol
		ARow = Frm1.vspdData2.ActiveRow

		If ACol > iColumnLimit Then
				Frm1.vspdData2.Col = iColumnLimit : frm1.vspdData2.Row = 0  	 	 	 	 		
				iRet = DisplayMsgBox("900030", "X", Trim(frm1.Vspddata2.text), "X")
				Exit Function  
		End If   
    
		Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_NONE
    
		ggoSpread.Source = Frm1.vspdData2
    
		ggoSpread.SSSetSplit(ACol)    
    
		Frm1.vspdData2.Col = ACol
		Frm1.vspdData2.Row = ARow
    
		Frm1.vspdData2.Action = Parent.SS_ACTION_ACTIVE_CELL     
    
		Frm1.vspdData2.ScrollBars = parent.SS_SCROLLBAR_BOTH
    End If   
	
	Set gActiveElement = document.activeElement    
		
End Function

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function  FncExit()
	Dim IntRetCD
	Dim var1,var2
	
	FncExit = False

    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange
    
    If lgBlnFlgChgValue = True or var2 = True Then					'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
  
    FncExit = True
	
	Set gActiveElement = document.activeElement    
	
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.3 Common Group - 3
' Description : This part declares 3rd common function group
'=======================================================================================================
'*******************************************************************************************************



'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================================
Function  DbDelete() 
    Dim strVal

    DbDelete = False														
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtRcptNo=" & Trim(frm1.txtRcptNo.value)				'��: ���� ���� ����Ÿ    
    strVal = strVal & "&txtAdjustNo=" & Trim(frm1.txtAdjustNo.value)				'��: ���� ���� ����Ÿ    

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
   
    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    DbDelete = True                                                         
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()												        '���� ������ ���� ���� 
	Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
	Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    Call InitVariables()                                                      'Initializes local global variables
    Call SetDefaultVal()
    Call DisableRefPop()
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData	
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================
Function DbQueryOk()													'��: ��ȸ ������ �������	
	With frm1
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field        
        Call SetToolbar("1111100000001111")                                     '��ư ���� ���� 
         Call DbQuery2()
        
    End With
    
 
	Call txtDocCur_OnChange()
	Call DisableRefPop()
	lgBlnFlgChgValue = False	
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'=======================================================================================================
Function  DbQuery() 
    Dim strVal
    
    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'��: 
			strVal = strVal & "&txtAdjustNo=" & Trim(.txtAdjustNo.value)				'��ȸ ���� ����Ÿ 
			
			
			
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001					'��: 
			strVal = strVal & "&txtAdjustNo=" & Trim(.txtAdjustNo.value)				'��ȸ ���� ����Ÿ 
		End If
    End With

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 

	Call RunMyBizASP(MyBizASP, strVal)												'��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function  DbQueryOk1()
	With frm1
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field        
        Call SetToolbar("1110100000001111")                                     '��ư ���� ���� 
        Call DbQuery2()
        
    End With
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'=======================================================================================================
Function  DbSave() 
    Dim lngRows 
    Dim lGrpcnt
    DIM strVal     
    Dim strDel
    Dim RowD
    DIM GrpCntD
    DIM strValD
    DIM strItemSEQ	'�����׸� �Ķ���� 

    DbSave = False                                                          
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   
	Err.Clear 
	
	With frm1
		.txtFlgMode.value = lgIntFlgMode
		.txtMode.value = parent.UID_M0002
	End With
    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data ���� ��Ģ 
    ' 0: Sheet��, 1: Flag , 2: Row��ġ, 3~N: �� ����Ÿ 

    lGrpCnt = 1
    
    GrpCntD = 1: strValD = ""	'�����׸� �Ķ���� 
    
	'=======================================================================
	'2001.06.18 Song,MunGil �����׸� �Է�/������ �ɷ� �����ϰ� ������.
	'=======================================================================
	With frm1.vspdData2
	For RowD = 1 To .MaxRows
		.Row = RowD
		.Col = 0
'		If (.Text = ggoSpread.InsertFlag or .Text = ggoSpread.UpdateFlag) then
		If Trim(.Text) <> ggoSpread.DeleteFlag then
			strValD = strValD & "C" & parent.gColSep & RowD & parent.gColSep
			strValD = strValD & "1" & parent.gColSep
			.Col = C_DtlSeq 
			strValD = strValD & Trim(.Text) & parent.gColSep
			.Col = C_CtrlCd
			strValD = strValD & Trim(.Text) & parent.gColSep
			.Col = C_CtrlVal
			strValD = strValD & Trim(.Text) & parent.gRowSep
										
			GrpCntD = GrpCntD + 1
		End If
	Next
	End With				

	
	frm1.txtMaxRows2.value = GrpCntD - 1									'Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread2.value  = strValD				

	'���Ѱ����߰� start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'���Ѱ����߰� end
		
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'���� �����Ͻ� ASP �� ���� 

    DbSave = True
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function  DbSaveOk()													'��: ���� ������ ���� ���� 
    ggoSpread.SSDeleteFlag 1
    
	Call ggoOper.ClearField(Document, "2")										'Clear Contents  Field
    Call InitVariables()															'Initializes local global variables
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData	

	Call DbQuery()		
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************




'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbQuery2()
	Dim strVal	
	Dim lngRows
		
	Dim strSelect
	Dim strFrom
	Dim strWhere 	
	
	Dim strTableid
	Dim strColid
	Dim strColNm	
	Dim strMajorCd	
	Dim strNmwhere
	Dim i,Indx1
	Dim arrVal,arrTemp
	
	Err.Clear
	
	With frm1
	  
	   
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.ColM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , "
		strSelect = strSelect & " 1 , LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.ColM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')),CHAR(8) "
  		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_RCPT_ADJUST_DTL C (NOLOCK), A_RCPT_ADJUST D (NOLOCK) "
		
					
		strWhere =			  " D.ADJUST_NO =  " & FilterVar(UCase(.txtAdjustNO.value), "''", "S") & " "		
		strWhere = strWhere & " AND D.ADJUST_NO  =  C.ADJUST_NO  "		
		strWhere = strWhere & "	AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD = B.CTRL_CD "
		strWhere = strWhere & " AND D.ACCT_CD = B.ACCT_CD "
		strWhere = strWhere & " AND B.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
   					
		frm1.vspdData2.ReDraw = False
			
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = frm1.vspdData2
			arrTemp =  Split(lgF2By2,Chr(12))
			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next
			lgF2By2 = Join(arrTemp,Chr(12))			
			ggoSpread.SSShowData lgF2By2							
			
			For lngRows = 1 To frm1.vspdData2.Maxrows
				frm1.vspddata2.Row = lngRows	
				frm1.vspddata2.Col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspddata2.Col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspddata2.Col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspddata2.Col = C_ColNm
					strColNm = frm1.vspddata2.text	
					frm1.vspddata2.Col = C_MajorCd					
					strMajorCd = frm1.vspddata2.text	
					
					frm1.vspddata2.Col = C_CtrlVal
					
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text , "''", "S") & " " 
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") & " "
					End If				 
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.Col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
					End If
				End If								
			Next					
			
		End If 		
	
		Call SetSpread2ColorAR()
	End With
	
	Call LayerShowHide(0)
	
	frm1.vspdData2.ReDraw = True
	
	DbQuery2 = True
End Function

'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function  DbQueryOk2()
	Call SetSpread2ColorAR()
    Call txtDocCur_OnChange()
End Function

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    Dim arrVal
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY = " & FilterVar(frm1.txtDocCur.value, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		arrVal = Split(lgF0, Chr(11))  
		'frm1.txtDocCurNm.value = arrVal(0)
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If	  
End Sub

'===================================== DisableRefPop()  =======================================
'	Name : DisableRefPop()
'	Description :
'====================================================================================================
Sub DisableRefPop()
	IF lgIntFlgMode = parent.OPMD_UMODE Then
		RefPop.innerHTML="<font color=""#777777"">����������</font>"
	ELse 
		RefPop.innerHTML="<A href=""vbscript:OpenRcptNo()"">����������</A>"
	End if

END sub
'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' �Աݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtRcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' �ܾ� 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' û��ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtAdjustAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	
End Sub

'=======================================================================================================
'   Event Name : InputCtrlVal
'   Event Desc :
'=======================================================================================================  
Sub InputCtrlVal(ByVal Row)
	Dim strAcctCd		
	Dim ii
			
	lgBlnFlgChgValue = True
		
	strAcctCd	= Trim(frm1.txtAcctCd.value)		
		
	Call AutoInputDetail(strAcctCd,Trim(frm1.txtDeptCd.value), frm1.txtAdjustDt.text, Row)

End Sub




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.5 Spread Popup method 
' Description : This part declares spread popup method
'=======================================================================================================
'*******************************************************************************************************




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
	Dim indx

	Select Case Trim(UCase(gActiveSpdSheet.Name))
	
		Case "VSPDDATA2"
			Call DeleteHSheet(frm1.hItemSeq.value)
		

			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'�����׸� �׸��� �ʱ�ȭ 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2Color()  
	End Select
End Sub


'=======================================================================================================
'   Event Name : txtAdjustDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtAdjustDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtAdjustDt.Action = 7                
        Call SetFocusToDocument("M")
		Frm1.txtAdjustDt.Focus 
    End If
End Sub
'=======================================================================================================
'   Event Name : txtAdjustDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtAdjustDt_Change() 
    lgBlnFlgChgValue = True
End Sub
'==========================================================================================
'   Event Name : txtAcctCd_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAcctCd_OnChange
 lgBlnFlgChgValue = True
 If Trim(frm1.txtAcctCd.value) <> "" Then
	Call DbQuery4()
 End If
End Sub
'==========================================================================================
'   Event Name : txtAdjustDt_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAdjustDt_OnChange
 lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtAdjustAmt_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAdjustAmt_Change
 lgBlnFlgChgValue = True
 frm1.txtAdjustLocAmt.Text = 0
End Sub

'==========================================================================================
'   Event Name : txtAdjustLocAmt_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAdjustLocAmt_Change
 lgBlnFlgChgValue = True
End Sub

'==========================================================================================
'   Event Name : txtAddesc_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtAdDesc_OnChange
 lgBlnFlgChgValue = True
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>

<!--
'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
'====================================================================================================== -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
			    <TR HEIGHT=23>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A>&nbsp;|&nbsp;<Span id="RefPop"><A HREF="VBSCRIPT:OpenRcptNo()">����������</A></Span></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">		
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>û���ȣ</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtAdJustNo" MAXLENGTH=18 STYLE="TEXT-ALIGN: Left" tag ="12XXXU" ALT="û���ȣ"><IMG align=top name=btnPrpaymNo src="../../image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript:OpenAdJustNo"></TD>								
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=40% >
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ݹ�ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRcptNo" SIZE=20 MAXLENGTH=30 STYLE="TEXT-ALIGN: Left" tag="24" ALT="�����ݹ�ȣ"></TD>
								<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="24" ALT="�ŷ�ó">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="24" ALT="�ŷ�ó��"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�Ա�����</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtRcptDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�Ա�����" tag="24X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>�μ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="24" ALT="ȸ��μ�">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=25 tag="24" ALT="ȸ��μ���"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>������ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" SIZE=20 MAXLENGTH=30 STYLE="TEXT-ALIGN: Left" tag="24" ALT="������ȣ"></TD>
								<TD CLASS=TD5 NOWRAP>�ŷ���ȭ|ȯ��</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDocCur" SIZE=10 MAXLENGTH=4 tag="24NXXU" STYLE="TEXT-ALIGN: left" ALT="�ŷ���ȭ">&nbsp;
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> align ="top" name="txtXchRate" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="ȯ��" tag="24X5Z" id=OBJECT7></OBJECT>');</SCRIPT></TD>
						
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�Աݱݾ�|�ڱ�</TD>
							    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtRcptAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�Աݱݾ�" tag="24X2" ></OBJECT>');</SCRIPT>&nbsp;
							    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtRcptLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�Աݱݾ�(�ڱ�)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>�ܾ�|�ڱ�</TD>
							    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�ܾ�" tag="24X2"></OBJECT>');</SCRIPT>&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�ܾ�(�ڱ�)" tag="24X2" ></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtRcptDesc" SIZE=90 MAXLENGTH=128 tag="24" ALT="����"></TD>
							</TR>						
						</TABLE>
					</TD>
				</TR>
				<TR height="60%">
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>û������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtAdjustDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="û������" tag="22X1" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtAcctCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="22XXXU" ALT="�����ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtAcctCd.value,1)">&nbsp;<INPUT TYPE=TEXT NAME="txtAcctNm" SIZE=25 tag="24" ALT="������"></TD>
							</TR>
<!--							<TR>
								<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtAdDocCur" SIZE=10 MAXLENGTH=4 tag="24NXXU" STYLE="TEXT-ALIGN: left" ALT="�ŷ���ȭ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript: CALL OpenPopup(frm1.txtAdDocCur.value,2)">&nbsp;<INPUT TYPE=TEXT NAME="txtAdDocCurNm" SIZE=20 tag="24" ALT="�ŷ���ȭ��"></TD></TD>
								<TD CLASS=TD5 NOWRAP>ȯ��</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtAdXchRate" CLASS=FPDS140 title=FPDOUBLESINGLE ALT="ȯ��" tag="24X5Z" id=OBJECT7></OBJECT>');</SCRIPT></TD>											
							</TR>-->
							<TR>
								
								<TD CLASS="TD5" NOWRAP>û��ݾ�|�ڱ�</TD>
							    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtAdjustAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="û��ݾ�" tag="22X2" id=OBJECT4></OBJECT>');</SCRIPT>&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtAdjustLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="û��ݾ�(�ڱ�)" tag="21X2" id=OBJECT5></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>����/ȸ����ǥ��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTEMPGlNo" SIZE=18 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="������ǥ��ȣ"> /
								<INPUT TYPE=TEXT NAME="txtGlNo" SIZE=18 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="ȸ����ǥ��ȣ"> </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���</TD>
								<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtAdDesc" SIZE=70 MAXLENGTH=128 tag="21" ALT="���"></TD>
							</TR>	
							<TR HEIGHT="55%">
								<TD WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" name=vspdData2 width="100%" tag="2" TITLE="SPREAD" id=OBJECT6> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA TYPE=hidden Class=hidden name=txtSpread2 tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows2"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hbankcd"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hbanknm"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hbankacct"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hClsAmt"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hClsLocAmt"   tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hConfFg"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hGlNo"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hNoteNo"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctNm"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hSttlAmt"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hSttlLocAmt"  tag="24" TABINDEX="-1">
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

