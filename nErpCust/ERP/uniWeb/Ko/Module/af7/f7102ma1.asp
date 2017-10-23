
<%@ LANGUAGE="VBSCRIPT" %>

<!--'======================================================================================================
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : PRERECEIPT
'*  3. Program ID           : f7102ma1
'*  4. Program Name         : ������ û�� 
'*  5. Program Desc         : ������ û�� 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/10/10
'*  8. Modified date(Last)  : 2002/01/10
'*  9. Modifier (First)     : Hee Jung, Kim
'* 10. Modifier (Last)      : Chung Ku, Heo
'* 11. Comment              : �ּ��� �� Ȱ���սô�.
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--'=======================================================================================================
'												1. �� �� �� 
'======================================================================================================= -->
<!--'=======================================================================================================
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
<SCRIPT LANGUAGE="VBScript"		SRC="../ag/AcctCtrl4.vbs">           </SCRIPT>

<SCRIPT LANGUAGE=vbscript>

Option Explicit																	'��: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->

'=======================================================================================================
'                                               1.2 Global ����/��� ����  
'	.Constant�� �ݵ�� �빮�� ǥ��.
'	.���� ǥ�ؿ� ����. prefix�� g�� �����.
'	.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=======================================================================================================

'@PGM_ID
Const BIZ_PGM_ID         = "f7102mb1.asp"					'��: F_PrRcpt_Sttl �� CRUD

Const gIsShowLocal = "Y"
<%
Const gIsShowLocal = "Y"
%>


Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2

Dim C_ItemSeq	
Dim C_AcctCd	
Dim C_ACCT_PB	
Dim C_AcctNm	
Dim C_STTL_AMT	
Dim C_ITEM_LOC_AMT
Dim C_STTL_LOC_AMT
Dim C_SttlDESC
Dim C_DrCRFG


Dim lgStrPrevKeyDtl
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgPrevNo
Dim lgNextNo
Dim lgCurrRow

Dim IsOpenPop	                'Popup
Dim gSelframeFlg

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




'======================================================================================================
' Name : initSpreadPosVariables()
' Description : �׸���(��������) �÷� ���� ���� �ʱ�ȭ 
'=======================================================================================================
Sub initSpreadPosVariables()
	C_ItemSeq	   = 1																'��: Spread Sheet �� Columns �ε��� 
	C_AcctCd	   = 2
	C_ACCT_PB	   = 3
	C_AcctNm	   = 4
	C_STTL_AMT	   = 5
	C_ITEM_LOC_AMT = 6
	C_STTL_LOC_AMT = 7
	C_SttlDESC     = 8
	C_DrCRFG       = 9
End Sub

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
    lgStrPrevKey = 0                            'initializes Previous Key
    lgStrPrevKeyDtl = 0                         'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'======================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=======================================================================================================
Sub SetDefaultVal()
	frm1.txtDocCur.value = parent.gCurrency	
	frm1.txtSttlDt.text  = UniConvDateAToB(dtToday, parent.gServerDateFormat,gDateFormat)
	frm1.txtXchRate.text = 1
End Sub

'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("I","*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE" , "MA") %>
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	
	With frm1
		ggoSpread.Source = .vspdData
		ggoSpread.Spreadinit "V20021126",,parent.gAllowDragDropSpread    
		.vspdData.ReDraw = False    
		
		.vspdData.MaxCols = C_DrCRFG + 1 												'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.vspdData.Col = .vspdData.MaxCols													'������Ʈ�� ��� Hidden Column
		.vspdData.ColHidden = True    
		.vspdData.MaxRows = 0
        
		Call AppendNumberPlace("6","3","0")
		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetFloat  C_ItemSeq     ," No"                , 5,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,2,,,"0","999"
	    ggoSpread.SSSetEdit	  C_AcctCd      ," �����ڵ�"      ,15, ,,18
	    ggoSpread.SSSetButton C_ACCT_PB
	    ggoSpread.SSSetEdit	  C_AcctNm      ," �����ڵ��"    ,20
	    ggoSpread.SSSetFloat  C_STTL_AMT    ," û��ݾ�"      ,17, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_ITEM_LOC_AMT," û��ݾ�(�ڱ�)",17, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_STTL_LOC_AMT," û��ݾ�(�ڱ�)",17, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetEdit	  C_SttlDESC    ," ���"          ,30,,,128
	    ggoSpread.SSSetEdit   C_DrCRFG      ," "    ,5	    
   	  
		Call ggoSpread.SSSetColHidden(C_STTL_LOC_AMT,C_STTL_LOC_AMT,True)
		call ggoSpread.MakePairsColumn(C_AcctCd,C_ACCT_PB)
	    Call ggoSpread.SSSetColHidden(C_DrCRFG,C_DrCRFG,True)	    		
    
		.vspdData.ReDraw		= True
	End With

	Call SetSpreadLock()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1
		ggoSpread.Source = .vspdData
		.vspdData.Redraw = False

		ggoSpread.SpreadLock C_ItemSeq,		-1, C_ItemSeq
  	    ggoSpread.SpreadLock C_AcctCd,		-1, C_AcctCd
  	    ggoSpread.SpreadLock C_ACCT_PB,		-1, C_ACCT_PB
		ggoSpread.SpreadLock C_ACCTNM,		-1, C_ACCTNM
		ggoSpread.SSSetRequired C_STTL_AMT,	-1, -1
		.vspdData.Redraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False

		ggoSpread.SSSetProtected C_ItemSeq,     pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_AcctCd,      pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ACCTNM,      pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_STTL_AMT,    pvStartRow, pvEndRow
		
		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpread2ColorF
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpread2ColorF()
	Dim Row
    With frm1
		ggoSpread.Source = .vspdData2
		.vspdData2.ReDraw = False
		                
		For Row = 1 To .vspdData2.MaxRows
			ggoSpread.SSSetProtected	C_DtlSeq,		Row,	Row
			ggoSpread.SSSetProtected	C_Ctrlcd,		Row,	Row
			ggoSpread.SSSetProtected	C_CtrlNm,		Row,	Row
			ggoSpread.SSSetProtected	C_CtrlValNm,	Row,	Row

			.vspdData2.Col = C_DRFg			
			If (.vspdData2.Text = "C" And .vspdData2.Text <> "") _
                            Or .vspdData2.text = "Y" Or .vspdData2.text = "DC" Then
				ggoSpread.SSSetRequired C_CtrlVal, Row, Row
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
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
     
            C_ItemSeq			= iCurColumnPos(1)
	        C_AcctCd			= iCurColumnPos(2)
	        C_ACCT_PB			= iCurColumnPos(3)
	        C_AcctNm			= iCurColumnPos(4)
	        C_STTL_AMT		    = iCurColumnPos(5)
	        C_ITEM_LOC_AMT	    = iCurColumnPos(6)
	        C_STTL_LOC_AMT	    = iCurColumnPos(7)
	        C_SttlDESC			= iCurColumnPos(8)
	        C_DrCRFG            = iCurColumnPos(9)
	End Select    
End Sub

'=======================================================================================================
'Description : ȸ����ǥ �������� �˾� 
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtSttlGlNo.value)	'ȸ����ǥ ��ȣ 

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'=======================================================================================================
'Description : ������ǥ �������� �˾� 
'=======================================================================================================
Function OpenPopupTempGL()
	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName
	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtSttlTEMPGlNo.value)	'������ǥ ��ȣ 

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
'   Function Name : OpenAcctCd(Byval strCode, Byval iWhere)
'   Function Desc : 
'=======================================================================================================
Function OpenAcctInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	arrParam(0) = "�����ڵ� �˾�"										' �˾� ��Ī 
	arrParam(1) = "A_ACCT,A_ACCT_GP"										' TABLE ��Ī 
	arrParam(2) = Trim(strCode)							 					' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "	' Where Condition
	arrParam(5) = "�����ڵ�"			
	
    arrField(0) = "A_ACCT.ACCT_CD"											' Field��(0)
    arrField(1) = "A_ACCT.ACCT_NM"											' Field��(1)
    arrField(2) = "A_ACCT_GP.GP_CD"											' Field��(2)
    arrField(3) = "A_ACCT_GP.GP_NM"											' Field��(3)
    
    arrHeader(0) = "�����ڵ�"											' Header��(0)
    arrHeader(1) = "�����ڵ��"											' Header��(1)
	arrHeader(2) = "�׷��ڵ�"										' Header��(2)
	arrHeader(3) = "�׷��"										' Header��(3)				
    
	IsOpenPop = True

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Call SetActiveCell(FRM1.vspdData,C_AcctCd,FRM1.vspdData.ActiveRow ,"M","X","X")
		Exit Function
	Else
		With frm1
			.vspdData.Col = C_AcctCd
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_ACCTNM
			.vspdData.Text = arrRet(1)

			Call vspdData_Change(C_AcctCd , .vspdData.Row)		 ' ������ �о�ٰ� �˷��� 
			Call SetActiveCell(.vspdData,C_AcctCd,.vspdData.ActiveRow ,"M","X","X")
		End With
	End If	
End Function

'=========================================================================================================
'	Name : OpenSttlmentNo()
'	Description : Ref ȭ���� call�Ѵ�. : û���ȣ 
'========================================================================================================= 
Function OpenSttlmentNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("f7506ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f7506ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
   
	IsOpenPop = True

	arrParam(0) = frm1.txtSttlmentNo.value				' �˻������� ������� �Ķ���� 

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	    
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then		
		frm1.txtSttlMentNo.focus
		Exit Function
	Else		
		Call SetSttlmentNo(arrRet)
	End If

End Function
'======================================================================================================
'   Function Name : SetSttlmentNo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetSttlmentNo(Byval arrRet)

	frm1.txtSttlmentNo.value	= arrRet(0)
	Call txtSttlMentNo_Change() 
	frm1.txtSttlMentNo.focus
End Function

'========================================================================================================= 
'	Name : OpenRefPreRcptNo()
'	Description : Ref ȭ���� call�Ѵ�. 
'========================================================================================================= 
Function OpenRefPreRcptNo()
	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim IntRetCD
	
	IF lgIntFlgMode = parent.OPMD_UMODE THEN Exit Function
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("f7102ra2")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "f7102ra2", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = ""				' �˻������� ������� �Ķ���� 
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
		Call SetRefPreRcptNo(arrRet)
	End If
End Function

 '------------------------------------------  SetRefPreRcptNo()  ---------------------------------------
'	Name : SetRefPreRcptNo()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function  SetRefPreRcptNo(Byval arrRet)
	lgBlnFlgChgValue = True
	With frm1
		.htxtPrrcptNo.Value	  = arrRet(0)	
		.txtPrRcptNo.value    = arrRet(0)		
		.txtPrrcptDt.text	  = arrRet(1)		
		.txtDeptCd.Value	  = arrRet(2)		
		.txtDeptNm.Value	  = arrRet(3)		
		.txtBpCd.Value		  = arrRet(4)		
		.txtBpNm.Value		  = arrRet(5)		
		.txtRefNo.Value		  = arrRet(6)	
		.txtDocCur.value	  = arrRet(7)		
		.txtXchRate.text	  = arrRet(8)	
		.txtGlNo.value		  = arrRet(9)	
		.txtTempGlNo.value	  = arrRet(10)	
		.txtPrrcptAmt.Text	  = arrRet(11)	
		.txtPrrcptLocAmt.Text = arrRet(12)	
		.txtBalAmt.text		  = arrRet(13)
		.txtBalLocAmt.text	  = arrRet(14)
		.txtPrrcptDesc.value  = arrRet(15)
		
		.txtSttlDocCur.value = Trim(.txtDocCur.value)
		
		Call txtDocCur_OnChange()
		
		If frm1.vspddata.maxrows >0 Then
			Call SetToolbar("1110111100111111")										    '��ư ���� ���� 
		Else
			Call SetToolbar("1110011100111111")										    '��ư ���� ���� 
		End If	
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
'=======================================================================================================
Sub Form_Load()
    Call LoadInfTB19029()   
    Call ggoOper.ClearField(Document, "1")	                                                      'Load table , B_numeric_format
    'Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)

    Call ggoOper.LockField(Document, "N")                                       'Lock  Suitable  Field
    
    Call InitSpreadSheet()                                                        'Setup the Spread sheet
	Call InitCtrlSpread()
	Call InitCtrlHSpread()	    
    Call InitVariables()                                                          'Initializes local global variables
    Call SetDefaultVal()    
    
    Call SetToolbar("11100000000011111")
	
	frm1.txtPrrcptNo.focus
	Set gActiveElement = document.activeElement
	lgBlnFlgChgValue = false

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
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim var1, var2
    FncQuery = False                                                        
    
    Err.Clear                                                               
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData											'��: Preset spreadsheet pointer 
	var1 = ggoSpread.SSCheckChange
	
	ggoSpread.Source = frm1.vspdData2											'��: Preset spreadsheet pointer 
	var2 = ggoSpread.SSCheckChange
	
    If lgBlnFlgChgValue = True or var1 = True or var2 = True Then									'��: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")					'Clear Contents Field
    Call InitVariables()											'Initializes local global variables
    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
	
	Call SetToolbar("11100000000001111")
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'This function check indispensable field
		Exit Function
    End If
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery()													'��: Query db data
           
    FncQuery = True																
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    Dim var1, var2
    
    FncNew = False                                                          
	'-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData											'��: Preset spreadsheet pointer 
	var1 = ggoSpread.SSCheckChange
	
	ggoSpread.Source = frm1.vspdData2											'��: Preset spreadsheet pointer 
	var2 = ggoSpread.SSCheckChange
	
    If lgBlnFlgChgValue = True or var1 = True or var2 = True Then									'��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	'-----------------------
    'Erase condition area
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    
    Call InitVariables()                                                      'Initializes local global variables
    Call SetDefaultVal()        

    Call txtDocCur_OnChange()    
    Call DisableRefPop()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    
    lgBlnFlgChgValue = False        
    
    FncNew = True                                                          
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : 
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncDelete() 
	Dim IntRetCD
    
    FncDelete = False                                                      
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then										'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                               
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")			'Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete																	'��: Delete db data
    
    FncDelete = True         
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim var1, var2
    
    
    FncSave = False                                                         
    
    On Error Resume Next                                                   
    Err.Clear                                                                   
    '-----------------------
    'Precheck area
    '-----------------------
	ggoSpread.Source = frm1.vspdData											'��: Preset spreadsheet pointer 
	var1 = ggoSpread.SSCheckChange
	
	ggoSpread.Source = frm1.vspdData2											'��: Preset spreadsheet pointer 
	var2 = ggoSpread.SSCheckChange
	
    If lgBlnFlgChgValue = false and var1 = false and var2 = false Then									'��: Check If data is chaged
       IntRetCD = DisplayMsgBox("900001","x","x","x")							'No data changed!!
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData											'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then								'��: Check required field(Multi area)
		Exit Function
    End If

    ggoSpread.Source = frm1.vspdData2											'��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then								'��: Check required field(Multi area)
		Exit Function
    End If

    If CheckSpread3 = False then 
		IntRetCD = DisplayMsgBox("110420","x","x","x")							'�����׸��� �ʼ��Է� �����Դϴ�."	
        Exit Function
    End if
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave()																'��: Save db data
    
    FncSave = True                                                       
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy() 
	Dim  IntRetCD
	 
    if frm1.vspdData.MaxRows < 1 then Exit Function
	
	frm1.vspdData.ReDraw = False
		
	ggoSpread.Source = frm1.vspdData
	ggoSpread.CopyRow
	
	cALL MaxSpreadVal(frm1.vspdData, C_ItemSeq , frm1.vspdData.ActiveRow)
	
	Call SetSpreadColor(frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow)
	    
	frm1.vspdData.Col = C_AcctCd
	frm1.vspdData.Text = ""

	frm1.vspdData.Col = C_ACCTNM
	frm1.vspdData.Text = ""

	frm1.vspdData.ReDraw = True
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then Exit Function
    
    With frm1.vspdData
        .Row = .ActiveRow
        .Col = 0
        If .Text = ggoSpread.InsertFlag Then
            .Col = C_ItemSeq
            DeleteHSheet(.Text)
        End if
   
        ggoSpread.Source = frm1.vspdData	
        ggoSpread.EditUndo

		if frm1.vspddata.Maxrows < 1 Then Exit Function

        .Row = .ActiveRow
        .Col = 0
		If .row = 0 then 
			Exit Function
		Else
			If .Text = ggoSpread.InsertFlag Then
			    .Col = C_Acctcd
				If Len(Trim(.Text)) > 0 Then             
					.Col = C_ItemSeq
					frm1.hItemSeq.Value = .Text
					ggoSpread.Source = frm1.vspdData2
					ggoSpread.ClearSpreadData
					Call DbQuery3(.ActiveRow)
				End If	
			Else
			    .Col = C_ItemSeq
			    frm1.hItemSeq.Value = .Text
				ggoSpread.Source = frm1.vspdData2
				ggoSpread.ClearSpreadData
			    Call DbQuery2(.ActiveRow)
			End If
		End If        
    End With
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(Byval pvRowcnt) 
	Dim iCurRowPos
	Dim imRow
    Dim ii
    
	On Error Resume Next                                                          '��: If process fails
    Err.Clear   
	
    FncInsertRow = False                                                         '��: Processing is NG

    If IsNumeric(Trim(pvRowcnt)) Then 
		imRow  = Cint(pvRowcnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
    End If                              
        
	With frm1.vspdData        
		iCurRowPos = .ActiveRow
        .ReDraw = False
        ggoSpread.Source = frm1.vspdData
		ggoSpread.InsertRow ,imRow
		
		For ii = .ActiveRow To  .ActiveRow + imRow - 1
			Call MaxSpreadVal(frm1.vspdData, C_ItemSeq, ii)
		Next  
		.Col = 2																	' �÷��� ���� ��ġ�� �̵�      
		.Row = 	ii - 1
		.Action = 0
		
		.Col = C_DrCRFG
		.Row = ii
		.Text = "CR"		

        Call SetSpreadColor(iCurRowPos + 1, iCurRowPos + imRow)
        .ReDraw = True
	End With        

    Call ggoOper.LockField(Document, "I")									'This function lock the suitable field
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
    
    If frm1.vspdData.maxrows > 0 Then
		If Trim(frm1.hprrcptno.value) = "" Then
			Call SetToolbar("111011110011111")		
		Else
			Call SetToolbar("111111110011111")		
		End If
    End If
	
    Set gActiveElement = document.ActiveElement   
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


'=======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim DelItemSeq

    If frm1.vspdData.MaxRows < 1 Then Exit Function

    With frm1.vspdData    
		.Row = .ActiveRow
		.Col = C_ItemSeq
        DelItemSeq = .Text
        
    	ggoSpread.Source = frm1.vspdData 

   		lDelRows = ggoSpread.DeleteRow
    End With
    
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData
    Call DeleteHsheet(DelItemSeq)
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint() 
    On Error Resume Next
    call parent.FncPrint()                                                 
	    		
	Set gActiveElement = document.activeElement    

End Function

'=====================s==================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function FncPrev() 
    On Error Resume Next                                               
End Function

'=======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function FncNext() 
    On Error Resume Next                                               
End Function

'=======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                          
	    		
	Set gActiveElement = document.activeElement    

End Function

'=======================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: "Will you destory previous data"
	
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
Function DbDelete()     
	DbDelete = false     
     Dim strVal
    With frm1
		.txtFlgMode.value = lgIntFlgMode
		.txtMode.value = parent.UID_M0003
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003						'��: 
		strVal = strVal & "&txtFlgMode=" & Trim(.txtFlgMode.value)					'��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtSttlmentNo=" & Trim(.txtSttlmentNo.value)					'��ȸ ���� ����Ÿ 
		strVal = strVal & "&txtPrRcptNo=" & Trim(.txtPrRcptNo.value)					'��ȸ ���� ����Ÿ 
	End With 

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
	
	Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 
	DbDelete = TRUE      
	
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()												        '���� ������ ���� ���� 
    '----------------------- 
    Call ggoOper.ClearField(Document, "1")                                  'Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  'Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
    
    Call InitVariables()                                                      'Initializes local global variables
    Call SetDefaultVal()        

    Call txtDocCur_OnChange()    
    Call DisableRefPop()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
    
    lgBlnFlgChgValue = False        
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
    Dim strVal

    DbQuery = False                                                             
    Call LayerShowHide(1)
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode="     & parent.UID_M0001					'��: 
			strVal = strVal & "&txtSttlmentNo=" & Trim(.txtSttlmentNo.value)					'��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey="    & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows="      & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID & "?txtMode="     & parent.UID_M0001					'��: 
			strVal = strVal & "&txtSttlmentNo=" & Trim(.txtSttlmentNo.value)					'��ȸ ���� ����Ÿ 
			strVal = strVal & "&lgStrPrevKey="    & lgStrPrevKey
			strVal = strVal & "&lgStrPrevKeyDtl=" & lgStrPrevKeyDtl
			strVal = strVal & "&txtMaxRows="      & .vspdData.MaxRows
		End If
    End With

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
	
	Call RunMyBizASP(MyBizASP, strVal)										    '��: �����Ͻ� ASP �� ���� 

    DbQuery = True                                                              
End Function

'=======================================================================================================
' Function Name : DbQueryOk1
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk1()
     With frm1
		Call SetSpreadLock()
        '-----------------------
        'Reset variables area
        '-----------------------
        lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
        
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
		Call SetToolbar("1111111100111111")										    '��ư ���� ���� 

        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text 
            Call DbQuery2(1)
        End If
    End With
    
    Call txtDocCur_OnChange()
    Call DisableRefPop()
    lgBlnFlgChgValue = False        
End Function

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lngRows 
    Dim lGrpcnt
    Dim strVal 
    Dim strDel
    Dim RowD, GrpCntD, strValD, strItemSEQ	'�����׸� �Ķ���� 

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
    strVal = ""
    strDel = ""
    
    GrpCntD = 1: strValD = ""	'�����׸� �Ķ���� 
    
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
	    For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0
			Select Case .Text
				Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					If .Text = ggoSpread.InsertFlag Then
						strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet�� 2�� �̹Ƿ� ���� 
					Else
						strVal = strVal & "U" & parent.gColSep & lngRows & parent.gColSep				'U=Update
					End If

			        .Col = C_ItemSeq		'1                         
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        strItemSEQ = Trim(.Text)	'ITEM_SEQ 
			        .Col = C_AcctCd		   
			        strVal = strVal & Trim(.Text) & parent.gColSep
			        .Col = C_STTL_AMT	   
			        strVal = strVal & UNIConvNum(.Text,0) & parent.gColSep
			        .Col = C_ITEM_LOC_AMT	
			        strVal = strVal & UNIConvNum(.Text,0) & parent.gColSep
			        .Col = C_STTL_LOC_AMT	
			        strVal = strVal & UNIConvNum(.Text,0) & parent.gColSep
			        .Col = C_SttlDESC		
			        strVal = strVal & Trim(.Text) & parent.gRowSep		        

			        lGrpCnt = lGrpCnt + 1
			        
			        '=======================================================================
			        '2001.06.18 Song,MunGil �����׸� �Է�/������ �ɷ� �����ϰ� ������.
			        '=======================================================================
			        With frm1.vspdData3
						For RowD = 1 To .MaxRows
							.Row = RowD
							.Col = C_DtlSeq
							
							If strItemSEQ = Trim(.Text) Then
								strValD = strValD & "C" & parent.gColSep & RowD & parent.gColSep
								.Col = 1	'2
								strValD = strValD & Trim(.Text) & parent.gColSep
								.Col = 2    '3
								strValD = strValD & Trim(.Text) & parent.gColSep
								.Col = 3    '4
								strValD = strValD & Trim(.Text) & parent.gColSep
								.Col = 5	'5
								strValD = strValD & Trim(.Text) & parent.gRowSep
								
								GrpCntD = GrpCntD + 1
							End If
						Next
					End With
			        
			    Case ggoSpread.DeleteFlag
					strDel = strDel & "D" & parent.gColSep & lngRows & parent.gColSep
					
			        .Col = C_ItemSeq	'2
			        strDel = strDel & Trim(.Text) & parent.gRowSep				        '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 

					lGrpcnt = lGrpcnt + 1             
			End Select
	    Next
	End With
	
	frm1.txtMaxRows.value = lGrpCnt-1										'Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread.value =  strDel & strVal									'Spread Sheet ������ ���� 

	frm1.txtMaxRows3.value = GrpCntD - 1									'Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread3.value  = strValD										'Spread Sheet ������ ���� 

	'���Ѱ����߰� start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'���Ѱ����߰� end
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'���� �����Ͻ� ASP �� ���� 
        
    DbSave = True                                                           
End Function

'=======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData		
	ggoSpread.Source = frm1.vspdData3
	ggoSpread.ClearSpreadData
   
	Call InitVariables()
	Call FncQuery()
End Function




'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.4 User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
'*******************************************************************************************************





'=======================================================================================================
' Function Name : DbQuery2																				
' Function Desc : F_PrRcpt_Sttl_Dtl ��ȸ 
'=======================================================================================================
Function DbQuery2(ByVal Row)
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
	    .vspdData.row = Row
	    .vspdData.col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	    
        If CopyFromData(.hItemSeq.Value) = True Then
			Call SetSpread2ColorF() 
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.STTL_NO, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & .hItemSeq.Value & ", " 
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A , A_ACCT_CTRL_ASSN B , F_PRRCPT_STTL_DTL C , F_PRRCPT_STTL D  "
		
		strWhere =			  " D.PRRCPT_NO  =  " & UCase(FilterVar(.htxtPrrcptNo.value, "''", "S")) & "  "
		strWhere = strWhere & " AND D.STTLMENT_NO =  " & UCase(FilterVar(.hSttlmentNo.value, "''", "S")) & "  "
		strWhere = strWhere & " AND D.STTL_NO = " & .hItemSeq.Value & " "
		strWhere = strWhere & " AND D.STTLMENT_NO    =  C.STTLMENT_NO    "
		strWhere = strWhere & " AND D.PRRCPT_NO    =  C.PRRCPT_NO    "
		strWhere = strWhere & " AND D.STTL_NO  =  C.STTL_NO "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
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
				frm1.vspddata2.row = lngRows	
				frm1.vspddata2.col = C_Tableid 
				If Trim(frm1.vspddata2.text) <> "" Then
					frm1.vspddata2.col = C_Tableid
					strTableid = frm1.vspddata2.text
					frm1.vspddata2.col = C_Colid
					strColid = frm1.vspddata2.text
					frm1.vspddata2.col = C_ColNm
					strColNm = frm1.vspddata2.text	
					frm1.vspddata2.col = C_MajorCd					
					strMajorCd = frm1.vspddata2.text	
					
					frm1.vspddata2.col = C_CtrlVal
					
					strNmwhere = strColid & " =   " & FilterVar(frm1.vspddata2.text, "''", "S") & "  " 
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If				 
					
					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   
						frm1.vspddata2.col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
					End If
				End If								
				
				strVal = strVal & Chr(11) & .hItemSeq.Value 
				
				frm1.vspddata2.col = C_DtlSeq
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_CtrlCd
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_CtrlNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_CtrlVal
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_CtrlPB
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_CtrlValNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_Seq
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_Tableid
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_Colid
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_ColNm
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_Datatype
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_DataLen
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_DRFg
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_MajorCd
				strVal = strVal & Chr(11) & frm1.vspddata2.text
				frm1.vspddata2.col = C_MajorCd + 1
				strVal = strVal & Chr(11) & frm1.vspddata2.text																				
				strVal = strVal & Chr(11) & Chr(12)									
			Next					

			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal	
		End If 		
       
		Call SetSpread2ColorF() 	
		
	End With
	
	Call LayerShowHide(0)
	
	frm1.vspdData2.ReDraw = True
	
	DbQuery2 = True
End Function

Sub PrevspdDataRestore(pActiveSheetName)

	Dim indx, indx1

	For indx = 0 to frm1.vspdData.MaxRows
        frm1.vspdData.Row    = indx
        frm1.vspdData.Col    = 0
		
		If frm1.vspdData.Text <> "" Then
			Select Case frm1.vspdData.Text			
				Case ggoSpread.InsertFlag					
					frm1.vspdData.Col = C_ItemSeq					
					Call DeleteHsheet(frm1.vspdData.Text)					
				Case ggoSpread.UpdateFlag		
					For indx1 = 0 to frm1.vspdData3.MaxRows					
						frm1.vspdData3.Row = indx1
						frm1.vspdData3.Col = 0
						Select Case frm1.vspdData3.Text 
							Case ggoSpread.UpdateFlag
								frm1.vspdData.Col = C_ItemSeq
								frm1.vspdData3.Col = 1					
								If UCase(Trim(frm1.vspdData.Text)) = UCase(Trim(frm1.vspdData3.Text)) Then
									Call DeleteHsheet(frm1.vspdData.Text)										
									Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtPrrcptNo.Value) 
								End If
						End Select
					Next
					'ggoSpread.Source = frm1.vspdData					
					'ggoSpread.EditUndo
					
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.txtPrrcptNo.Value) 

					'ggoSpread.Source = frm1.vspdData
					'ggoSpread.EditUndo

			End Select
			
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName

End Sub


'=======================================================================================================
' Function Name : DbQueryOk2
' Function Desc : DbQuery2�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk2()
	Call SetSpread2ColorF()
    Call txtDocCur_OnChange()
    
    lgBlnFlgChgValue = False        
End Function

'=======================================================================================================
' Function Name : CheckSpread3
' Function Desc : ����ÿ�  �ʼ����� check �ϱ����� ȣ��Ǵ� Function
'=======================================================================================================
Function CheckSpread3()
	Dim indx, jj
	Dim tmpDrCrFG

	CheckSpread3 = False

	With frm1
		For jj = 1 To .vspdData.MaxRows
			.vspdData.row = jj
			.vspdData.col = C_DrCRFG
			tmpDrCrFG = Left(.vspddata.Text,1)

	 		For indx = 1 to .vspdData3.MaxRows
			    .vspdData3.Row = indx
			    .vspdData3.Col = 14

			    If (tmpDrCrFG = .vspddata3.Text) Or .vspddata3.Text = "DC" Then
  					.vspdData3.Col = 5
					If Trim(.vspdData3.Text) = "" Then
						Exit Function
			  		End If
			    End If
			Next
		Next	
	End With

	CheckSpread3 = True
End Function

'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	End If
End Sub

'==========================================================================================
'   Event Name : txtSttlDocCur_onChange
'   Event Desc : 
'==========================================================================================
Sub txtSttlDocCur_onChange()
    lgBlnFlgChgValue = True
    If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtSttlDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
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
		
		RefPop.innerHTML="<A href=""vbscript:OpenRefPreRcptNo()"">����������</A>"
	End if

END sub
'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()
	With frm1
		' �����ݾ� 
		ggoOper.FormatFieldByObjectOfCur .txtPrrcptAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
		' �ܾ� 
		ggoOper.FormatFieldByObjectOfCur .txtBalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		ggoSpread.Source = frm1.vspdData
		' û��ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_STTL_AMT,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
	End With
End Sub

'=======================================================================================================
'   Event Name : InputCtrlVal
'   Event Desc :
'=======================================================================================================  
Sub InputCtrlVal(ByVal Row)
	Dim strAcctCd		
	Dim ii
			
	lgBlnFlgChgValue = True
		
	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.Col = C_AcctCd
	frm1.vspdData.Row = Row		
	strAcctCd	= Trim(frm1.vspdData.text)		
		
		
	Call AutoInputDetail(strAcctCd,Trim(frm1.txtDeptCd.value), frm1.txtSttlDt.text, Row)
	For ii = 1 To frm1.vspdData2.MaxRows
		frm1.vspddata2.col = C_CtrlVal
		frm1.vspddata2.row = ii
					
		If Trim(frm1.vspddata2.text) <> "" Then
			Call CopyToHSheet2(frm1.vspdData.ActiveRow,ii)			 			
		End if
	Next
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
		Case "VSPDDATA"
		    Call PrevspdDataRestore(gActiveSpdSheet) 
			For indx = 0 To frm1.vspdData.MaxRows
				frm1.vspdData.Row = indx
				frm1.vspdData.Col = 0
				If frm1.vspdData.Text = ggoSpread.InsertFlag Then
					frm1.vspdData.Col = C_ItemSeq
					Call DeleteHSheet(frm1.vspdData.Text)  
				End If
			Next

			ggoSpread.Source = frm1.vspdData
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet()
			Call ggoSpread.ReOrderingSpreadData()
		Case "VSPDDATA2"
		    Call PrevspdDataRestore(gActiveSpdSheet)
			For indx = 0 To frm1.vspdData.MaxRows
				frm1.vspdData.Row = indx
				frm1.vspdData.Col = 0
				If frm1.vspdData.Text = ggoSpread.InsertFlag Then
					frm1.vspdData.Col = C_AcctCd
					frm1.vspdData.Text = ""
					frm1.vspdData.Col = C_AcctNm
					frm1.vspdData.Text = ""
					frm1.vspdData.Col = C_ItemSeq
					Call DeleteHSheet(frm1.vspdData.Text)
				End If
			Next

			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.RestoreSpreadInf()
			Call InitCtrlSpread()			'�����׸� �׸��� �ʱ�ȭ 
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2Color()  
	End Select
End Sub

Sub PrevspdData2Restore(pActiveSheetName)

	Dim indx, indx1

	For indx = 0 to frm1.vspdData2.MaxRows
        frm1.vspdData2.Row    = indx
        frm1.vspdData2.Col    = 0

		If frm1.vspdData2.Text <> "" Then
			Select Case frm1.vspdData2.Text
				Case ggoSpread.InsertFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData	
					        ggoSpread.EditUndo							
						End If
					Next
				Case ggoSpread.UpdateFlag
					frm1.vspdData2.Col = C_HItemSeq
					For indx1 = 0 to frm1.vspdData.MaxRows
						frm1.vspdData.Row = indx1
						frm1.vspdData.Col = C_ItemSeq
						If frm1.vspdData.Text = frm1.vspdData2.Text Then
							Call DeleteHsheet(frm1.vspdData.Text)
							ggoSpread.Source = frm1.vspdData
							ggoSpread.EditUndo
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.txtPrrcptNo.Value) 
						End If
					Next

				Case ggoSpread.DeleteFlag

			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName

End Sub


Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)

	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal

	on Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	With frm1

		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text
	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  THEN " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  END , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("DC", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  THEN " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  THEN " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " END	, " & strItemSeq & ", "	
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD, '')), CHAR(8) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_GL_DTL C (NOLOCK), A_GL_ITEM D (NOLOCK) "
		
		strWhere =			  " D.GL_NO = " & FilterVar(UCase(pInvalue1), "''", "S")   
		strWhere = strWhere & " AND D.ITEM_SEQ = " & strItemSeq & " "
		strWhere = strWhere & " AND D.GL_NO  =  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then
			arrTempRow =  Split(lgF2By2, Chr(12))
			For Indx1 = 0 To Ubound(arrTempRow) - 1
				arrTempCol = split(arrTempRow(indx1), Chr(11))
				If Trim(arrTempCol(8)) <> "" Then
					strTableid = arrTempCol(8)
					strColid   = arrTempCol(9)
					strColNm   = arrTempCol(10)
					strMajorCd = arrTempCol(15)
					
					strNmwhere = strColid & " =   " & FilterVar(arrTempCol(C_CtrlVal), "''", "S") & "  " 

					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " AND MAJOR_CD =  " & FilterVar(strMajorCd, "''", "S") & "  "
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						arrVal = Split(lgF0, Chr(11))
						arrTempCol(6) = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & strItemSeq
				strVal = strVal & Chr(11) & arrTempCol(1)
				strVal = strVal & Chr(11) & arrTempCol(2)
				strVal = strVal & Chr(11) & arrTempCol(3)
				strVal = strVal & Chr(11) & arrTempCol(4)
				strVal = strVal & Chr(11) & arrTempCol(5)
				strVal = strVal & Chr(11) & arrTempCol(6)
				strVal = strVal & Chr(11) & arrTempCol(7)
				strVal = strVal & Chr(11) & arrTempCol(8)
				strVal = strVal & Chr(11) & arrTempCol(9)
				strVal = strVal & Chr(11) & arrTempCol(10)
				strVal = strVal & Chr(11) & arrTempCol(11)
				strVal = strVal & Chr(11) & arrTempCol(12)
				strVal = strVal & Chr(11) & arrTempCol(13)
				strVal = strVal & Chr(11) & arrTempCol(15)
				strVal = strVal & Chr(11) & Indx1 + 1
				strVal = strVal & Chr(11) & Chr(12)
			Next
			ggoSpread.Source = .vspdData3
			ggoSpread.SSShowData strVal	
		End If 		

		If Row = CurrRow Then
			Call CopyFromData (strItemSeq)
		End If

		Call LayerShowHide(0)
		Call RestoreToolBar()
		
'		Call SetSpread2Color()
		

	End With

	If Err.number = 0 Then
		fncRestoreDbQuery2 = True
	End If

'	Set gActiveElement = document.ActiveElemen

End Function



'*******************************************************************************************************
'=======================================================================================================
' Area Name   : 4.5.6 Spread OCX Tag Event
' Description : This part declares Spread OCX Tag Event
'=======================================================================================================
'*******************************************************************************************************





'=======================================================================================================
'   Event Name : vspdData_onfocus
'   Event Desc :
'=======================================================================================================
Sub vspdData_onfocus()

End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
    
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData
    
    ggoSpread.Source = frm1.vspdData
	frm1.vspddata.Row = frm1.vspddata.ActiveRow

	If Row <= 0 then
	    Exit Sub
    End if

	If frm1.vspdData.ActiveCol = C_Acct_PB then
	    Exit Sub
    End if
 
  	frm1.vspdData.Col = C_AcctCd
	
    If Len(frm1.vspdData.Text) > 0 Then

	Else
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	End If	
End Sub

'======================================================================================================
'   Event Name :vspddata_DblClick
'   Event Desc :
'======================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
     Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
    End If     
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
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
'   Event Name : vspdData_LeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
        With frm1        
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq
            .hItemSeq.value = .vspdData.Text
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.ClearSpreadData
        End With
        
        lgCurrRow = NewRow
       
        Call DbQuery2(lgCurrRow)
    End If
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
    Dim strData
    
    ggoSpread.Source = frm1.vspdData
        
	With frm1.vspdData        
        If Row > 0 And Col = C_ACCT_PB Then
            .Col = C_AcctCd
            .Row = Row                                   
            Call OpenAcctInfo(.Text)
        End If    
    End With
End Sub

'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )
                
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    frm1.vspdData.Col = 0
    
    If Col = C_AcctCd And frm1.vspdData.Text = ggoSpread.InsertFlag Then
        frm1.vspdData.Col = C_ItemSeq
        frm1.hItemSeq.value = frm1.vspdData.Text
        frm1.vspdData.Col = C_AcctCd
        
        If Len(frm1.vspdData.Text) > 0 Then
		    frm1.vspdData.Row = Row
		    frm1.vspdData.Col = 1	
			DeleteHsheet frm1.vspdData.Text
    
            Call DbQuery3 (Row)
            Call InputCtrlVal(Row)	
            Call SetSpread2ColorF()
        End If 
    ElseIf Col = C_STTL_AMT  Then
		frm1.vspddata.row = frm1.vspddata.activerow
		frm1.vspddata.col = C_ITEM_LOC_AMT
		frm1.vspddata.text = 0
    
		frm1.vspddata.row = frm1.vspddata.activerow
		frm1.vspddata.col = C_STTL_LOC_AMT
		frm1.vspddata.text = 0
    End If    
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'======================================================================================================
'   Event Name :vspddata_KeyPress
'   Event Desc :
'=======================================================================================================
Sub vspddata_KeyPress(KeyAscii )
     lgBlnFlgChgValue = True                                            'Indicates that value changed
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim intListGrvCnt 
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'=======================================================================================================
'   Event Name : txtSttlDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtSttlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtSttlDt.Action = 7   
        Call SetFocusToDocument("M")
		Frm1.txtSttlDt.Focus                     
    End If
End Sub
'=======================================================================================================
'   Event Name : txtSttlDt_Change()
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub  txtSttlDt_Change() 
	Dim iRows
    If frm1.vspddata.maxrows >0 Then
		For iRows = 1 To frm1.vspddata.maxrows
			frm1.vspddata.row = iRows
			frm1.vspddata.col = C_ITEM_LOC_AMT
			frm1.vspddata.text = 0
    		frm1.vspddata.col = C_STTL_LOC_AMT
			frm1.vspddata.text = 0
		Next
	End If	
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtSttlMentNo_Change()
'   Event Desc : 
'=======================================================================================================
Sub  txtSttlMentNo_Change() 
    frm1.hSttlmentNo.value = Trim(frm1.txtSttlMentNo.value)
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  --> 
</HEAD>

<!--'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A>&nbsp;|&nbsp;<Span id="RefPop"><A href="vbscript:OpenRefPreRcptNo()">����������</A></Span></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
		<!--ù��° TAB  -->
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>û���ȣ</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtSttlMentNo" MAXLENGTH=18 tag ="12XXXU" ALT="û���ȣ"><IMG align=top name=btnPrpaymNo src="../../../CShared/image/btnPopup.gif"  TYPE="BUTTON" onclick="vbscript:OpenSttlmentNo"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH="100%" HEIGHT=30% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ݹ�ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPrRcptNo" SIZE=20 MAXLENGTH=20 tag="24" ALT="�����ݹ�ȣ"></TD>
								<TD CLASS="TD5" NOWRAP>�Ա�����</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtPrrcptDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�Ա�����" tag="24" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="24" ALT="�ŷ�ó">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="24" ALT="�ŷ�ó��"></TD>
								<TD CLASS="TD5" NOWRAP>�μ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="24" ALT="�μ�">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=25 tag="24" ALT="�μ���"></TD>
								
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" TYPE="Text" SIZE=10 tag="24" ></TD>
								<TD CLASS="TD5" NOWRAP>ȯ��</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 80px" title=FPDOUBLESINGLE ALT="ȯ��" tag="24X5Z" id=fpDoubleSingle1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>ȸ����ǥ��ȣ</TD>
		                        <TD CLASS="TD6" NOWRAP><INPUT NAME="txtGlNo" ALT="ȸ����ǥ��ȣ" TYPE="Text" SIZE=15 STYLE="TEXT-ALIGN: Left" tag="24" ></TD>
		                        <TD CLASS="TD5" NOWRAP>������ǥ��ȣ</TD>
		                        <TD CLASS="TD6" NOWRAP><INPUT NAME="txtTempGlNo" ALT="������ǥ��ȣ" TYPE="Text" SIZE=15 STYLE="TEXT-ALIGN: Left" tag="24" ></TD>
		                    </TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ݾ�|�ڱ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtPrrcptAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�" tag="24X2" id=fpDoubleSingle2></OBJECT>');</SCRIPT>&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtPrrcptLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�����ݾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle3></OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS="TD5" NOWRAP>�ܾ�|�ڱ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�ܾ�" tag="24X2" id=fpDoubleSingle8></OBJECT>');</SCRIPT>&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtBalLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="�ܾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle9></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>���</TD>
								<TD CLASS="TD656" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtPrrcptDesc" SIZE=90 MAXLENGTH=128 tag="24" ALT="���"></TD>
							</TR>
						</TABLE>
					</TD>
					</TR>
					<TR>
					<TD WIDTH="100%" HEIGHT=70% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>û������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtSttlDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="û������" tag="22X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>û��ݾ�|�ڱ�</TD>
							    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSttlAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="û��ݾ�" tag="24X2" id=OBJECT4></OBJECT>');</SCRIPT>&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSttlLocAmt CLASS=FPDS140 title=FPDOUBLESINGLE ALT="û��ݾ�(�ڱ�)" tag="24X2" id=OBJECT5></OBJECT>');</SCRIPT></TD>
								<!--<TD CLASS=TD5 NOWRAP>ȯ��</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtSttlXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 80px" title=FPDOUBLESINGLE ALT="ȯ��" tag="24X5Z" id=fpDoubleSingle1></OBJECT>');</SCRIPT></TD>
							--></TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������ǥ��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSttlTEMPGlNo" SIZE=18 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="������ǥ��ȣ"></TD>
								<TD CLASS=TD5 NOWRAP>ȸ����ǥ��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSttlGlNo" SIZE=18 MAXLENGTH=20 STYLE="TEXT-ALIGN: Left" tag="24XXXU" ALT="ȸ����ǥ��ȣ"> </TD>
							</TR>
							<TR>
								<TD WIDTH="100%" HEIGHT=50% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> tag="2" HEIGHT="100%" name=vspdData width="100%" id=fpSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						    <TR>
								<TD WIDTH="100%" HEIGHT=50% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> Tag="2" HEIGHT="100%" name=vspdData2 width="100%" id=fpSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
			<IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread    tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3   tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"  tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtFlgMode"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtPrrcptNo"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hItemSeq"		 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"	 tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hSttlmentNo"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtRefNo"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtSttlDocCur"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 tag="2" width="100%" TABINDEX="-1" id=OBJECT1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

