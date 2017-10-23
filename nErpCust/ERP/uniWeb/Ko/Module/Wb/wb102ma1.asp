<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : VB101MA1
'*  4. Program Name         : Company History(���������̷���ȸ)
'*  5. Program Desc         : ���������̷���ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/12/28
'*  8. Modified date(Last)  : 2004/12/28
'*  9. Modifier (First)     : LSHSAT
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'***********************************************************************k*********************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '��: indicates that All variables must be declared in advance 


'********************************************  1.2 Global ����/��� ����  *********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global ��� ����  ====================================
'==========================================================================================================

Const BIZ_PGM_ID = "WB102MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const C_SHEETMAXROWS    = 100	                                      '��: Visble row

'========================================================================================================= 
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
Dim lgBlnFlawChgFlg	
Dim lgOldRow

Dim lgMpsFirmDate, lgLlcGivenDt											 '��: �����Ͻ� ���� ASP���� �����ϹǷ� Dim 

Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
Dim cboOldVal          
Dim IsOpenPop          
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2        

Dim C_CO_CD
Dim C_FISC_YEAR
Dim C_REP_TYPE
Dim C_CO_NM
Dim C_CO_ADDR
Dim C_OWN_RGST_NO
Dim C_LAW_RGST_NO
Dim C_REPRE_NM
Dim C_REPRE_RGST_NO
Dim C_TEL_NO
Dim C_COMP_TYPE1
Dim C_DEBT_MULTIPLE
Dim C_COMP_TYPE2
Dim C_TAX_OFFICE
Dim C_HOLDING_COMP_FLG
Dim C_IND_CLASS
Dim C_IND_TYPE
DIm C_FOUNDATION_DT
Dim C_REP_TYPE_CD


'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
sub InitSpreadPosVariables()
	
	C_CO_CD = 1
	C_FISC_YEAR = 2
	C_REP_TYPE = 3
	C_CO_NM = 4
	C_CO_ADDR = 5
	C_OWN_RGST_NO = 6
	C_LAW_RGST_NO = 7
	C_REPRE_NM = 8
	C_REPRE_RGST_NO = 9
	C_TEL_NO = 10 
	C_COMP_TYPE1 = 11
	C_DEBT_MULTIPLE = 12
	C_COMP_TYPE2 = 13
	C_TAX_OFFICE = 14
	C_HOLDING_COMP_FLG = 15
	C_IND_CLASS = 16
	C_IND_TYPE = 17
	C_FOUNDATION_DT = 18
	C_REP_TYPE_CD = 19
	
end sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    '-----------------------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""

    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
	lgOldRow = 0

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False


'	frm1.txtCO_CD.value = parent.wgCO_CD
'	frm1.txtco_cd.focus  
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub



'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name :InitComboBox_Five()
'	Description : 
'------------------------------------------------------------------------------------------------------------
Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," dbo.ufn_TB_MINOR('W1018', '" & C_REVISION_YM & "')  "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

 
'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenCompanyInfo()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

Function OpenCompanyInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���� �˾�"						' �˾� ��Ī 
	arrParam(1) = "TB_COMPANY_HISTORY"						' TABLE ��Ī 
	arrParam(2) = strCode							' Code Condition
	arrParam(3) = ""								' Name Cindition
	arrParam(4) = ""								' Where Condition
	arrParam(5) = "����"

    arrField(0) = "Upper(CO_CD)"					' Field��(0)
    arrField(1) = "CO_NM"						' Field��(1)

    arrHeader(0) = "�����ڵ�"						' Header��(0)
    arrHeader(1) = "���θ�"						' Header��(1)

	arrRet = window.showModalDialog("wb101ra1.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCO_CD.focus
	    Exit Function
	Else
		Call SetCompanyInfo(arrRet,iWhere)
	End If	

End Function



'------------------------------------------  SetItemInfo()  -------------------------------------------------
'	Name : SetCostInfo()
'	Description : Popup���� Return�Ǵ� �� setting
'------------------------------------------------------------------------------------------------------------
Function SetCompanyInfo(Byval arrRet,byval iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtCO_CD.focus
			.txtCO_CD.value     = arrRet(0)
			.txtCO_FULLNM.value = arrRet(1)
			
			Call FncQuery
		End If
'		lgBlnFlgChgValue = False
	End With

End Function

Sub txtCO_CD_onChange()	' �����ڵ� ����� 
	Dim arrVal
	
	If Len(frm1.txtCO_CD.Value) > 0 Then
		If CommonQueryRs("CO_NM", " TB_COMPANY_HISTORY " , " CO_CD = '" & frm1.txtCO_CD.Value &"' ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
	    	arrVal				= Split(lgF0, Chr(11))
			frm1.txtCO_FULLNM.Value	= arrVal(0)
		Else
			Call DisplayMsgBox("970000", "x",frm1.txtCO_CD.alt,"x")
			frm1.txtCO_FULLNM.Value	= ""
		End If
	Else
		frm1.txtCO_FULLNM.Value = ""
	End If

End Sub



'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
    If Frm1.vspdData.MaxRows > 0 Then
        Call vspdData_Click(1 , 1)
		Frm1.vspdData.focus
        Set gActiveElement = document.ActiveElement
	End If
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
            
			C_CO_CD						= iCurColumnPos(1)
			C_FISC_YEAR					= iCurColumnPos(2)
			C_REP_TYPE					= iCurColumnPos(3)
			C_CO_NM						= iCurColumnPos(4)
			C_CO_ADDR					= iCurColumnPos(5)
			C_OWN_RGST_NO				= iCurColumnPos(6)
			C_LAW_RGST_NO				= iCurColumnPos(7)
			C_REPRE_NM					= iCurColumnPos(8)
			C_REPRE_RGST_NO				= iCurColumnPos(9)
			C_TEL_NO					= iCurColumnPos(10)
			C_COMP_TYPE1				= iCurColumnPos(11)
			C_DEBT_MULTIPLE				= iCurColumnPos(12)
			C_COMP_TYPE2				= iCurColumnPos(13)
			C_TAX_OFFICE				= iCurColumnPos(14)
			C_HOLDING_COMP_FLG			= iCurColumnPos(15)
			C_IND_CLASS					= iCurColumnPos(16)
			C_IND_TYPE					= iCurColumnPos(17)
			C_FOUNDATION_DT				= iCurColumnPos(18)
			C_REP_TYPE_CD				= iCurColumnPos(19)
    End Select    
End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  

	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    

		.ReDraw = false
		.MaxCols   = C_REP_TYPE_CD + 1                                          ' ��:��: Add 1 to Maxcols
		.Col       = .MaxCols                                                        ' ��:��: Hide maxcols
		.ColHidden = True                                                            ' ��:��:
		.MaxRows = 0

		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 

		Call GetSpreadColumnPos("A")  

		ggoSpread.SSSetEdit     C_CO_CD,					"�����ڵ�",					7
		ggoSpread.SSSetEdit     C_FISC_YEAR,				"�������",					7, 2
		ggoSpread.SSSetEdit     C_REP_TYPE,					"�Ű���",					7
		ggoSpread.SSSetEdit     C_CO_NM,					"���θ�",					15
		ggoSpread.SSSetEdit     C_CO_ADDR,					"���μ�����",				20
		ggoSpread.SSSetEdit     C_OWN_RGST_NO,				"����ڵ�Ϲ�ȣ",			12, 2
		ggoSpread.SSSetEdit     C_LAW_RGST_NO,				"���ε�Ϲ�ȣ",				12, 2
		ggoSpread.SSSetEdit     C_REPRE_NM,					"��ǥ�ڸ�",					10
		ggoSpread.SSSetEdit     C_REPRE_RGST_NO,			"��ǥ���ֹι�ȣ",			13, 2
		ggoSpread.SSSetEdit     C_TEL_NO,					"�������ȭ��ȣ",			13
		ggoSpread.SSSetEdit     C_COMP_TYPE1,				"�߼ұ������",				10
		ggoSpread.SSSetEdit     C_DEBT_MULTIPLE,			"���Ա� ���",				10
		ggoSpread.SSSetEdit     C_COMP_TYPE2,				"�Ϲݹ��� �ش翩��",		15
		ggoSpread.SSSetEdit     C_TAX_OFFICE,				"���Ҽ�����",				10
		ggoSpread.SSSetEdit     C_HOLDING_COMP_FLG,			"����ȸ���ش翩��",			13
		ggoSpread.SSSetEdit     C_IND_CLASS,				"����",						10
		ggoSpread.SSSetEdit     C_IND_TYPE,					"����",						20
		ggoSpread.SSSetEdit     C_FOUNDATION_DT,			"����������",				10, 2
		ggoSpread.SSSetEdit     C_REP_TYPE_CD,					"�Ű���",					10
		
		Call ggoSpread.SSSetColHidden(C_REP_TYPE_CD,C_REP_TYPE_CD,True)	

		Call SetSpreadLock 
		.ReDraw = true
	    
	End With
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub


'========================================================================================================= 
Sub Form_Load()
    Call InitVariables																'��: Initializes local global variables
    Call LoadInfTB19029																'��: Load table , B_numeric_format
	Call AppendNumberPlace("6","4","0")
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitComboBox

    Call SetToolBar("1100000000000111")

    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	frm1.txtco_cd.focus 
	frm1.cboREP_TYPE.value = ""

    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed

'	FncQuery

End Sub

'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'========================================================================================
Function FncQuery() 
    Dim IntRetCD

    FncQuery = False
    Err.Clear

  '-----------------------
    'Check previous data area
    '----------------------- 
'    If lgBlnFlgChgValue = True Then
'		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
'		If IntRetCD = vbNo Then
'			Exit Function
'		End If
'    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

  '-----------------------
    'Query function call area
    '----------------------- 

    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True
End Function


'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables

    Call SetToolbar("1110100000000111")

	frm1.txtCO_CD.focus

    FncNew = True

End Function


'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim strYear,strMonth,strDay
    Dim strYear1,strMonth1,strDay1

	FncSave = False
	Err.Clear

	'-----------------------
	'Precheck area
	'-----------------------
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
	    Exit Function
	End If
	'-----------------------
	'Check content area
	'-----------------------
	If Not chkField(Document, "2") Then                             '��: Check contents area
	   Exit Function
	End If

'	If CompareDateByFormat(frm1.txtFISC_Start_DT.text,frm1.txtFISC_End_DT.text,frm1.txtFISC_Start_DT.Alt,frm1.txtFISC_End_DT.Alt, _
'        	               "970024",frm1.txtFISC_Start_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
'	   frm1.txtFISC_Start_DT.focus
'	   Exit Function
'	End If
  
'	If CompareDateByFormat(frm1.txtHOME_ANY_START_DT.text,frm1.txtHOME_ANY_END_DT.text,frm1.txtHOME_ANY_START_DT.Alt,frm1.txtHOME_ANY_END_DT.Alt, _
'       	               "970024",frm1.txtHOME_ANY_START_DT.UserDefinedFormat,parent.gComDateType, true) = False Then
'	   frm1.txtHOME_ANY_START_DT.focus
'	   Exit Function
'	End If
  
	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False then
		Exit Function
	End If

	FncSave = True
End Function


'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    lgIntFlgMode = parent.OPMD_CMODE											'Indicates that current mode is Crate mode

     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'This function lock the suitable field
    
	lgBlnFlgChgValue = True

    frm1.txtCO_CD_Body.value = ""

    frm1.txtCO_CD_Body.focus
    
End Function


'========================================================================================
Function FncCancel()
     On Error Resume Next
End Function


'========================================================================================
Function FncInsertRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncDeleteRow()
     On Error Resume Next
End Function


'========================================================================================
Function FncPrint()
     On Error Resume Next
    parent.FncPrint()
End Function


'========================================================================================
Function FncPrev()

End Function


'========================================================================================
Function FncNext()

End Function


'========================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_SINGLE)												'��: ȭ�� ���� 
End Function


'========================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")

		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

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
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
	
	With frm1.vspdData 
		.Row = Row	
		.Col = C_CO_CD		: WriteCookie "gCoCd", .Text
		.Col = C_FISC_YEAR	: WriteCookie "gFiscYear", .Text
		.Col = C_REP_TYPE_CD: WriteCookie "gRepType", .Text
	
		Call PgmJump("WB101MA1")
	End With
End Sub

Function PgmJump(Byval pMnuID)
	Dim objConn , PostString
	
	WriteCookie "gActivePgmID",pMnuID
	
	Set objConn = CreateObject("uniConnector.cGlobal") 
	PostString = objConn.GetAspPostString 
	'window.open "../../SessionTrans.asp?" & PostString 
	window.open "../../uniToolbar.Asp?SLX=Y&DPCP=" & pMnuID & "&arg="
End Function

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub


'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtco_cd=" & Trim(frm1.txtco_cd.value)				'��: ���� ���� ����Ÿ 

	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function


'========================================================================================
Function DbQuery()

    Err.Clear

    DbQuery = False
    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal     & "&lgStrPrevKey="		& lgStrPrevKey             '��: Next key tag
    strVal = strVal     & "&txtMaxRows="		& Frm1.vspdData.MaxRows         '��: Max fetched data


	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
 '   Call LayerShowHide(0)
End Function

'========================================================================================
Function DbQueryOk()
	lgIntFlgMode      =  parent.OPMD_UMODE                                               '��: Indicates that current mode is Create mode

	Call SetToolbar("1100000000000111")												'��: Set ToolBar
    Call InitData()
    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

	Frm1.vspdData.focus

End Function

'========================================================================================
Function DbSave() 

    Err.Clear
	DbSave = False

    Dim strVal

    Call LayerShowHide(1) 

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	End With

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()
    frm1.txtCO_CD.value = frm1.txtCO_CD_Body.value 
    lgBlnFlgChgValue = False
    'FncQuery
End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
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
    
   	If lgOldRow <> Row Then
		
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = row
	
		lgOldRow = Row
		  		
	End If
       frm1.vspdData.Row = Row
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If
	
		frm1.vspdData.Col = 1
		frm1.vspdData.Row = NewRow
	
End Sub


'=======================================================================================================
'   Event Name : 
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================

Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
    End If
End Sub




'=======================================================================================================
'   Event Name : Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================

Sub txtFISC_YEAR_Keypress(Key)
    If Key = 13 Then
        FncQuery()
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
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white>�������� �̷���ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCO_CD" MAXLENGTH="10" SIZE=10 ALT ="�����ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenCompanyInfo(frm1.txtco_cd.value,0)"> <INPUT NAME="txtCO_FULLNM" MAXLENGTH="30" SIZE=30 ALT ="���θ�" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>�������</TD>
									<TD CLASS="TD6" NOWRAP>
<script language =javascript src='./js/wb102ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�Ű���</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="11XXXU"></SELECT></TD>
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
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
                        			<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
									<script language =javascript src='./js/wb102ma1_vaSpread1_vspdData.js'></script>
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
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabindex="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" tabindex="-1"></iframe>
</DIV>

</BODY>
</HTML>

