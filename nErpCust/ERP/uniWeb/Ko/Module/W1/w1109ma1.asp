<%@ LANGUAGE="VBScript" CODEPAGE=949%>
<% Option Explicit %>
<% session.CodePage=949 %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : Tax
'*  2. Function Name        : 
'*  3. Program ID           : W1109MA1
'*  4. Program Name         : ��3ȣ��3(3) �μӸ���-����������� 
'*  5. Program Desc         : ��3ȣ��3(3) �μӸ���-����������� ���/��ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2004/12/29
'*  8. Modified date(Last)  : 2004/12/30
'*  9. Modifier (First)     : LSHSAT
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'
'���İ� ���� 
'      22.����� >0 �ΰ�� 47ȣ(��) (117)�� �� < 0 �̸� ����	����	WC0024	���μ� ��47ȣ(��) �ֿ���������� ������� ��ȸ����ݾ��� �������� �ʽ��ϴ�.	%1�� %2�� %3�� �������� �ʽ��ϴ�.
'***********************************************************************k*********************** -->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  �α����� ������ �����ڵ带 ����ϱ� ����  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                '��: indicates that All variables must be declared in advance 


'********************************************  1.2 Global ����/��� ����  *********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->

'============================================  1.2.1 Global ��� ����  ====================================
'==========================================================================================================

Const BIZ_PGM_ID 		= "W1109MB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 		= "W1109MB2.asp"											 '��: �����Ͻ� ���� ASP�� 
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

Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

Dim C_GP_CD
Dim C_PAR_GP_CD
Dim C_GP_NM
Dim C_FORM_REP_NO
Dim C_AMT
Dim C_SUM_FG
Dim C_GP_LVL


'============================================  �ʱ�ȭ �Լ�  ====================================
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
sub InitSpreadPosVariables()
	
	C_GP_CD = 1
	C_PAR_GP_CD = 2
	C_GP_NM = 3
	C_FORM_REP_NO = 4
	C_AMT = 5
	C_SUM_FG = 6
	C_GP_LVL = 7
	
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

    lgRefMode = False

End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub



'------------------------------------------  OpenCalType()  -------------------------------------------------
'	Name :InitComboBox()
'	Description : 
'------------------------------------------------------------------------------------------------------------
Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub


'============================================  �׸��� �Լ�  ====================================
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
            
			C_GP_CD						= iCurColumnPos(1)
			C_PAR_GP_CD					= iCurColumnPos(2)
			C_GP_NM						= iCurColumnPos(3)
			C_FORM_REP_NO				= iCurColumnPos(4)
			C_AMT						= iCurColumnPos(5)
			C_SUM_FG					= iCurColumnPos(6)
			C_GP_LVL					= iCurColumnPos(7)
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
		.MaxCols   = C_GP_LVL + 1                                          ' ��:��: Add 1 to Maxcols
		.Col       = .MaxCols                                                        ' ��:��: Hide maxcols
		.ColHidden = True                                                            ' ��:��:
		.MaxRows = 0

		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 

		Call GetSpreadColumnPos("A")  

		ggoSpread.SSSetEdit     C_GP_CD,				"�����ڵ�",				10
		ggoSpread.SSSetEdit     C_PAR_GP_CD,			"���������ڵ�",				10
		ggoSpread.SSSetEdit     C_GP_NM,				"��������",				25
		ggoSpread.SSSetEdit     C_FORM_REP_NO,			"�ڵ�",					7, 2
		ggoSpread.SSSetFloat    C_AMT,					"�ݾ�",					15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec,  "1", True,   "Z"
		ggoSpread.SSSetEdit		C_SUM_FG,				"��꿩��",				10
		ggoSpread.SSSetEdit		C_GP_LVL,				"����",					5

		Call ggoSpread.SSSetColHidden(C_GP_CD,C_GP_CD,True)	
		Call ggoSpread.SSSetColHidden(C_PAR_GP_CD,C_PAR_GP_CD,True)	
		Call ggoSpread.SSSetColHidden(C_SUM_FG,C_SUM_FG,True)	
		Call ggoSpread.SSSetColHidden(C_GP_LVL,C_GP_LVL,True)	

		Call SetSpreadLock 
		.ReDraw = true
	    
	End With
	
	
    
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
'      ggoSpread.Source = frm1.vspdData
'      ggoSpread.SpreadLockWithOddEvenRowColor()

   With frm1

    .vspdData.ReDraw = False
    
    ggoSpread.SpreadLock C_GP_NM, -1, C_GP_NM
    ggoSpread.SpreadLock C_FORM_REP_NO, -1, C_FORM_REP_NO
	ggoSpread.SSSetProtected C_AMT, 1, 1
	ggoSpread.SSSetProtected C_AMT, 5, 5
	ggoSpread.SSSetProtected C_AMT, 10, 10
	ggoSpread.SSSetProtected C_AMT, 28, 28
	ggoSpread.SSSetProtected C_AMT, 31, 31
	ggoSpread.SSSetProtected C_AMT, 34, 34

    'ggoSpread.SSSetRequired C_SUMMARY_DESC, -1, -1
    'ggoSpread.SpreadLock C_COMPANY_NM, -1, C_COMPANY_NM
    'ggoSpread.SSSetRequired C_STOCK_RATE, -1, -1
    'ggoSpread.SSSetRequired C_ACQUIRE_AMT, -1, -1
    'ggoSpread.SSSetRequired C_COMPANY_TYPE, -1, -1
    'ggoSpread.SpreadLock C_OWN_RGST_NO, -1, C_OWN_RGST_NO
	'ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With

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


'============================================  ��ȸ���� �Լ�  ====================================
Sub CheckFISC_DATE()	' ��û������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.
	Dim sFiscYear, sRepType, sCoCd, sFISC_START_DT, sFISC_END_DT, datMonCnt, i, datNow
	
	sCoCd		= frm1.txtCO_CD.value
	sFiscYear	= frm1.txtFISC_YEAR.text
	sRepType	= frm1.cboREP_TYPE.value
	
	call CommonQueryRs("FISC_START_DT, FISC_END_DT"," TB_COMPANY_HISTORY "," CO_CD= '" & sCoCd & "' AND FISC_YEAR='" & sFiscYear & "' AND REP_TYPE='" & sRepType & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    If lgF0 <> "" Then 
		sFISC_START_DT = CDate(lgF0)
	Else
		sFISC_START_DT = ""
	End if

    If lgF1 <> "" Then 
		sFISC_END_DT = CDate(lgF1)
	Else
		sFISC_END_DT = ""
	End if
	
End Sub


'============================================  �� �Լ�  ====================================

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

    Call SetToolBar("1100100000010111")

	Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
	Call ggoOper.FormatDate(frm1.txtFISC_YEAR_Body, parent.gDateFormat,3)

    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    
    Call ggoOper.ClearField(Document, "2")
    Call InitData
    Call FncQuery()
    
End Sub

'==========================================================================================
'==========================================================================================
Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call CheckFISC_DATE
End Sub



'============================================  �������� �Լ�  ====================================

'========================================================================================
Function FncQuery() 
    Dim IntRetCD

    FncQuery = False
    Err.Clear

  '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

  '-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables

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
    Call InitData

    Call SetToolBar("1110100000011111")


	Call DbQuery2

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
	
    FncSave = False                                                         
    
    Err.Clear                                                               
    'On Error Resume Next                                                   

	'-----------------------
	'Condition copy to Check Field
	'-----------------------
	If Not chkField(Document, "1") Then                             '��: Check indispensable field
	   Exit Function
	End If
	Frm1.txtFISC_YEAR_Body.Value = Frm1.txtFISC_YEAR.Text
	Frm1.txtREP_TYPE_Body.Value = Frm1.cboREP_TYPE.Value
'	Frm1.txtBS_PL_FG.Value
	

	'-----------------------
	'Check content area
	'-----------------------
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR     '��: Check contents area
       Exit Function
    End If

	If Not chkField(Document, "2") Then                             '��: Check contents area
	   Exit Function
	End If

<%  '-----------------------
    'Precheck area
    '----------------------- %>
	If lgBlnFlgChgValue = False Then
	    IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
	    Exit Function
	End If
    

	'-----------------------
	' �ʼ��Է� �ݾ� Ȯ��																								
	'-----------------------
	If SubMakeSum <> True Then
		Exit Function
	End If

	If Verification = False Then Exit Function

<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function


' ---------------------- ���ĳ� ���� -------------------------
Function  Verification()
	Dim iSchRow, iSumAmt
	
	Verification = False

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "32", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = UNICdbl(Frm1.vspdData.text)

	'34 <= 0	����	WC0007	�Էµ� DATA�� �����ϴ�.
    If iSumAmt <= 0 Then
        Call DisplayMsgBox("WC0007", "X", "X", "X")                          
        Exit Function
    End If

	Verification = True	
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

'    frm1.txtCO_CD_Body.value = ""

'    frm1.txtCO_CD_Body.focus
    
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
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    ElseIf lgPrevNo = "" then
		Call DisplayMsgBox("900011", "X", "X", "X")
	End IF

    response.write lgPrevNo

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
    strVal = strVal & "&txtco_cd =" & lgPrevNo

	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
Function FncNext()
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						  '��: �����Ͻ� ó�� ASP�� ���°� 
    strVal = strVal & "&txtco_cd=" & lgNextNo

	Call RunMyBizASP(MyBizASP, strVal)
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
End Sub


'============================================  DB �＼�� �Լ�  ====================================

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'��: ���� ���� ����Ÿ 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&txtBS_PL_FG=" & Trim(frm1.txtBS_PL_FG.value)				'��: ��ȸ ���� ����Ÿ 

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


	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery = True
 '   Call LayerShowHide(0)
End Function

'========================================================================================
Function DbQuery2()

    Err.Clear

    DbQuery2 = False
    Call LayerShowHide(1)
    Dim strVal

    strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'��: ��ȸ ���� ����Ÿ 


	Call RunMyBizASP(MyBizASP, strVal)

    DbQuery2 = True
 '   Call LayerShowHide(0)
End Function

'========================================================================================
Function DbQueryOk()
	lgIntFlgMode      =  parent.OPMD_UMODE                                               '��: Indicates that current mode is Create mode

	Call SetToolbar("1101100000010111")												'��: Set ToolBar
    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
    Call SetSpreadLock

	Frm1.vspdData.focus

End Function

'========================================================================================
Function DbQueryOk2()
	lgIntFlgMode      =  parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
    Call SetSpreadLock

	Frm1.vspdData.focus

End Function

'========================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	With Frm1
		For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
			strVal = strVal & "C"  &  Parent.gColSep

            .vspdData.Col = C_GP_CD				: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
            .vspdData.Col = C_GP_NM     	    : strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
            .vspdData.Col = C_FORM_REP_NO		: strVal = strVal & Trim(.vspdData.Text) &  Parent.gColSep
            .vspdData.Col = C_AMT				: strVal = strVal & Trim(.vspdData.Text) &  Parent.gRowSep

            lGrpCnt = lGrpCnt + 1
                    
       Next
		.txtMode.value        =  Parent.UID_M0002
		.txtFlgMode.value     = lgIntFlgMode
		.txtSpread.value      = strVal
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True
End Function

'========================================================================================
Function DbSaveOk()
''    frm1.txtCO_CD.value = frm1.txtCO_CD_Body.value 
    lgBlnFlgChgValue = False
    Call FncQuery
End Function


'============================================  �̺�Ʈ �Լ�  ====================================

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
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub


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
       
    lgBlnFlgChgValue = True

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	
	Call SubMakeSum
	
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================

Sub vspdData_Click(ByVal Col, ByVal Row)

	'Call SetPopupMenuItemInf("0000111111")       

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
    If Row = 0 Then
       Exit Sub
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



'============================== �׸��� ����Ÿ ó�� ���� �Լ�  ========================================


'========================================================================================================
'   Event Name : SubMakeSum
'   Event Desc : This function is Sum Amt with spread sheet
'	FORM_REP_NO	:	01		( 02 + 03 - 04 ) �� ����Ͽ� �Է���.
'					02 ~ 04		����ڰ� ���ڷ� �����Է���.
'					05		( 06 + 07 + 08 ) �� ����Ͽ� �Է���.
'					06 ~ 09		����ڰ� ���ڷ� �����Է���.
'					10		11 ~ 25 ������ �հ踦 �Է���.
'					11 ~ 25		����ڰ� ���ڷ� �����Է���.
'					26		( 01 + 05 + 09 + 10 ) �� ����Ͽ� �Է���.
'					27 ~ 28		����ڰ� ���ڷ� �����Է���.
'					29		( 26 + 27 + 28 ) �� ����Ͽ� �Է���.
'					30		����ڰ� ���ڷ� �����Է���.
'					31		����ڰ� ���ڷ� �����Է���.
'					32		( 29 - 30 - 31 ) �� ����Ͽ� �Է���.
'========================================================================================================
Function SubMakeSum()
	Dim iSchRow
	Dim iSumAmt
	DIm iRowCnt

	iSumAmt = 0
	
	'01		( 02 + 03 - 04 ) �� ����Ͽ� �Է���.
	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "02", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "03", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "04", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt - UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "01", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	Frm1.vspdData.text  = iSumAmt

	'05		( 06 + 07 + 08 ) �� ����Ͽ� �Է���.
	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "06", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "07", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "08", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "05", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	Frm1.vspdData.text  = iSumAmt


	'10		11 ~ 25 ������ �հ踦 �Է���.
	iSumAmt = 0
	For iRowCnt = 11 to 25
		iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, Right("0" + CStr(iRowCnt), 2), 0)
	   	Frm1.vspdData.Row = iSchRow
		Frm1.vspdData.Col = C_AMT
		iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)
	Next

	' -- 2006.03 ���� 33, 34 �ڵ� �߰� 
	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "33", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)
	
	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "34", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)
	

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "10", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	Frm1.vspdData.text  = iSumAmt

	'26		( 01 + 05 + 09 + 10 ) �� ����Ͽ� �Է���.
	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "01", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "05", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "09", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "26", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	Frm1.vspdData.text  = iSumAmt


	'29		( 26 + 27 + 28 ) �� ����Ͽ� �Է���.
	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "26", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "27", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "28", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt + UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "29", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	Frm1.vspdData.text  = iSumAmt

	'32	( 29 - 30 - 31 ) �� ����Ͽ� �Է���.
	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "30", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt - UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "31", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	iSumAmt = iSumAmt - UNICDbl(Frm1.vspdData.text)

	iSchRow = Frm1.vspdData.SearchCol(C_FORM_REP_NO, 0, Frm1.vspdData.MaxRows, "32", 0)
   	Frm1.vspdData.Row = iSchRow
	Frm1.vspdData.Col = C_AMT
	Frm1.vspdData.text  = iSumAmt

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow iSchRow
	
	SubMakeSum = True
End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP_BAK"><font color=white>��3ȣ��3(3) �μӸ� �������</font></td>
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
									<TD CLASS="TD5">�������</TD>
									<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFISC_YEAR CLASS=FPDTYYYY title=FPDATETIME ALT="�������" tag="14X1" id=txtFISC_YEAR></OBJECT>');</SCRIPT>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�Ű���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>

				<TR>
					<TD WIDTH="100%" valign=top>
						<TABLE  CLASS="TB3" CELLSPACING=0>
							<TR>
								<TD HEIGHT="100%" NOWRAP COLSPAN="4">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtCO_CD_Body" tag="24" tabindex="-1">
<INPUT TYPE=hidden name=txtFISC_YEAR_Body  tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtREP_TYPE_Body" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtBS_PL_FG" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

e