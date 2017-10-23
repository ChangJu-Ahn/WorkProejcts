<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : �λ縶��Ÿ��� 
*  3. Program ID           : H9102ma1
*  4. Program Name         : H9102ma1
*  5. Program Desc         : �����������/��������/�ҵ�.���װ������ 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/04
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : mok young bin
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript"   SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h9102mb1.asp"						           '��: Biz Logic ASP Name
Const TAB1 = 1
Const TAB2 = 2
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '���� TAB�� ��ġ�� ��Ÿ���� Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType

Dim IsOpenPop						                                    ' Popup
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  Parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
	ElseIf flgs = 0 Then

		strTemp =  ReadCookie(CookieSplit)
		If strTemp = "" then 
		    Exit Function
		End if	
		If Err.number <> 0 Then
			Err.Clear
			 WriteCookie CookieSplit , ""
			Exit Function 
		End If

		 WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If

End Function
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
		
	Call  AppendNumberPlace("6", "3", "0")
	Call  AppendNumberPlace("9", "3", "1")

	Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)

	Call  ggoOper.FormatNumber(frm1.txtNational_pension_sub_rate,"100","0",false)
	'Call  ggoOper.FormatNumber(frm1.txtRnd_nontax_limit,"100","0",false)  ' 20040302 by lsn 
	Call  ggoOper.FormatNumber(frm1.txtIncome_tax_rate1,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIncome_tax_rate2,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtParia_med_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtLegal_contr_rate,"100","0",false)

	Call  ggoOper.FormatNumber(frm1.txtTaxLaw_contr_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtApp_contr_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtOurStock_contr_rate,"100","0",false)	'2004 �츮���ֱ�αݰ������� 
	
	Call  ggoOper.FormatNumber(frm1.txtHouse_fund_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIncome_card_rate1,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIncome_card_rate2,"100","0",false)

    Call  ggoOper.FormatNumber(frm1.txtIncome_card2_rate2,"100","0",false)	'����ī��(2003)
    Call  ggoOper.FormatNumber(frm1.txtFore_edu_rate,"100","0",false)	'�ܱ��αٷ����Ǳ����������(2003)    
    Call  ggoOper.FormatNumber(frm1.txtForeign_separate_tax_rate,"100","0",false)	'2004 �ܱ��αٷ��ںи������� 
	
	Call  ggoOper.FormatNumber(frm1.txtInvest_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIndiv_anu_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIndiv_anu2_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtHouse_repay_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtPer_edu_sub,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtMed_sub_bas_amt,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtFarm_tax,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtRes_tax,"100","0",false)

	Call  ggoOper.LockField(Document, "N")											'��: Lock Field
	Call SetToolbar("1110100000000111")												'��: Set ToolBar

	
	Call InitVariables

    Call changeTabs(TAB1)
    gIsTab     = "Y" ' <- "Yes"�� ���� Y(����) �Դϴ�.[V(����)�ƴմϴ�]
    gTabMaxCnt = 2   ' Tab�� ������ ���� �ּ���    

	Call CookiePage (0)                                                             '��: Check Cookie
    Call MainQuery()			
			
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900013",  Parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '��: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If
    
    Call InitVariables                                                           '��: Initializes local global variables

    Call  DisableToolBar( Parent.TBC_QUERY)
	If DBQuery=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
       
    FncQuery = True                                                              '��: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  Parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '��: Clear Condition Field

	Call  ggoOper.FormatNumber(frm1.txtNational_pension_sub_rate,"100","0",false)
	'Call  ggoOper.FormatNumber(frm1.txtRnd_nontax_limit,"100","0",false)  ' 20040302 by lsn 	
	Call  ggoOper.FormatNumber(frm1.txtIncome_tax_rate1,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIncome_tax_rate2,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtParia_med_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtLegal_contr_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtTaxLaw_contr_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtApp_contr_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtOurStock_contr_rate,"100","0",false)	'2004 �츮���ֱ�αݰ������� 
	
	Call  ggoOper.FormatNumber(frm1.txtHouse_fund_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIncome_card_rate1,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIncome_card_rate2,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIncome_card2_rate2,"100","0",false) '����ī��(2003)
    Call  ggoOper.FormatNumber(frm1.txtFore_edu_rate,"100","0",false)	'�ܱ��αٷ����Ǳ����������(2003)	
    Call  ggoOper.FormatNumber(frm1.txtForeign_separate_tax_rate,"100","0",false)	'2004 �ܱ��αٷ��ںи�������    
	
	Call  ggoOper.FormatNumber(frm1.txtInvest_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIndiv_anu_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtIndiv_anu2_rate,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtHouse_repay_rate,"100","0",false)
	
	Call  ggoOper.FormatNumber(frm1.txtPer_edu_sub,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtMed_sub_bas_amt,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtFarm_tax,"100","0",false)
	Call  ggoOper.FormatNumber(frm1.txtRes_tax,"100","0",false)
    
    Call  ggoOper.LockField(Document , "N")                                       '��: Lock  Field
    
    Call SetToolbar("11101000000011")
    Call InitVariables                                                        '��: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '��: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
    If lgIntFlgMode <>  Parent.OPMD_UMODE Then                                           '��: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  Parent.VB_YES_NO,"x","x")                        '��: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call  DisableToolBar( Parent.TBC_DELETE)
	If DBDelete=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If

    Set gActiveElement = document.ActiveElement   
    
    FncDelete = True                                                            '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 

    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '��: Check contents area
       Exit Function
    End If
    
    Call  DisableToolBar( Parent.TBC_SAVE)
	If DBSave=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
    
    FncSave = True                                                              '��: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '��: Protect system from crashing
End Function
'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport( Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( Parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900016",  Parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '��: Clear err status

    DbQuery = False                                                              '��: Processing is NG

    If LayerShowHide(1)=False Then
		Exit Function
    End If


    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0001                       '��: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '��: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic
    
    DbQuery = True                                                               '��: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
    If LayerShowHide(1)=False Then
		Exit Function
    End If

	With Frm1
		.txtMode.value        =  Parent.UID_M0002                                        '��: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '��: Processing is NG
End Function 
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '��: Clear err status
		
	DbDelete = False			                                                 '��: Processing is NG
		
    If LayerShowHide(1)=False Then
		Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode="          & Parent.UID_M0003                       '��: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '��: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '��: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '��: Run Biz logic
	
	DbDelete = True                                                              '��: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim strVal

	lgIntFlgMode      =  Parent.OPMD_UMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = false

	Call SetToolbar("1100100000000111")

    Call  ggoOper.LockField(Document, "Q")
    
	IF frm1.txtSub_fam_flag1.checked = True Then    
		Call ggoOper.SetReqAttr(frm1.txtStd_sub_amt2, "Q") 
	End If		
		   
	IF frm1.txtSub_fam_flag2.checked = True Then
		Call ggoOper.SetReqAttr(frm1.txtStd_sub_amt2, "D") 	
		Call ggoOper.SetReqAttr(frm1.txtSub_fam1, "Q")
		Call ggoOper.SetReqAttr(frm1.txtSub_fam1_amt, "Q")
		Call ggoOper.SetReqAttr(frm1.txtSub_fam2, "Q")
		Call ggoOper.SetReqAttr(frm1.txtSub_fam2_amt, "Q")	
	End If
	
    Set gActiveElement = document.ActiveElement   
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call InitVariables	
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call SetToolbar("1110100000000111")												'��: Set ToolBar
	Call InitVariables()
	Call MainNew()	
End Function

'========================================================================================================
' Name : PgmJump1(PGM_JUMP_ID)
' Desc : developer describe this line 
'========================================================================================================

Function PgmJump1(PGM_JUMP_ID)
    Call CookiePage(1)  ' Write Cookie
    PgmJump(PGM_JUMP_ID)
End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp()
	Dim arrRet
	Dim arrParam(1)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no1.value			' Code Condition
	arrParam(1) = ""'frm1.txtName1.value			' Name Cindition
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no1.value = arrRet(0)
		.txtName1.value = arrRet(1)
		Call  ggoOper.ClearField(Document, "2")					 '��: Clear Contents  Field
		Set gActiveElement = document.ActiveElement

		lgBlnFlgChgValue = False
	End With
End Sub

'==========================================  2.3.1 Tab Click ó��  =================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'===================================================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)
	
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	
	gSelframeFlg = TAB2
End Function


'========================================================================================================
' Name : SubOpenCollateralNoPop()
' Desc : developer describe this line Call Master L/C No PopUp
'========================================================================================================
Sub SubOpenCollateralNoPop()
	Dim strRet
	If gblnWinEvent = True Then Exit Sub
	gblnWinEvent = True
		
	strRet = window.showModalDialog("s1413pa1.asp", "", _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
       Exit Sub
	Else
       Call SetCollateralNo(strRet)
	End If	
End Sub

'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'****************************************************************************************************

Sub txtNon_tax_bas_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtNon_tax_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtNon_dinn_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtOversea_labor_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_bas_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_rate1_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_bas1_amt_Change()
	lgBlnFlgChgValue = True
End Sub
							                          
Sub txtIncome_calcu_bas_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_rate2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_calcu_bas1_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_rate3_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_bas2_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_calcu_bas2_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_rate4_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_bas3_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_rate5_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtNational_pension_sub_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtRnd_nontax_limit_Change()  ' 20040302 by lsn 
	lgBlnFlgChgValue = True
End Sub

Sub txtPer_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSpouse_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtFam_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtOld_sub_amt1_Change()	'2004 ��ο�����(65���̻�)
	lgBlnFlgChgValue = True
End Sub

Sub txtOld_sub_amt2_Change()	'2004 ��ο�����(70���̻�)
	lgBlnFlgChgValue = True
End Sub

Sub txtParia_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtChl_rear_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLady_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSmall_sub1_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSmall_sub2_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_tax_sub_bas_amt_Change()
	lgBlnFlgChgValue = True
	frm1.txtIncome_tax_sub_bas1_amt.value = frm1.txtIncome_tax_sub_bas_amt.value
End Sub

Sub txtIncome_tax_rate1_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_tax_bas_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_tax_sub_bas1_amt_Change()
	lgBlnFlgChgValue = True
	frm1.txtIncome_tax_sub_bas_amt.value = frm1.txtIncome_tax_sub_bas1_amt.value
End Sub

Sub txtIncome_tax_rate2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_tax_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub



Sub txtOther_insur_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtDisabled_insur_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMed_sub_bas_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMed_sub_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtParia_med_rate_Change()
	lgBlnFlgChgValue = True
End Sub
                					        
Sub txtPer_edu_sub_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtFam_edu_sub_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtKind_edu_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtUniv_edu_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub
                					        
Sub txtLegal_contr_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTaxLaw_contr_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtApp_contr_rate_Change()	
	lgBlnFlgChgValue = True
End Sub

Sub txtOurStock_contr_Change()	'2004 �츮���ֱ�αݰ������� 
	lgBlnFlgChgValue = True
End Sub

Sub txtHouse_fund_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtHouse_fund_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLong_house_loan_limit_Change() '��������������Ա��������ѵ���(2003)
	lgBlnFlgChgValue = True
End Sub

Sub txtLong_house_loan_limit1_Change() '2004 ��������������Ա��������ѵ��� (��ȯ�Ⱓ 15���̻�)
	lgBlnFlgChgValue = True
End Sub

Sub txtStd_sub_amt2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_card_rate1_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_card_rate2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_card_rate2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_card2_rate2_Change() '(2003)
	lgBlnFlgChgValue = True
End Sub

Sub txtCeremony_amt_Change()	'2004  ��ȥ/���/�̻������� 
	lgBlnFlgChgValue = True
End Sub

Sub txtForeign_separate_tax_rate_Change() '2004 �ܱ��αٷ��ںи������� 
	lgBlnFlgChgValue = True
End Sub

Sub txtFore_edu_rate_Change() 
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_card_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtInvest_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtInvest_rate2_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIndiv_anu_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIndiv_anu2_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIndiv_anu_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIndiv_anu2_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtHouse_repay_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMedPrint_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLong_Stock_save_rate_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLong_Stock_save_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtFarm_tax_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtRes_tax_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIncome_sub_rate4_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtOur_Stock_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtLong_Stock_save_rate1_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtDisabled_edu_limit_amt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtSub_fam1_Change()
		lgBlnFlgChgValue = True
End Sub
Sub txtSub_fam1_amt_Change()
		lgBlnFlgChgValue = True
End Sub
Sub txtSub_fam2_Change()
		lgBlnFlgChgValue = True
End Sub
Sub txtSub_fam2_amt_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtSub_fam_flag1_OnClick()
	lgBlnFlgChgValue = True
	
	Call ggoOper.SetReqAttr(frm1.txtStd_sub_amt2, "Q")
	Call ggoOper.SetReqAttr(frm1.txtSub_fam1, "D")
	Call ggoOper.SetReqAttr(frm1.txtSub_fam1_amt, "D")
	Call ggoOper.SetReqAttr(frm1.txtSub_fam2, "D")
	Call ggoOper.SetReqAttr(frm1.txtSub_fam2_amt, "D")			
End Sub

Sub txtSub_fam_flag2_OnClick()
	lgBlnFlgChgValue = True
	Call ggoOper.SetReqAttr(frm1.txtStd_sub_amt2, "D") 	
	Call ggoOper.SetReqAttr(frm1.txtSub_fam1, "Q")
	Call ggoOper.SetReqAttr(frm1.txtSub_fam1_amt, "Q")
	Call ggoOper.SetReqAttr(frm1.txtSub_fam2, "Q")
	Call ggoOper.SetReqAttr(frm1.txtSub_fam2_amt, "Q")			
	
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="AUTO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23><% ' ����ġ %>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
							<TR>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif"><img src="../../../Cshared/Image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�޿����û���</font></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="right"><img src="../../../Cshared/Image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../Cshared/Image/table/tab_up_bg.gif"><img src="../../../Cshared/Image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../Cshared/Image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����������û���</font></td>
								<td background="../../../Cshared/Image/table/tab_up_bg.gif" align="right"><img src="../../../Cshared/Image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR><% ' ����ġ ���� %>
	<TR HEIGHT=*>
		<TD CLASS="Tab11">
            <TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
   	            <TR>
   	                <TD WIDTH=100% VALIGN="TOP" HEIGHT="*">
					    <!-- TAB1 ù��° �� ���� -->
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
                        <TABLE width=100%>
							<TR>
							    <TD VALIGN=TOP colspan="2">
							        <FIELDSET CLASS="CLSFLD">
							            <TABLE HEIGHT=100% CELLPADDING="3" CELLSPACING=0 WIDTH=100% >
							                <TR>
							                    <TD CLASS=TD5 NOWRAP>�����������������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax_bas_amt name=txtNon_tax_bas_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�����������������"></OBJECT>');</SCRIPT>&nbsp;����</TD>
							                    <TD CLASS=TD5 NOWRAP>���������ѵ���(����)</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_tax_limit_amt name=txtNon_tax_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="���������ѵ���"></OBJECT>');</SCRIPT></TD>
                					        </TR>
                					        <TR>
							                    <TD CLASS=TD5 NOWRAP>�Ĵ������ѵ���(��)</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNon_dinn_amt name=txtNon_dinn_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�Ĵ������ѵ���"></OBJECT>');</SCRIPT></TD>
							                    <TD CLASS=TD5 NOWRAP>���ܱٷκ�����ѵ���(��)</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtOversea_labor_limit_amt name=txtOversea_labor_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="���ܱٷκ�����ѵ���"></OBJECT>');</SCRIPT></TD>
							                </TR>
                					        <TR>
							                    <TD CLASS=TD5 NOWRAP>����������ѵ�(��)</TD>
												<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtRnd_nontax_limit name=txtRnd_nontax_limit CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="����������ѵ�"></OBJECT>');</SCRIPT></TD>
							                    <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                					            <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
							                </TR>							                
							            </TABLE>
							        </FIELDSET>
							    </TD>    
							</TR>
							<TR>
							    <TD VALIGN=TOP>
									<table  WIDTH=100% CELLSPACING=0 CELLPADDING=0 >
										<tr>
											<td VALIGN=TOP>
												<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>�ٷμҵ����</LEGEND>
												    <table Height=100% WIDTH=100% CELLSPACING=0 CELLPADDING="7"  >
														<TR bgcolor=#D1E8F9>
															<TD WIDTH=40% Height="30" ALIGN=CENTER >�ѱ޿���</TD>
															<TD WIDTH=60% Height="30" ALIGN=Left>�ҵ������</TD>
														</TR>
														<TR bgcolor=#EEEEEC>
															<TD WIDTH=40% Height="20" ALIGN=CENTER ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_bas0_amt name=txtIncome_sub_bas0_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ����1"></OBJECT>');</SCRIPT>����</TD>	
															<TD WIDTH=60% Height="20" ALIGN=Left ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_rate1 name=txtIncome_sub_rate1 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X9Z" ALT="�ٷμҵ����2"></OBJECT>');</SCRIPT>&nbsp;% ����
														</TR>
														<TR bgcolor=#EEEEEC>
															<TD WIDTH=40% Height="20" ALIGN=CENTER ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_bas_amt name=txtIncome_sub_bas_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ����1"></OBJECT>');</SCRIPT>�ʰ�</TD>	
															<TD WIDTH=60% Height="20" ALIGN=Left ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_calcu_bas_amt name=txtIncome_calcu_bas_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ����4"></OBJECT>');</SCRIPT>
															&nbsp;+ �ʰ��ݾ� x &nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_rate2 name=txtIncome_sub_rate2 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X9Z" ALT="�ٷμҵ����5"></OBJECT>');</SCRIPT>%</td>	
														</TR>
														<TR bgcolor=#EEEEEC>
															<TD WIDTH=40% Height="20" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_bas1_amt name=txtIncome_sub_bas1_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ����3"></OBJECT>');</SCRIPT>�ʰ�</TD>
															<TD WIDTH=60% Height="20" ALIGN=Left><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_calcu_bas1_amt name=txtIncome_calcu_bas1_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ����6"></OBJECT>');</SCRIPT>
															&nbsp;+ �ʰ��ݾ� x &nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_rate3 name=txtIncome_sub_rate3 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X9Z" ALT="�ٷμҵ����7"></OBJECT>');</SCRIPT>%</td>
														</TR>
														<TR bgcolor=#EEEEEC>
															<TD WIDTH=40% Height="20" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_bas2_amt name=txtIncome_sub_bas2_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ����5"></OBJECT>');</SCRIPT>�ʰ�</TD>
															<TD WIDTH=60% Height="20" ALIGN=Left><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_calcu_bas2_amt name=txtIncome_calcu_bas2_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ����8"></OBJECT>');</SCRIPT>
															&nbsp;+ �ʰ��ݾ� x &nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_rate4 name=txtIncome_sub_rate4 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X9Z" ALT="�ٷμҵ����9"></OBJECT>');</SCRIPT>%</TD>
														</TR>
														<TR bgcolor=#EEEEEC>
															<TD WIDTH=40% Height="20" ALIGN=CENTER><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_bas3_amt name=txtIncome_sub_bas3_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ѵ���"></OBJECT>');</SCRIPT>�ʰ�</TD>
															<TD WIDTH=60% Height="20" ALIGN=Left><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_limit_amt name=txtIncome_sub_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ����9"></OBJECT>');</SCRIPT>
															&nbsp;+ �ʰ��ݾ� x &nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_sub_rate5 name=txtIncome_sub_rate5 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X9Z" ALT="�ٷμҵ����9"></OBJECT>');</SCRIPT>%</TD>
														</TR>
														<TR bgcolor=#EEEEEC>
															<TD WIDTH=40% Height="20" ALIGN=CENTER>&nbsp;</TD>
															<TD WIDTH=60% Height="20" ALIGN=Left>&nbsp;</TD>
														</TR>														
													</TABLE>		
												</FIELDSET>
											</td>
										</tr>
									
									</TABLE>
							    </TD> 
							    <TD  VALIGN=TOP>
									<table WIDTH=100% CELLSPACING=0 CELLPADDING=0>
									
									<tr>
										<td VALIGN=TOP>
											<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>��������</LEGEND>
									        <TABLE HEIGHT=100% WIDTH=100% CELLSPACING=0 CELLPADDING="3">
                						        <TR>
									                <TD CLASS=TD5 Height="20" NOWRAP>�� ��</TD>
                						            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtPer_sub_amt name=txtPer_sub_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�⺻"></OBJECT>');</SCRIPT></TD>
									                <TD CLASS=TD5 Height="20" NOWRAP>�����</TD>
                						            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSpouse_sub_amt name=txtSpouse_sub_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�����"></OBJECT>');</SCRIPT></TD>
                						        </TR>
                						        <TR>
									                <TD CLASS=TD5 Height="20" NOWRAP>�� ��</TD>
                						            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtFam_sub_amt name=txtFam_sub_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ξ�"></OBJECT>');</SCRIPT></TD>
									                <TD CLASS=TD5 Height="20" NOWRAP>&nbsp;</TD>
                						            <TD CLASS=TD6 Height="20" NOWRAP></TD>
                						        </TR>
									        </TABLE>
											</FIELDSET>
										</td>
									</tr>
									
									<tr>
										<TD VALIGN=TOP>
										    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>�߰�����</LEGEND>
										        <TABLE HEIGHT=100% WIDTH=100% CELLSPACING=0 CELLPADDING="5">
                							        <TR>
										                <TD CLASS=TD5 Height="20" NOWRAP>��ο��(65���̻�)</TD>
                							            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtOld_sub_amt1 name=txtOld_sub_amt1 CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="��ο��"></OBJECT>');</SCRIPT></TD>
										                <TD CLASS=TD5 Height="20" NOWRAP>��ο��(70���̻�)</TD>
                							            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtOld_sub_amt2 name=txtOld_sub_amt2 CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="��ο��"></OBJECT>');</SCRIPT></TD>
                							        </TR>
                							        <TR>
										                <TD CLASS=TD5 Height="20" NOWRAP>�ڳ����</TD>
                							            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtChl_rear_sub_amt name=txtChl_rear_sub_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ڳ����"></OBJECT>');</SCRIPT></TD>
										                <TD CLASS=TD5 Height="20" NOWRAP>�γ���</TD>
                							            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLady_sub_amt name=txtLady_sub_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�γ���"></OBJECT>');</SCRIPT></TD>
                							        </TR>
                							        <TR>
										                <TD CLASS=TD5 Height="20" NOWRAP>�����</TD>
                							            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtParia_sub_amt name=txtParia_sub_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�����"></OBJECT>');</SCRIPT></TD>
										                <TD CLASS=TD5 Height="20" NOWRAP>&nbsp;</TD>
                							            <TD CLASS=TD6 Height="20" NOWRAP>&nbsp;</TD>                							            
                							        </TR>
                							        <TR>
										                <TD CLASS=TD5 Height="20" NOWRAP>���ڳ��߰�2��</TD>
                							            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSmall_sub1_amt name=txtSmall_sub1_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�Ҽ��߰�1��"></OBJECT>');</SCRIPT></TD>
										                <TD CLASS=TD5 Height="20" NOWRAP>���ڳ��߰�3���̻�</TD>
                							            <TD CLASS=TD6 Height="20" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSmall_sub2_amt name=txtSmall_sub2_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�Ҽ��߰�2��"></OBJECT>');</SCRIPT></TD>
                							        </TR>                							        
										        </TABLE>
										    </FIELDSET>
										</TD>    
									</tr>
									<TR>
										<TD colspan="2" VALIGN=TOP>
											<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>���ݺ�������</LEGEND>
												<TABLE HEIGHT=100% WIDTH=100% CELLSPACING=0 CELLPADDING="5">
                								    <TR>
												        <TD CLASS=TD5 NOWRAP>���ο��ݳ��ξ���</TD>
                								        <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtNational_pension_sub_rate name=txtNational_pension_sub_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="���ݺ�������"></OBJECT>');</SCRIPT>%</TD>
												        <TD CLASS=TDT NOWRAP>&nbsp;</TD>
                								        <TD CLASS=TD6 NOWRAP></TD>
                								    </TR>
												</TABLE>
											</FIELDSET>
										</TD>
									</TR>
									</table>
							    </TD>   
							</TR>
							<TR>
							    <TD  VALIGN=TOP colspan="2">
							        <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>�ٷμҵ漼�װ���</LEGEND>
							            <TABLE HEIGHT=100%  WIDTH=100% CELLSPACING=0 CELLPADDING="3">
							                <TR>
							                    <TD CLASS=TD5 NOWRAP>���⼼��</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_tax_sub_bas_amt name=txtIncome_tax_sub_bas_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ漼�װ���1"></OBJECT>');</SCRIPT></TD>
							                    <TD CLASS=TD5 NOWRAP>������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_tax_rate1 name=txtIncome_tax_rate1 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ٷμҵ漼�װ���2"></OBJECT>');</SCRIPT>%</TD>
                					        </TR>
                					        <TR>
							                    <TD CLASS=TD5 NOWRAP>�ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_tax_limit_amt name=txtIncome_tax_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ѵ���"></OBJECT>');</SCRIPT></TD>
							                    <TD CLASS=TD5 NOWRAP>�ʰ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_tax_bas_amt name=txtIncome_tax_bas_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ漼�װ���3"></OBJECT>');</SCRIPT>+(
                					                                 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_tax_sub_bas1_amt name=txtIncome_tax_sub_bas1_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ٷμҵ漼�װ���4"></OBJECT>');</SCRIPT>�ʰ��ݾ�
                					                                 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_tax_rate2 name=txtIncome_tax_rate2 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ٷμҵ漼�װ���5"></OBJECT>');</SCRIPT>%)
                					            </TD>
                					        </TR>
							            </TABLE>
							        </FIELDSET>
							    </TD>    
							</TR>
			    		    <TR>
				        	    <TD VALIGN="TOP" COLSPAN=2>
									<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>Ư������</LEGEND>
									    <table Height=100% WIDTH=100% CELLSPACING=0>
											<TR bgcolor=#D1E8F9>
              							        <TD CLASS="TD5" NOWRAP>���뿩��</TD>
	                   						    <TD CLASS="TD6">
                                                    <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtSub_fam_flag1" NAME="txtSub_fam_flag" TAG="21X" VALUE="Y"><LABEL FOR="txtSub_fam_flag1">YES:(���Ρ�����ڡ��ξ簡��)</LABEL></TD>

              							        <TD CLASS="TD5" NOWRAP>�����������</TD>
	                   						    <TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSub_fam1 name=txtSub_fam1 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X2Z" ALT="�����������"></OBJECT>');</SCRIPT>��&nbsp;����&nbsp;&nbsp;
												                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSub_fam1_amt name=txtSub_fam1_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="21X2Z" ALT="��������ݾ�"></OBJECT>');</SCRIPT></TD>
								 			</TR>
											<TR>
              									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
	                   							<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" ID="txtSub_fam_flag2" NAME="txtSub_fam_flag" TAG="21X" VALUE="N" CHECKED><LABEL FOR="txtSub_fam_flag2">NO&nbsp;:(ǥ�ذ���)</LABEL>
      							                				<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtStd_sub_amt2 name=txtStd_sub_amt2 CLASS=FPDS90 title=FPDOUBLESINGLE tag="21X2" ALT="ǥ�ذ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
              							        <TD CLASS="TD5" NOWRAP></TD>
	                   						    <TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSub_fam2 name=txtSub_fam2 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X2Z" ALT="�����������"></OBJECT>');</SCRIPT>��&nbsp;�̻�&nbsp;&nbsp;
												                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSub_fam2_amt name=txtSub_fam2_amt CLASS=FPDS140 title=FPDOUBLESINGLE tag="21X2Z" ALT="��������ݾ�"></OBJECT>');</SCRIPT></TD>
											</TR>											
										</TABLE>		
									</FIELDSET>
				        	    </TD>			    		    
			    		    </TR>
					    </TABLE>
					    </DIV><!-- ù��° �� ���� -->
    
					    <!-- TAB2 �� ���� -->
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
                        <TABLE width=100%>
							<TR>
							    <TD VALIGN=TOP>
							        <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>Ư������</LEGEND>
										<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                					        <TR>
							                    <TD CLASS=TD5 NOWRAP>��Ÿ��������ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtOther_insur_limit_amt name=txtOther_insur_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="��Ÿ��������ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
							                    <TD CLASS=TD5 NOWRAP>��������� ���庸�� �ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtDisabled_insur_limit_amt name=txtDisabled_insur_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="��������뺸������ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
                					        </TR>
                					        <TR>
                					            <TD CLASS=TD5 NOWRAP>�Ƿ���������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMed_sub_bas_amt name=txtMed_sub_bas_amt CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�Ƿ���������"></OBJECT>');</SCRIPT>%&nbsp;</TD>
							                    <TD CLASS=TD5 NOWRAP>�Ƿ������ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMed_sub_limit_amt name=txtMed_sub_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�Ƿ������ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
											</TR>
                					        <TR>							                    
							                    <TD CLASS=TD5 NOWRAP>����/�����/������Ƿ�������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtParia_med_rate name=txtParia_med_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="��ι�����Ƿ�������"></OBJECT>');</SCRIPT>%&nbsp;</TD>
                					            <TD CLASS=TD5 NOWRAP>�����Ư��������1���ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtDisabled_edu_limit_amt" name=txtDisabled_edu_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�����Ư��������1���ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
											</TR>
                					        <TR>
							                    <TD CLASS=TD5 NOWRAP>���α����������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtPer_edu_sub name=txtPer_edu_sub CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="���α����������"></OBJECT>');</SCRIPT>%&nbsp;</TD>
                					            <TD CLASS=TD5 NOWRAP>��ġ��������1���ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtKind_edu_limit_amt name=txtKind_edu_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="��ġ��������1���ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
							                </TR>
                					        <TR>
                								<TD CLASS=TD5 NOWRAP>���߰�����1���ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtFam_edu_sub_amt name=txtFam_edu_sub_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="���߰�����1���ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
							                    <TD CLASS=TD5 NOWRAP>���б�����1���ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtUniv_edu_limit_amt name=txtUniv_edu_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="���б�����1���ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
							                </TR>
                					        <TR>
							                    <TD CLASS=TD5 NOWRAP>�������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLegal_contr_rate name=txtLegal_contr_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�������"></OBJECT>');</SCRIPT>%&nbsp;&nbsp;&nbsp;&nbsp;
                					            ��Ư�� ��73�� ��α�<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="Object1" name=txtTaxLaw_contr_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="��Ư�� ��73�� ��α�"></OBJECT>');</SCRIPT>%&nbsp;</TD>
							                    <TD CLASS=TD5 NOWRAP>�������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtApp_contr_rate name=txtApp_contr_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�������"></OBJECT>');</SCRIPT>%&nbsp;</TD>
							                </TR>
                					        <TR>
							                    <TD CLASS=TD5 NOWRAP>�츮���ֱ�αݰ�������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtOurStock_contr_rate name=txtOurStock_contr_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�츮���ֱ�αݰ�������"></OBJECT>');</SCRIPT>%</TD>
							                    <TD CLASS=TD5 NOWRAP>�����ڱ��ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtHouse_fund_limit_amt name=txtHouse_fund_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�����ڱ��ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
							                </TR>							                
                					        <TR>
                					            <TD CLASS=TD5 NOWRAP>�����ڱݰ�����</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> valign=middle id=txtHouse_fund_rate name=txtHouse_fund_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�����ڱݰ�����"></OBJECT>');</SCRIPT>%&nbsp;</TD>
                					            <TD CLASS=TD5 NOWRAP>ī�����ѵ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_card_limit_amt name=txtIncome_card_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="ī�����ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
							                </TR>
                					        <TR>
							                    <TD CLASS=TD5 NOWRAP>ī����׿���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIncome_card_rate1 name=txtIncome_card_rate1 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="ī����׿���"></OBJECT>');</SCRIPT>%
                					            �ʰ�������<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle43 name=txtIncome_card_rate2 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ʰ�������"></OBJECT>');</SCRIPT>%&nbsp;(�ſ�ī��)</TD>
                								<TD CLASS=TD5 NOWRAP>��ȥ/���/�̻�������</TD>
                								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtCeremony_amt name=txtCeremony_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="��ȥ/���/�̻�������"></OBJECT>');</SCRIPT>
							                </TR>
							                <TR>
												<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                								<TD CLASS=TD6 NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�ʰ�������<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle43 name=txtIncome_card2_rate2 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ʰ�������"></OBJECT>');</SCRIPT>%&nbsp;(����ī��)</TD>
                								<TD CLASS=TD5 NOWRAP>�ܱ��αٷ��ڱ������������</TD>
                								<TD CLASS=TD6 NOWRAP>������<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtFore_edu_rate name=txtFore_edu_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ܱ��αٷ��ڱ�����,������"></OBJECT>');</SCRIPT>%</TD>
							                </TR>                										    
							                <TR>
                					            <TD CLASS=TD5 NOWRAP>�ܱ��αٷ��ںи�����������</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtForeign_separate_tax_rate name=txtForeign_separate_tax_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ܱ��αٷ��ڱ�����,������"></OBJECT>');</SCRIPT>%</TD>
												<TD CLASS=TD5 NOWRAP>��������������Ա����ڻ�ȯ�ѵ���</TD>
                								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLong_house_loan_limit name=txtLong_house_loan_limit CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="��������������Ա����ڻ�ȯ�ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
							                </TR>
							                <TR>
                					            <TD CLASS=TD5 NOWRAP>ǥ�ذ���</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtStd_sub_amt name=txtStd_sub_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="24X2Z" ALT="ǥ�ذ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
												<TD CLASS=TD5 NOWRAP>��������������Ա����ڻ�ȯ�ѵ���(��ȯ�Ⱓ 15���̻�)</TD>
                								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLong_house_loan_limit1 name=txtLong_house_loan_limit1 CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="��������������Ա����ڻ�ȯ�ѵ���"></OBJECT>');</SCRIPT>&nbsp;</TD>
							                </TR>
 							            </TABLE>
							        </FIELDSET>
							    </TD>
							</TR>
							<TR>
								<td valign=top>
									<table width=100%>
										<tr>
											<TD WIDTH=50%>
												<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>�����������ڼҵ����&nbsp;�ѵ���</LEGEND>
													<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                										<TR>
															<TD CLASS=TD5 NOWRAP>�ٷμҵ�ݾ���</TD>
                											<TD CLASS=TD6 NOWRAP>2001.12.31����<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtInvest_rate name=txtInvest_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ٷμҵ�ݾ���"></OBJECT>');</SCRIPT>%&nbsp;&nbsp;
                											                                   ����<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtInvest_rate2" name=txtInvest_rate2 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ѵ���"></OBJECT>');</SCRIPT>%</TD>
                										</TR>
													</TABLE>
												</FIELDSET>
											</TD>
											<TD WIDTH=50%>
												<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>�츮�����⿬�ݾ׼ҵ����</LEGEND>
													<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0 ID="Table1">
                										<TR>
															<TD CLASS=TD5 NOWRAP>�ѵ���</TD>
                											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtOur_Stock_limit_amt" name=txtOur_Stock_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�츮�������� �⿬�� �ҵ���� �ѵ�"></OBJECT>');</SCRIPT>&nbsp;</TD>
                										</TR>
													</TABLE>
												</FIELDSET>
											</TD>
										</tr>
									</table>
								</td>
							</TR>
							<TR>
							    <td VALIGN=TOP>
									<Table WIDTH=100%>
										<tr >
											<TD  WIDTH=50%>
												<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>���ο�������(2001������)</LEGEND>
													<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                										<TR>    
															<TD CLASS=TD5 NOWRAP>������</TD>
                											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIndiv_anu_rate name=txtIndiv_anu_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="������"></OBJECT>');</SCRIPT>%</TD>
															<TD CLASS=TD5 NOWRAP>�ѵ���</TD>
                											<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIndiv_anu_limit_amt name=txtIndiv_anu_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ѵ���"></OBJECT>');</SCRIPT></TD>
                										</TR>
													</TABLE>
												</FIELDSET>
											</TD>    
								
											<TD  WIDTH=50%>
												<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>���ο�������(2001������)</LEGEND>
													<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                										<TR>    
															<TD CLASS=TD5 NOWRAP>������</TD>
			            									<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIndiv_anu2_rate name=txtIndiv_anu2_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="������"></OBJECT>');</SCRIPT>%</TD>
															<TD CLASS=TD5 NOWRAP>�ѵ���</TD>
															<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtIndiv_anu2_limit_amt name=txtIndiv_anu2_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ѵ���"></OBJECT>');</SCRIPT></TD>
                										</TR>
													</TABLE>
												</FIELDSET>
											</TD>  
									    </tr>
								    </Table>  
							    </TD>
							</TR>
							
							<TR>
							    <TD  VALIGN=TOP>
							        <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>�����ڱ����ڼ��װ���</LEGEND>
							            <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                					        <TR>    
							                    <TD CLASS=TD5 NOWRAP>���ڻ�ȯ����</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtHouse_repay_rate name=txtHouse_repay_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="���ڻ�ȯ����"></OBJECT>');</SCRIPT>%</TD>
							                    <TD CLASS=TDT NOWRAP>&nbsp;</TD>
                					            <TD CLASS=TD6 NOWRAP></TD>
                					        </TR>
							            </TABLE>
							        </FIELDSET>
							    </TD>    
							</TR>
							<TR>
							    <TD  VALIGN=TOP>
							        <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>�Ƿ�����������</LEGEND>
							            <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                					        <TR>    
							                    <TD CLASS=TD5 NOWRAP>�����Ƿ��</TD>
                					            <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtMedPrint name=txtMedPrint CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�Ƿ�����������"></OBJECT>');</SCRIPT>�̻�</TD>
							                    <TD CLASS=TDT NOWRAP>&nbsp;</TD>
                					            <TD CLASS=TD6 NOWRAP></TD>
                					        </TR>
							            </TABLE>
							        </FIELDSET>
							    </TD>    
							</TR>							
							<TR>
							    <TD  VALIGN=TOP>
							        <table WIDTH=100%>
							        <TR>
										<TD WIDTH=50%>
											<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>�ֽ����༼�װ���</LEGEND>
												<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=1>
													<TR>
                							            <TD CLASS=TD5 NOWRAP>����ֽ�����</TD>
                							            <TD CLASS=TD6 NOWRAP>&nbsp;���⵵&nbsp;���Ծ���&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLong_Stock_save_rate name=txtLong_Stock_save_rate CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="���Ծ���"></OBJECT>');</SCRIPT>%&nbsp;</TD>
                							        </TR>
													<TR>
                							            <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                							            <TD CLASS=TD6 NOWRAP>�����⵵&nbsp;���Ծ���&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id="txtLong_Stock_save_rate1" name=txtLong_Stock_save_rate1 CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="���Ծ���"></OBJECT>');</SCRIPT>%&nbsp;</TD>
                							        </TR>
                							        <TR>
                							            <TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                							            <TD CLASS=TD6 NOWRAP>1�� 1����&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtLong_Stock_save_limit_amt name=txtLong_Stock_save_limit_amt CLASS=FPDS115 title=FPDOUBLESINGLE tag="21X2Z" ALT="�ѵ�"></OBJECT>');</SCRIPT>�ѵ�&nbsp;</TD>
                							        </TR>
											    </TABLE>
											</FIELDSET>
										</TD>
										<TD WIDTH=50%>
											<FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>��Ÿ ���� �׸�</LEGEND>
												<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
													<TR>
														<TD CLASS=TD5 NOWRAP>��Ư��</TD>
                									    <TD CLASS=TD6 NOWRAP>�����ڱ����ڼ���&nbsp;��������&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtFarm_tax name=txtFarm_tax CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="��Ư��"></OBJECT>');</SCRIPT>%</TD>
													</TR>
													<TR>
                									    <TD CLASS=TD5 NOWRAP>�ֹμ�</TD>
                									    <TD CLASS=TD6 NOWRAP>�ҵ漼��&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtRes_tax name=txtRes_tax CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X6Z" ALT="�ֹμ�"></OBJECT>');</SCRIPT>%</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                									    <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
													</TR>
													<TR>
														<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
                									    <TD CLASS=TD6 NOWRAP>&nbsp;</TD>
													</TR>
												</TABLE>
											</FIELDSET>
										</TD>
									</TR>
									</table>
							    </TD>    
							</TR>
					    </TABLE>
					    </DIV><!-- 2 �� ���� -->
                    </TD>
                </TR>
            </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
</FORM>

</BODY>
</HTML>

