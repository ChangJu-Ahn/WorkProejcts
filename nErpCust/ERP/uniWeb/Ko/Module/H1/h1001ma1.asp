<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : H1001ma1
*  4. Program Name         : H1001ma1
*  5. Program Desc         : 기준정보관리/회사Rule등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/03
*  8. Modified date(Last)  : 2003/05/15
*  9. Modifier (First)     : YBI
* 10. Modifier (Last)      : LSN
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "H1001mb1.asp"
Const WARRANT_TYPE_MAJOR = "S0002"
Const DEL_TYPE_MAJOR     = "S0003"
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029B("I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strYear
	Dim strMonth
	Dim strInsurDt
	Dim stReturnrInsurDt

	lgKeyStream = "1" & parent.gColSep       'You Must append one character( parent.gColSep)

End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0081", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call  SetCombo2(frm1.cboFamily_type, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0082", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call  SetCombo2(frm1.cboIntern_type, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0083", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call  SetCombo2(frm1.cboSave_script_type, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0084", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call  SetCombo2(frm1.cboBas_strt_mm, iCodeArr, iNameArr, Chr(11))

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0084", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr =  lgF0
    iNameArr =  lgF1
    Call  SetCombo2(frm1.cboBas_end_mm, iCodeArr, iNameArr, Chr(11))

End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
    Err.Clear
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call  AppendNumberPlace("7", "3", "3")
	Call  AppendNumberPlace("8", "2", "0")
	Call  AppendNumberPlace("9", "3", "2")
	Call  AppendNumberRange("0", "-12x34", "13x440")
	Call  AppendNumberRange("1", "1", "31")
	Call  AppendNumberRange("2", "100x00", "x99")
	
	Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")

    Call  ggoOper.FormatNumber(frm1.txtBas_strt_dd, 31, 0)
    Call  ggoOper.FormatNumber(frm1.txtBas_end_dd, 31, 0)

    Call  ggoOper.FormatNumber(frm1.txtPay_prov_dd, 31, 0)
    Call  ggoOper.FormatNumber(frm1.txtPay_bas_dd, 31, 0)
    Call  ggoOper.FormatNumber(frm1.txtDilig_dd, 31, 0)
	
    Call SetDefaultVal()
	Call SetToolbar("1100100000001111")
	
	Call InitVariables
    Call InitComboBox
    call MainQuery()			
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
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("Q")
    
	Call  DisableToolBar( parent.TBC_QUERY)
    If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
              
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
    Call SetToolbar("11101000000011")
    Call SetDefaultVal
    Call InitVariables                                                        '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False
    Err.Clear
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call MakeKeyStream("D")
    
	Call  DisableToolBar( parent.TBC_DELETE)
    If DbDelete = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncDelete = True
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
	Dim strBasStrtMm
	Dim strBasStrtDd
	Dim strBasEndMm
	Dim strBasEndDd

    FncSave = False    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

    if  frm1.txtBas_strt_dd.text = "" then
        frm1.txtBas_strt_dd.text = "0"
    end if
   
    if  frm1.txtBas_end_dd.text = "" then
        frm1.txtBas_end_dd.text = "0"
    end if

    if  frm1.txtpay_prov_dd.text = "" then
        frm1.txtpay_prov_dd.text = "0"
    end if

    if  frm1.txtpay_bas_dd.text = "" then
        frm1.txtpay_bas_dd.text = "0"
    end if

    if  frm1.txtdilig_dd.text = "" then
        frm1.txtdilig_dd.text = "0"
    end if
    
	strBasStrtMm = frm1.cboBas_strt_mm.value
	strBasStrtDd = frm1.txtBas_strt_dd.text
	strBasEndMm = frm1.cboBas_end_mm.value
	strBasEndDd = frm1.txtBas_end_dd.text
	
    If strBasStrtDd = 0 Then
        strBasStrtDd = 31
    End If       
    
    If strBasEndDd = 0 Then
        strBasEndDd = 31
    End If       

    If strBasStrtMm > strBasEndMm Then
	    Call  DisplayMsgBox("800445","X","X","X")	               '일단위를 확인하십시요.
	    frm1.cboBas_end_mm.focus
        Set gActiveElement = document.activeElement
        Exit Function
    ElseIf strBasStrtMm = strBasEndMm Then
        If  UNICDbl(strBasStrtDd) >  UNICDbl(strBasEndDd) Then
	        Call  DisplayMsgBox("113118","X","X","X")	               '시작일이 종료일보다 빨라야 합니다.
	        frm1.txtBas_end_dd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End if 
    End If

    if   UNICDbl(frm1.txtanut_comp_rate1.text) > 100 then
        call  DisplayMsgBox("970027", "x","회사부담율","x")
        frm1.txtanut_comp_rate1.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtanut_comp_rate2.text) > 100 then
        call  DisplayMsgBox("970027", "x","회사부담율","x")
        frm1.txtanut_comp_rate2.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtanut_prsn_rate1.text) > 100 then
        call  DisplayMsgBox("970027", "x","본인부담율","x")
        frm1.txtanut_prsn_rate1.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtanut_prsn_rate2.text) > 100 then
        call  DisplayMsgBox("970027", "x","본인부담율","x")
        frm1.txtanut_prsn_rate2.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtanut_retire_rate1.text) > 100 then
        call  DisplayMsgBox("970027", "x","퇴직전환율","x")
        frm1.txtanut_retire_rate1.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtanut_retire_rate2.text) > 100 then
        call  DisplayMsgBox("970027", "x","퇴직전환율","x")
        frm1.txtanut_retire_rate2.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtmed_comp_rate.text) > 100 then
        call  DisplayMsgBox("970027", "x","회사부담율","x")
        frm1.txtmed_comp_rate.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtmed_prsn_rate.text) > 100 then
        call  DisplayMsgBox("970027", "x","본인부담율","x")
        frm1.txtmed_prsn_rate.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtemploy_rate.text) > 100 then
        call  DisplayMsgBox("970027", "x","고용보험율","x")
        frm1.txtemploy_rate.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtre_tax_sub1.text) > 100 then
        call  DisplayMsgBox("970027", "x","퇴직세액공제","x")
        frm1.txtre_tax_sub1.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtre_incom_sub.text) > 100 then
        call  DisplayMsgBox("970027", "x","기타퇴직","x")
        frm1.txtre_incom_sub.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    if   UNICDbl(frm1.txtre_speci_sub.text) > 100 then
        call  DisplayMsgBox("970027", "x","명예퇴직","x")
        frm1.txtre_speci_sub.focus
        Set gActiveElement = document.ActiveElement
        exit function
    end if

    Call MakeKeyStream("S")
	Call  DisableToolBar( parent.TBC_SAVE)
    If DbSave = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
            
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")				     '☜: Data is changed.  Do you want to continue? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode =  parent.OPMD_CMODE												     '⊙: Indicates that current mode is Crate mode
    
    Call  ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call  ggoOper.LockField(Document, "N")									     '⊙: This function lock the suitable field
    Call SetToolbar("11101000000011")
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                            '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
	With Frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

    lgBlnFlgChgValue = false

	Call SetToolbar("1100100000001111")												'⊙: Set ToolBar

    Call  ggoOper.LockField(Document, "Q")
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
	Call InitVariables()
End Function

'========================================================================================================
' Name : txtmed_entr_flag1_Change
' Desc : developer describe this line 
'========================================================================================================
Sub txtmed_entr_flag1_OnClick()
	lgBlnFlgChgValue = True
End Sub
Sub txtmed_entr_flag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtmed_retire_flag1_OnClick()
	lgBlnFlgChgValue = True
End Sub
Sub txtmed_retire_flag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtmed_en_re_flag1_OnClick()
	lgBlnFlgChgValue = True
End Sub
Sub txtmed_en_re_flag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtanut_entr_flag1_OnClick()
	lgBlnFlgChgValue = True
End Sub
Sub txtanut_entr_flag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtanut_retire_flag1_OnClick()
	lgBlnFlgChgValue = True
End Sub
Sub txtanut_retire_flag2_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtanut_en_re_flag1_OnClick()
	lgBlnFlgChgValue = True
End Sub
Sub txtanut_en_re_flag2_OnClick()
	lgBlnFlgChgValue = True
End Sub


Sub txtMed_type1_OnClick()
	lgBlnFlgChgValue = True
End Sub
Sub txtMed_type2_OnClick()
	lgBlnFlgChgValue = True
End Sub


'========================================================================================================
' Name : cboFamily_type_OnChange
' Desc : developer describe this line 
'========================================================================================================
Sub cboFamily_type_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub cboSave_script_type_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub cboIntern_type_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub cboBas_strt_mm_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub cboBas_end_mm_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub txtmed_entr_flag_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub txtmed_retire_flag_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub txtmed_en_re_flag_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub txtanut_entr_flag_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub txtanut_retire_flag_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub txtanut_en_re_flag_OnChange()
		lgBlnFlgChgValue = True
End Sub

Sub txtBas_strt_dd_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtBas_end_dd_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtMed_comp_rate_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtMed_prsn_rate_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtAnut_comp_rate1_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtAnut_comp_rate2_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtAnut_prsn_rate1_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtAnut_prsn_rate2_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtEmploy_rate_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtpay_prov_dd_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtRe_tax_sub1_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtpay_bas_dd_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtRe_incom_sub_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtdilig_dd_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtRe_speci_sub_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtre_sub_limit_Change()
		lgBlnFlgChgValue = True
End Sub

Sub txtAnut_retire_rate1_Change()
		lgBlnFlgChgValue = True
End Sub
Sub txtAnut_retire_rate2_Change()
		lgBlnFlgChgValue = True
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
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
								<TD BACKGROUND"../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" WIDTH="10" HEIGHT="23"></td>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" CLASS="CLSMTAB" ALIGN="center"><FONT COLOR=white>회사RULE등록</font></td>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN="TOP">
						<TABLE <%=LR_SPACE_TYPE_60%>>
				            <TR>
				                <TD VALIGN="TOP" colspan=2>
				            		<FIELDSET CLASS="CLSFLD">
				            		<TABLE CLASS="BasicTB" CELLSPACING=0>
				            			<TR>
				            				<TD CLASS="TD5" NOWRAP>가족수당지급기준</TD>
				            				<TD CLASS="TD6"><SELECT NAME="cboFamily_type" ALT="가족수당지급기준" CLASS ="cbonormal" TAG="21"></SELECT></TD>
				            				<TD CLASS="TD5" NOWRAP>저축불입처리기준</TD>
				            				<TD CLASS="TD6"><SELECT NAME="cboSave_script_type" ALT="저축불입처리기준" CLASS ="cbonormal" TAG="21"></SELECT></TD>
				            			</TR>
				            			<TR>
				            				<TD CLASS="TD5" NOWRAP>수습사원처리기준</TD>
				            				<TD CLASS="TD6">
				            					<SELECT NAME="cboIntern_type" ALT="수습사원처리기준" CLASS ="cbonormal" TAG="21"></SELECT>
				            				</TD>
				            				<TD CLASS="TD5" NOWRAP>보험계산기준기간</TD>
				            				<TD CLASS="TD6">
				            				    <SELECT NAME="cboBas_strt_mm" ALT="보험계산기준기간" CLASS ="cbosmall" TAG="21"></SELECT>
	                   							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtBas_strt_dd NAME=txtBas_strt_dd ALT="보험계산기준기간" CLASS=FPDS40 TITLE=FPDOUBLESINGLE TAG="21X81"></OBJECT>');</SCRIPT>일 ~ 
				            					<SELECT NAME="cboBas_end_mm" ALT="보험계산기준기간" STYLE="WIDTH: 50px" TAG="21"></SELECT>
	                   							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtBas_end_dd NAME=txtBas_end_dd ALT="보험계산기준기간" CLASS=FPDS40 TITLE=FPDOUBLESINGLE TAG="21X81"></OBJECT>');</SCRIPT>일
				            				</TD>
				            			</TR>
				            		</TABLE>
				            		</FIELDSET>
				            	</TD>
				            </TR>
                            <TR>
						        <TD VALIGN="TOP">
						            <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>건강보험처리</LEGEND>
						            <TABLE CLASS="BasicTB" CELLSPACING=0>
        					        	<TR>
              							    <TD CLASS="TD5" NOWRAP>회사부담율</TD>
	                   						<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtMed_comp_rate NAME=txtMed_comp_rate ALT="회사부담율" CLASS=FPDS65 TITLE=FPDOUBLESINGLE TAG="21X7Z"></OBJECT>');</SCRIPT>%</TD>
	                   					</TR>
			        		        	<TR>
					    	           		<TD CLASS="TD5" NOWRAP>본인부담율</TD>
						            		<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtMed_prsn_rate NAME=txtMed_prsn_rate ALT="본인부담율" CLASS=FPDS65 TITLE=FPDOUBLESINGLE TAG="21X7Z"></OBJECT>');</SCRIPT>%</TD>
						    	        </TR>
        				        		<TR>
		        		        			<TD CLASS="TD5" NOWRAP>중도입사자처리</TD>
				                			<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" ID="txtmed_entr_flag1" NAME="txtmed_entr_flag" TAG="21X" VALUE="Y" CHECKED><LABEL FOR="txtmed_entr_flag1">공제함</LABEL>&nbsp;&nbsp;&nbsp;
				                			                <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtmed_entr_flag2" NAME="txtmed_entr_flag" TAG="21X" VALUE="N"><LABEL FOR="txtmed_entr_flag2">공제안함</LABEL></TD>
                                        </TR>
        				        		<TR>
		        		        			<TD CLASS="TD5" NOWRAP>중도퇴사자처리</TD>
				                			<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtmed_retire_flag" ID="txtmed_retire_flag1" TAG="21X" VALUE="Y" CHECKED><LABEL FOR="txtmed_retire_flag1">공제함</LABEL>&nbsp;&nbsp;&nbsp;
				                			                <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtmed_retire_flag" ID="txtmed_retire_flag2" TAG="21X" VALUE="N"><LABEL FOR="txtmed_retire_flag2">공제안함</LABEL></TD>
                                        </TR>
        				        		<TR>
		        		        			<TD CLASS="TD5" NOWRAP>중도입퇴사자처리</TD>
				                			<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtmed_en_re_flag" ID="txtmed_en_re_flag1" TAG="21X" VALUE="Y" CHECKED><LABEL FOR="txtmed_en_re_flag1">공제함</LABEL>&nbsp;&nbsp;&nbsp;
				                			                <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtmed_en_re_flag" ID="txtmed_en_re_flag2" TAG="21X" VALUE="N"><LABEL FOR="txtmed_en_re_flag2">공제안함</LABEL></TD>
                                        </TR>
                                        
                                        <TR>
		        		        			<TD CLASS="TD5" NOWRAP>건강보험처리기준</TD>
				                			<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtMed_type" ID="txtMed_type1" TAG="21X" VALUE="1" CHECKED><LABEL FOR="txtMed_type1">보수총액</LABEL>&nbsp;&nbsp;&nbsp;
				                			                <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtMed_type" ID="txtMed_type2" TAG="21X" VALUE="2"><LABEL FOR="txtMed_type2">표준보수월액</LABEL></TD>
                                        </TR>
                                        
                                        
				    		        <% Call SubFillRemBodyTD56(2) %>
				    		        </TABLE>
				    		        </FIELDSET>
						        </TD>
				    	        <TD VALIGN="TOP">
						            <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT>국민연금처리</LEGEND>
						            <TABLE CLASS="BasicTB" CELLSPACING=0>
        					        	<TR>
              							    <TD CLASS="TD5" NOWRAP></TD>
	                   						<TD CLASS="TD6">&nbsp;&nbsp;1년이상자&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1년미만자</TD>
	                   					</TR>
        					        	<TR>
              							    <TD CLASS="TD5" NOWRAP>회사부담율</TD>
	                   						<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtAnut_comp_rate1 NAME=txtAnut_comp_rate1 CLASS=FPDS90 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="회사부담율"></OBJECT>');</SCRIPT>%&nbsp; 
	                   						                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtAnut_comp_rate2 NAME=txtAnut_comp_rate2 CLASS=FPDS90 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="회사부담율"></OBJECT>');</SCRIPT>%</TD>
	                   					</TR>
			        		        	<TR>
					    	           		<TD CLASS="TD5" NOWRAP>본인부담율</TD>
						            		<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtAnut_prsn_rate1 NAME=txtAnut_prsn_rate1 CLASS=FPDS90 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="본인부담율"></OBJECT>');</SCRIPT>%&nbsp; 
						            		                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtAnut_prsn_rate2 NAME=txtAnut_prsn_rate2 CLASS=FPDS90 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="본인부담율"></OBJECT>');</SCRIPT>%</TD>
						    	        </TR>
			        		        	<TR>
					    	           		<TD CLASS="TD5" NOWRAP>퇴직전환율</TD>
						            		<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtAnut_retire_rate1 NAME=txtAnut_retire_rate1 CLASS=FPDS90 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="퇴직전환율"></OBJECT>');</SCRIPT>%&nbsp; 
						            		                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtAnut_retire_rate2 NAME=txtAnut_retire_rate2 CLASS=FPDS90 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="퇴직전환율"></OBJECT>');</SCRIPT>%</TD>
						    	        </TR>
        				        		<TR>
		        		        			<TD CLASS="TD5" NOWRAP>중도입사자처리</TD>
				                			<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtanut_entr_flag" TAG="21" VALUE="Y" CHECKED ID="txtanut_entr_flag1"><LABEL FOR="txtanut_entr_flag1">공제함</LABEL>&nbsp;&nbsp;&nbsp;
				                			                <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtanut_entr_flag" TAG="21" VALUE="N" ID="txtanut_entr_flag2"><LABEL FOR="txtanut_entr_flag2">공제안함</LABEL></TD>
                                        </TR>
        				        		<TR>
		        		        			<TD CLASS="TD5" NOWRAP>중도퇴사자처리</TD>
				                			<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtanut_retire_flag" TAG="21" VALUE="Y" CHECKED ID="txtanut_retire_flag1"><LABEL FOR="txtanut_retire_flag1">공제함</LABEL>&nbsp;&nbsp;&nbsp;
				                			                <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtanut_retire_flag" TAG="21" VALUE="N" ID="txtanut_retire_flag2"><LABEL FOR="txtanut_retire_flag2">공제안함</LABEL></TD>
                                        </TR>
        				        		<TR>
		        		        			<TD CLASS="TD5" NOWRAP>중도입퇴사자처리</TD>
				                			<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtanut_en_re_flag" TAG="21" VALUE="Y" CHECKED ID="txtanut_en_re_flag1"><LABEL FOR="txtanut_en_re_flag1">공제함</LABEL>&nbsp;&nbsp;&nbsp;
				                			                <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtanut_en_re_flag" TAG="21" VALUE="N" ID="txtanut_en_re_flag2"><LABEL FOR="txtanut_en_re_flag2">공제안함</LABEL></TD>
                                        </TR>
        				        		<TR HEIGHT=6>
		        		        			<TD CLASS="TD5"></TD>
				                			<TD CLASS="TD6"></TD>
                                        </TR>
				    		        </TABLE>
				    		        </FIELDSET>
						        </TD>
                            </TR>
				    	    <TR>
				        	    <TD VALIGN="TOP" COLSPAN=2>
						            <FIELDSET CLASS="CLSFLD">
						            <TABLE CLASS="BasicTB" CELLSPACING=0>
	    				        	    <TR>
              							    <TD CLASS="TD5" NOWRAP>고용보험율</TD>
	                   						<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtEmploy_rate NAME=txtEmploy_rate CLASS=FPDS65 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="회사부담율"></OBJECT>');</SCRIPT>%</TD>
              							    <TD CLASS="TD5" NOWRAP>급여지급일</TD>
	                   						<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtpay_prov_dd NAME=txtpay_prov_dd CLASS=FPDS40 TITLE=FPDOUBLESINGLE TAG="22X81" ALT="급여지급일"></OBJECT>');</SCRIPT>
	                   						</TD>
	        						    </TR>
	    				        	    <TR>
              							    <TD CLASS="TD5" NOWRAP>퇴직세액공제</TD>
	                   						<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtRe_tax_sub1 NAME=txtRe_tax_sub1 CLASS=FPDS65 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="퇴직세액공제"></OBJECT>');</SCRIPT>%</TD>
              							    <TD CLASS="TD5" NOWRAP>급여기준일</TD>
	                   						<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtpay_bas_dd NAME=txtpay_bas_dd CLASS=FPDS40 TITLE=FPDOUBLESINGLE TAG="22X81" ALT="급여기준일"></OBJECT>');</SCRIPT>
	                   						</TD>
	        						    </TR>
	    				        	    <TR>
              							    <TD CLASS="TD5" NOWRAP>퇴직특별공제</TD>
	                   						<TD CLASS="TD6">기타퇴직&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtRe_incom_sub NAME=txtRe_incom_sub CLASS=FPDS65 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="기타퇴직"></OBJECT>');</SCRIPT>%</TD>
              							    <TD CLASS="TD5" NOWRAP>근태기준일</TD>
	                   						<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtdilig_dd NAME=txtdilig_dd CLASS=FPDS40 TITLE=FPDOUBLESINGLE tag="22X81" ALT="근태기준일"></OBJECT>');</SCRIPT>
	                   						</TD>
	        						    </TR>
                                        <TR>
              							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
	                   						<TD CLASS="TD6">명예퇴직&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtRe_speci_sub NAME=txtRe_speci_sub CLASS=FPDS65 TITLE=FPDOUBLESINGLE TAG="21X9Z" ALT="명예퇴직"></OBJECT>');</SCRIPT>%</TD>
              							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
	                   						<TD CLASS="TD6"></TD>
	        						    </TR>
	    				        	    <TR>
              							    <TD CLASS="TD5" NOWRAP>퇴직세액한도액=</TD>
	                   						<TD CLASS="TD6">근속연수 * &nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> ID=txtre_sub_limit NAME=txtre_sub_limit CLASS=FPDS140 TITLE=FPDOUBLESINGLE TAG="21X2Z" ALT="퇴직세액한도액"></OBJECT>');</SCRIPT></TD>
              							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
	                   						<TD CLASS="TD6"></TD>
	        						    </TR>
							        </TABLE>
				    		        </FIELDSET>
				        	    </TD>
			    		    </TR>
			    		    <TR><TD>&nbsp;</TD></TR>
			    		    <TR><TD>&nbsp;</TD></TR>
			    		    <TR><TD>&nbsp;</TD></TR>
			    		    <TR><TD>&nbsp;</TD></TR>
			    		    <TR><TD>&nbsp;</TD></TR>
					    </TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>

