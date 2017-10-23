
<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : VAT
'*  3. Program ID           : a6117ma1
'*  4. Program Name         : �ΰ������� 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004.05.10
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Eun Kyung , KANG
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '#########################################################################################################
'1. �� �� �� 
'##############################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/IncSvrCcm.inc"  --><!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css"><!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit'��: indicates that All variables must be declared in advance

'******************************************  1.2 Global ����/��� ����  ***********************************
'1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID = "a6117mb1.asp"'��: �����Ͻ� ���� ASP�� 

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns

Dim C_VAT_NO              
Dim C_ISSUED_DT
Dim C_IO_FG            
Dim C_BP_CD 
Dim C_BP_PB          
Dim C_BP_NM 
DIM C_REG_NO    
Dim C_MADE_VAT_FG              
Dim C_VAT_TYPE      
Dim C_VAT_TYPE_NM  
Dim C_VAT_TYPE_PB     
Dim C_NET_LOC_AMT        
Dim C_VAT_LOC_AMT         
Dim C_CARD_NO 
Dim C_CARD_PB              
Dim C_REPORT_BIZ_AREA_CD  
Dim C_REPORT_BIZ_AREA_PB
Dim C_BIZ_AREA_CD   
Dim C_BIZ_AREA_PB
Dim C_GL_NO     
Dim C_TEMP_GL_NO


Dim C_issue_dt_fg_cd
Dim C_issue_dt_fg_nm
Dim C_issue_dt_kind_cd
Dim C_issue_dt_kind_nm

Const C_SHEETMAXROWS = 100

 '==========================================  1.2.2 Global ���� ����  =====================================
'1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
'Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
'Dim lgIntFlgMode               ' Variable is for Operation Status

'Dim lgStrPrevKey
'Dim lgLngCurRows

Dim lgStrPrevVatKey

 '==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
 '----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop
'Dim lgSortKey

 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
'Dim lgStrComDateType'Company Date Type�� ����(��� Mask�� �����.)

 '#########################################################################################################
'2. Function�� 
'
'���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'           2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : initSpreadPosVariables()
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

   C_VAT_NO                = 1     
	C_ISSUED_DT             = 2  
	C_IO_FG                 = 3  
	C_BP_CD                 = 4  
	C_BP_PB                 = 5  
	C_BP_NM                 = 6  
	C_REG_NO                = 7  
	C_MADE_VAT_FG           = 8  
	C_VAT_TYPE              = 9  
	C_VAT_TYPE_NM           = 10
	C_VAT_TYPE_PB           = 11
	C_issue_dt_fg_cd        = 12
	C_issue_dt_fg_nm        = 13
	C_issue_dt_kind_cd      = 14
	C_issue_dt_kind_nm      = 15
	C_NET_LOC_AMT           = 16
	C_VAT_LOC_AMT           = 17
	C_CARD_NO               = 18
	C_CARD_PB               = 19
	C_REPORT_BIZ_AREA_CD    = 20
	C_REPORT_BIZ_AREA_PB    = 21
	C_BIZ_AREA_CD           = 22
	C_BIZ_AREA_PB           = 23
	C_GL_NO                 = 24
	C_TEMP_GL_NO            = 25
     					
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$




 '==========================================  2.1.1 InitVariables()  ======================================
'Name : InitVariables()
'Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevVatKey = ""
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgPageNo  = 0
End Sub

 '******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'���: ȭ���ʱ�ȭ 
'����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 

'==========================================  2.2.1 SetDefaultVal()  ========================================
'Name : SetDefaultVal()
'Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
Dim strSvrDate
Dim strYear, strMonth, strDay,  EndDate, StartDate

    strSvrDate = "<%=GetSvrDate%>"

	Call ExtractDateFrom(strSvrDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	EndDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	StartDate = UNIDateAdd("M", -1, EndDate, parent.gDateFormat)

	frm1.txtIssuedDtFr.Text = StartDate
	frm1.txtIssuedDtTo.Text = EndDate

End Sub



'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================== 2.2.3 InitSpreadSheet() =================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()

    Call initSpreadPosVariables()    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20091212",,parent.gAllowDragDropSpread


    With frm1.vspdData
		
        .ReDraw = False
        '.ColHidden = True
        
       .MaxCols = C_TEMP_GL_NO + 1
        '.Col = .MaxCols'��: ������Ʈ�� ��� Hidden Column

        .MaxRows = 0
                           'patch version
        Call GetSpreadColumnPos("A")
        
        ggoSpread.SSSetEdit      C_VAT_NO,             "��꼭��ȣ",   18, 3
        ggoSpread.SSSetDate      C_ISSUED_DT,          "������",       10, 2, parent.gDateFormat
        ggoSpread.SSSetEdit      C_IO_FG,              "���ⱸ��",     8, 3
        ggoSpread.SSSetEdit      C_BP_CD,              "�ŷ�ó�ڵ�",   10, 3
        ggoSpread.SSSetButton    C_BP_PB
        ggoSpread.SSSetEdit      C_BP_NM,              "�ŷ�ó��",     20, 3
        ggoSpread.SSSetEdit      C_REG_NO,              "����ڹ�ȣ",     12, 3
        
        ggoSpread.SSSetEdit      C_MADE_VAT_FG,        "�ΰ�������",   2, 3
        ggoSpread.SSSetEdit      C_VAT_TYPE,           "",                 2, 3        
        ggoSpread.SSSetEdit      C_VAT_TYPE_NM,        "�ΰ�������",   15, 3
        ggoSpread.SSSetButton    C_VAT_TYPE_PB
        ggoSpread.SSSetCombo     C_issue_dt_fg_cd,  "���ڼ��ݰ�꼭���࿩��",	15
		ggoSpread.SSSetCombo     C_issue_dt_fg_nm,  "���ڼ��ݰ�꼭���࿩��",	15    
        ggoSpread.SSSetCombo     C_issue_dt_kind_cd,  "���ڼ��ݰ�꼭����",	15
		ggoSpread.SSSetCombo     C_issue_dt_kind_nm,  "���ڼ��ݰ�꼭����",	15        
        ggoSpread.SSSetFloat     C_NET_LOC_AMT,        "���ް���",     20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec
        ggoSpread.SSSetFloat     C_VAT_LOC_AMT,        "�ΰ�����",     20, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec
        ggoSpread.SSSetEdit      C_CARD_NO,            "ī���ȣ",     20, 3
        ggoSpread.SSSetButton    C_CARD_PB
        ggoSpread.SSSetEdit      C_REPORT_BIZ_AREA_CD, "�Ű�����",   10, 3
        ggoSpread.SSSetButton    C_REPORT_BIZ_AREA_PB
        ggoSpread.SSSetEdit      C_BIZ_AREA_CD ,       "�߻������",   10, 3
        ggoSpread.SSSetButton    C_BIZ_AREA_PB
        ggoSpread.SSSetEdit      C_GL_NO,              "��ǥ��ȣ",     10, 5
        ggoSpread.SSSetEdit      C_TEMP_GL_NO,         "������ȣ", 10, 5

		
        

        
		Call ggoSpread.MakePairsColumn(C_BP_CD,              C_BP_PB              ,"1")
		Call ggoSpread.MakePairsColumn(C_VAT_TYPE,           C_VAT_TYPE_PB        ,"1")
		Call ggoSpread.MakePairsColumn(C_CARD_NO           , C_CARD_PB            ,"1")
		Call ggoSpread.MakePairsColumn(C_REPORT_BIZ_AREA_CD, C_REPORT_BIZ_AREA_PB ,"1")
		Call ggoSpread.MakePairsColumn(C_BIZ_AREA_CD,        C_BIZ_AREA_PB        ,"1")
		Call ggoSpread.MakePairsColumn(C_issue_dt_kind_cd, C_issue_dt_kind_nm, "1")
		Call ggoSpread.MakePairsColumn(C_issue_dt_fg_cd, C_issue_dt_fg_nm, "1")

		Call ggoSpread.SSSetColHidden(C_issue_dt_kind_cd, C_issue_dt_kind_cd, True)
		Call ggoSpread.SSSetColHidden(C_issue_dt_fg_cd, C_issue_dt_fg_cd, True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
        Call ggoSpread.SSSetColHidden(C_VAT_TYPE,C_VAT_TYPE,True)
        Call ggoSpread.SSSetColHidden(C_MADE_VAT_FG,C_MADE_VAT_FG,True)
		
		
        .ReDraw = True

    
    End With
    
    Call SetSpreadLock     
End Sub

'================================== 2.2.4 SetSpreadLock() ===============================
' Function Name : SetSpreadLock()

' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()

    With frm1.vspdData
        .ReDraw = False

        'ggoSpread.SSSetRequired C_BDG_PLAN_AMT, -1,-1
        ggoSpread.SpreadLock C_VAT_NO,      -1, C_VAT_NO
        ggoSpread.SpreadLock C_IO_FG,       -1, C_IO_FG
        ggoSpread.SpreadLock C_BP_NM,       -1, C_BP_NM      ' ,-1
        ggoSpread.SpreadLock C_REG_NO,       -1, C_REG_NO      ' ,-1
        ggoSpread.SpreadLock C_GL_NO,       -1, C_GL_NO
        ggoSpread.SpreadLock C_TEMP_GL_NO,  -1, C_TEMP_GL_NO       ',-1
        
        ggoSpread.SSSetProtected .MaxCols,  -1, -1
        .ReDraw = True
		
    End With

End Sub


'================================== 2.2.5 SetSpreadColor() ==============================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

    With frm1

        .vspdData.ReDraw = False

        ' �ʼ� �Է� �׸����� ���� 
        ' SSSetRequired(ByVal Col, ByVal Row, Optional Row2)
        ggoSpread.SSSetProtected  C_VAT_NO,              pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_ISSUED_DT,           pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_IO_FG,               pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_BP_CD,               pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_BP_NM,               pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_REG_NO,               pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_MADE_VAT_FG,         pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_VAT_TYPE_NM,         pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_NET_LOC_AMT,         pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_VAT_LOC_AMT,         pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_CARD_NO,             pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_REPORT_BIZ_AREA_CD,  pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_BIZ_AREA_CD,         pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_GL_NO,               pvStartRow, pvEndRow
        ggoSpread.SSSetProtected  C_TEMP_GL_NO,          pvStartRow, pvEndRow

        .vspdData.ReDraw = True
    
    End With

End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'Name : InitComboBox()
'Description : Combo Display
'========================================================================================================= 



Sub InitComboBox()

    Dim arrData
 
   ggoSpread.Source = frm1.vspdData

	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("DT004", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1

	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_issue_dt_kind_cd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_issue_dt_kind_nm
 
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1020", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1

	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_issue_dt_fg_cd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_issue_dt_fg_nm

    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1003", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboIoFg ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("DT004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboissue_dt_kind ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboissue_dt_fg ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("DT004", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboissue_dt_kind2 ,lgF0  ,lgF1  ,Chr(11))
    
    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1020", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboissue_dt_fg2 ,lgF0  ,lgF1  ,Chr(11))
    
    
End Sub



'========================================== 2.4.2 Open???()  =============================================
'Name : Open???()
'Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 
'--------------------------------------------------------------------------------------------------------- 
'   Function Name : OpenVatNoInfo()
'   Function Desc : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenVatNoInfo(Byval strCode, Byval Cond)
	Dim iCalledAspName
	Dim arrRet
		
	If IsOpenPop = True Then Exit Function	

	iCalledAspName = AskPRAspName("a6114ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a6114ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	     
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtVatNo.focus
		Exit Function
	Else
		Call SetVatNoInfo(arrRet,Cond)	
	End If	
End Function

'--------------------------------------------------------------------------------------------------------- 
'   Function Name : SetChgNoInfo(Byval arrRet)
'   Function Desc : 
'--------------------------------------------------------------------------------------------------------- 
Function SetVatNoInfo(Byval arrRet, Byval Cond)
	Select Case Cond
		Case "VatNo"
			frm1.txtVatNo.focus
			frm1.txtVatNo.Value	= arrRet(0)
	End Select	
End Function

'------------------------------------------  OpenBp()  ---------------------------------------
'Name : OpenBp()
'Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
    Dim arrRet
    Dim arrParam(5)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True
        
    arrParam(0) = strCode' Code Condition
    arrParam(1) = ""' ä�ǰ� ����(�ŷ�ó ����)
    arrParam(2) = ""' FrDt
    arrParam(3) = ""' ToDt
    arrParam(4) = "T"' B :���� S: ���� T: ��ü 
    arrParam(5) = ""' SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 

    arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
    "dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False

    If arrRet(0) = "" Then
        frm1.txtBpCd.focus
        Exit Function
    Else
	    frm1.txtBpCd.focus
	    frm1.txtBpCd.Value    = arrRet(0)		
    	frm1.txtBpNm.Value    = arrRet(1)		
        lgBlnFlgChgValue = True
    End If
End Function

'=======================================================================================================
'    Name : OpenReportBizArea()
'    Description : Bp Cd PopUp
'=======================================================================================================
Function OpenReportBizArea()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
    
    If IsOpenPop = True  Then Exit Function

    IsOpenPop = True

    arrParam(0) = "���ݽŰ����� �˾�"                    ' �˾� ��Ī 
    arrParam(1) = "B_TAX_BIZ_AREA"                        ' TABLE ��Ī 
    arrParam(2) = Trim(frm1.txtReportBizArea.Value)
    arrParam(3) = ""
    arrParam(4) = ""            
    arrParam(5) = "���ݽŰ������ڵ�"                    '�����ʵ��� �� ��Ī 
    
    arrField(0) = "TAX_BIZ_AREA_CD"                               ' Field��(0)
    arrField(1) = "TAX_BIZ_AREA_NM"                               ' Field��(1)
    
    arrHeader(0) = "���ݽŰ������ڵ�"                       ' Header��(0)
    arrHeader(1) = "���ݽŰ������"                       ' Header��(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        frm1.txtReportBizArea.focus    
        Exit Function
    Else
        Call SetReportBizArea(arrRet)
    End If    
End Function

'=======================================================================================================
'    Name : SetReportBizArea()
'    Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetReportBizArea(byval arrRet)
    frm1.txtReportBizArea.focus    
    frm1.txtReportBizArea.Value    = arrRet(0)        
    frm1.txtReportBizAreaNm.Value    = arrRet(1)        
    lgBlnFlgChgValue = True
End Function

'=======================================================================================================
'    Name : OpenVatType()
'    Description : Bp Cd PopUp
'=======================================================================================================
Function OpenVatType()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)
      
    If IsOpenPop = True  Then Exit Function

    IsOpenPop = True
    arrParam(0) = "�ΰ��������˾�"                    ' �˾� ��Ī 
    arrParam(1) = "B_MINOR"                                ' TABLE ��Ī 
    arrParam(2) = Trim(frm1.txtVatType.Value)
    arrParam(3) = ""
    arrParam(4) = "MAJOR_CD=" & FilterVar("B9001", "''", "S") & " "            
    arrParam(5) = "�ΰ����ڵ�"                    '�����ʵ��� �� ��Ī 
    
    arrField(0) = "MINOR_CD"                               ' Field��(0)
    arrField(1) = "MINOR_NM"                               ' Field��(1)
    
    arrHeader(0) = "�ΰ�������"                       ' Header��(0)
    arrHeader(1) = "�ΰ���������"                       ' Header��(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    If arrRet(0) = "" Then
        frm1.txtVatType.focus
        Exit Function
    Else
        Call SetVatType(arrRet)
    End If   
End Function

'=======================================================================================================
'    Name : SetVatType()
'    Description :
'=======================================================================================================
Function SetVatType(byval arrRet)
    frm1.txtVatType.focus
    frm1.txtVatType.Value   = arrRet(0)        
    frm1.txtVatTypeNm.Value = arrRet(1)        
    lgBlnFlgChgValue = True
End Function


'============================================================
'���� �˾� 
'============================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    Select Case iWhere
    
        Case "BpCd_Spread"
           arrParam(0) = strCode                                ' Code Condition
           arrParam(1) = ""                                     ' ä�ǰ� ����(�ŷ�ó ����)
           arrParam(2) = ""                                     ' FrDt
           arrParam(3) = ""                                     ' ToDt
           arrParam(4) = "T"                                    ' B :���� S: ���� T: ��ü 
           arrParam(5) = ""                                     ' SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 

        Case "VatType_Spread"
            arrParam(0) = "�ΰ��������˾�"                 ' �˾� ��Ī 
            arrParam(1) = "B_MINOR "                            ' TABLE ��Ī 
            arrParam(2) = strCode                               ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = "MAJOR_CD=" & FilterVar("B9001", "''", "S") & " "                    ' Where Condition
            arrParam(5) = "�ΰ�������"                  ' �����ʵ��� �� ��Ī 

            arrField(0) = "MINOR_CD"                            ' Field��(0)
            arrField(1) = "MINOR_NM"                            ' Field��(1)
    
            arrHeader(0) = "�ΰ�������"                     ' Header��(0)
            arrHeader(1) = "�ΰ���������"                   ' Header��(1)
    
        Case "CardCd_Spread"
            arrParam(0) = "�ſ�ī�� �˾�"                   ' �˾� ��Ī 
            arrParam(1) = "B_CREDIT_CARD"                       ' TABLE ��Ī 
            arrParam(2) = strCode                               ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""                                    ' Where Condition
            arrParam(5) = "�ſ�ī��"                        ' �����ʵ��� �� ��Ī 

            arrField(0) = "CREDIT_NO"                           ' Field��(0)
            arrField(1) = "CREDIT_NM"                           ' Field��(1)
    
            arrHeader(0) = "�ſ�ī���ȣ"                   ' Header��(0)
            arrHeader(1) = "�ſ�ī���"                     ' Header��(1)
    
        Case "ReportBizAreaCd_Spread"
            arrParam(0) = "���ݽŰ����� �˾�"             ' �˾� ��Ī 
            arrParam(1) = "B_TAX_BIZ_AREA"                      ' TABLE ��Ī 
            arrParam(2) = strCode                               ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""
            arrParam(5) = "���ݽŰ������ڵ�"            
    
            arrField(0) = "TAX_BIZ_AREA_CD"                     ' Field��(0)
            arrField(1) = "TAX_BIZ_AREA_NM"                     ' Field��(1)

            arrHeader(0) = "���ݽŰ������ڵ�"             ' Header��(0)
            arrHeader(1) = "���ݽŰ������"               ' Header��(1)
 
         Case "BizAreaCd_Spread"
            arrParam(0) = "����� �˾�"                     ' �˾� ��Ī 
            arrParam(1) = "B_BIZ_AREA"                          ' TABLE ��Ī 
            arrParam(2) = strCode                               ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""
            arrParam(5) = "������ڵ�"            
    
            arrField(0) = "BIZ_AREA_CD"                          ' Field��(0)
            arrField(1) = "BIZ_AREA_NM"                          ' Field��(1)

            arrHeader(0) = "������ڵ�"                     ' Header��(0)
            arrHeader(1) = "������"                       ' Header��(1)
       
    End Select    

    IsOpenPop = True
    
    If iWhere = "BpCd_Spread" Then
       arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
           "dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
    Else
        arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
            "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    End If
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
        Exit Function
    Else
        With frm1
            Select Case iWhere
            Case "BpCd_Spread"
                .vaSpread1.Col  = C_BP_CD
                .vaSpread1.Text = arrRet(0)
                .vaSpread1.Col  = C_BP_NM
                .vaSpread1.Text = arrRet(1)
                '.vaSpread1.Col  = C_REG_NO
                '.vaSpread1.Text = arrRet(2)

                Call vspdData_Change(.vspdData.Col,.vspdData.Row )

            Case "VatType_Spread"
                .vaSpread1.Col  = C_VAT_TYPE
                .vaSpread1.Text = arrRet(0)
                .vaSpread1.Col  = C_VAT_TYPE_NM
                .vaSpread1.Text = arrRet(1)
                Call vspdData_Change(.vspdData.Col,.vspdData.Row )

            Case "CardCd_Spread"
                .vaSpread1.Col  = C_CARD_NO
                .vaSpread1.Text = arrRet(0)            
                Call vspdData_Change(.vspdData.Col,.vspdData.Row )

            Case "ReportBizAreaCd_Spread"
                .vaSpread1.Col  = C_REPORT_BIZ_AREA_CD
                .vaSpread1.Text = arrRet(0)
                Call vspdData_Change(.vspdData.Col,.vspdData.Row )
                
            Case "BizAreaCd_Spread"
                .vspdData.Col  = C_BIZ_AREA_CD
                .vspdData.Text = arrRet(0)
                Call vspdData_Change(.vspdData.Col,.vspdData.Row )         

            End Select
        End With
    End If    

End Function


'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
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
            		C_VAT_NO                = iCurColumnPos(1)      
				C_ISSUED_DT             = iCurColumnPos(2)   
				C_IO_FG                 = iCurColumnPos(3)   
				C_BP_CD                 = iCurColumnPos(4)   
				C_BP_PB                 = iCurColumnPos(5)   
				C_BP_NM                 = iCurColumnPos(6)   
				C_REG_NO                = iCurColumnPos(7)   
				C_MADE_VAT_FG           = iCurColumnPos(8)   
				C_VAT_TYPE              = iCurColumnPos(9)   
				C_VAT_TYPE_NM           = iCurColumnPos(10)  
				C_VAT_TYPE_PB           = iCurColumnPos(11)  
				C_issue_dt_fg_cd        = iCurColumnPos(12)  
				C_issue_dt_fg_nm        = iCurColumnPos(13)  
				C_issue_dt_kind_cd      = iCurColumnPos(14)  
				C_issue_dt_kind_nm      = iCurColumnPos(15)  
				C_NET_LOC_AMT           = iCurColumnPos(16)  
				C_VAT_LOC_AMT           = iCurColumnPos(17)  
				C_CARD_NO               = iCurColumnPos(18)  
				C_CARD_PB               = iCurColumnPos(19)  
				C_REPORT_BIZ_AREA_CD    = iCurColumnPos(20)  
				C_REPORT_BIZ_AREA_PB    = iCurColumnPos(21)  
				C_BIZ_AREA_CD           = iCurColumnPos(22)  
				C_BIZ_AREA_PB           = iCurColumnPos(23)  
				C_GL_NO                 = iCurColumnPos(24)  
				C_TEMP_GL_NO            = iCurColumnPos(25)  
			
            
    End Select    
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '#########################################################################################################
'                                                3. Event�� 
'    ���: Event �Լ��� ���� ó�� 
'    ����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################
 '******************************************  3.1 Window ó��  *********************************************
'    Window�� �߻� �ϴ� ��� Even ó��    
'********************************************************************************************************* 
 '==========================================  3.1.1 Form_Load()  ======================================
'    Name : Form_Load()
'    Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    
    Call LoadInfTB19029                           '��: Load table , B_numeric_format        
    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)    
    Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking
    Call SetDefaultVal
	 
    Call InitSpreadSheet                          '��: Setup the Spread Sheet    
    Call InitComboBox
    Call InitVariables     
    

    '----------  Coding part  -------------------------------------------------------------
    'Call FncSetToolBar("New")
    Call SetToolbar("1100100100101111")

    frm1.txtIssuedDtFr.focus
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'    Document�� TAG���� �߻� �ϴ� Event ó��    
'    Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'    Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag ó��  *********************************************
'    Window�� �߻� �ϴ� ��� Even ó��    
'********************************************************************************************************* 

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

Function txtReportBizArea_onblur()
    If frm1.txtReportBizArea.value = "" Then
        frm1.txtReportBizAreaNm.value = ""
    End If
End Function

Function txtBpCd_onblur()
    If frm1.txtBpCd.value = "" Then
        frm1.txtBpNm.value = ""
    End If
End Function

'=======================================================================================================
'   Event Name : txtIssuedDtFr_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtIssuedDtFr_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDtFr.Action = 7
          Call SetFocusToDocument("M")
        frm1.txtIssuedDtFr.Focus
    End If
End Sub

Sub txtIssuedDtTo_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDtTo.Action = 7
          Call SetFocusToDocument("M")
        frm1.txtIssuedDtTo.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtIssuedDtFr_Keypress(Key)
'   Event Desc : ��ȸ�� �Ѵ�.
'=======================================================================================================
Sub txtIssuedDtFr_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssuedDtTo.focus
        FncQuery()
    End If
End Sub

Sub txtIssuedDtTo_Keypress(Key)
    If Key = 13 Then
		frm1.txtIssuedDtFr.focus
        FncQuery()
    End If
End Sub


Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
		.Row = Row

		Select Case Col
			Case  C_issue_dt_kind_nm
				.Col = Col
				intIndex = .Value
				.Col = C_issue_dt_kind_cd
				.Value = intIndex
			Case  C_issue_dt_fg_nm
				.Col = Col
				intIndex = .Value
				.Col = C_issue_dt_fg_cd
				.Value = intIndex
		End Select
	End With
End Sub



'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(Col, Row)

    If lgIntFlgMode = Parent.OPMD_CMODE Then
        Call SetPopupMenuItemInf("1001111111")
    Else
        Call SetPopupMenuItemInf("1101111111")
    End If       
    gMouseClickStatus = "SPC"    'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
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
        Exit Sub

    End If
    
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
    If Row <= 0 Then
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

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
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    Dim strSelect, strFrom, strWhere
    Dim strYear, strMonth, strDay, strDate
    Dim IntRetCD 
    Dim arrVal1, arrVal2
    Dim ii
    Dim intIndex

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    frm1.vspdData.row = Row

    Select Case Col
        Case  C_issue_dt_kind_nm
            frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_issue_dt_kind_cd
            frm1.vspdData.Value = intIndex
            
        Case  C_issue_dt_kind_cd
              frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_issue_dt_kind_nm
            frm1.vspdData.Value = intIndex
            
        Case  C_issue_dt_fg_nm
            frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_issue_dt_fg_cd
            frm1.vspdData.Value = intIndex
        
        Case  C_issue_dt_fg_cd
              frm1.vspdData.Col = Col
            intIndex = frm1.vspdData.Value
            frm1.vspdData.Col = C_issue_dt_fg_nm
            frm1.vspdData.Value = intIndex
                 
    End Select
    
    lgBlnFlgChgValue = True
    
End Sub

Sub vspdData_ScriptLeaveCell(Col, Row, NewCol, NewRow, Cancel)

    If frm1.vspdData.MaxRows = 0 Then                            'no data�� ��� vspdData_LeaveCell no ���� 
       Exit Sub                                                    'tab�̵��ÿ� �߸��� 140318 message ���� 
    End If
    
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
 '----------  Coding part  -------------------------------------------------------------   
    If CheckRunningBizProcess = True Then
       Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
        If lgStrPrevVatKey <> "" then  
          Call DisableToolBar(Parent.TBC_QUERY)
            If DbQuery = False Then
                Call RestoreToolBar()
                Exit Sub
            End if
       End If
    End if
        
End Sub

'==========================================================================================
' Event Name : vspdData_ButtonClicked
' Event Desc : ��ư �÷��� Ŭ���� ��� 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
    '---------- Coding part -------------------------------------------------------------
    With frm1.vspdData
    
        ggoSpread.Source = frm1.vspdData
        
        IF Row > 0 And Col = C_VAT_TYPE_PB Then
            .Col = C_VAT_TYPE
            .Row = Row
            Call OpenPopup(.Text, "VatType_Spread")
                
        ElseIf Row > 0 and Col = C_CARD_PB Then
            .Col = C_CARD_NO
            .Row = Row
            Call OpenPopup(.Text, "CardCd_Spread")

        ElseIf Row > 0 and Col = C_REPORT_BIZ_AREA_PB Then
            .Col = C_REPORT_BIZ_AREA_CD
            .Row = Row
            Call OpenPopup(.Text, "ReportBizAreaCd_Spread")

        ElseIf Row > 0 and Col = C_BIZ_AREA_PB Then
            .Col = C_BIZ_AREA_CD
            .Row = Row
            Call OpenPopup(.Text, "BizAreaCd_Spread")
        
        ElseIf Row > 0 and Col = C_BP_PB Then
            .Col = C_BP_CD
            .Row = Row
            Call OpenPopup(.Text, "BpCd_Spread")
        
        End If
        
    End With
    
End Sub


'#########################################################################################################
'                                                4. Common Function�� 
'    ���: Common Function
'    ����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 

'#########################################################################################################
'                                                5. Interface�� 
'    ���: Interface
'    ����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'          Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'    << ���뺯�� ���� �κ� >>
'     ���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'                �����ϵ��� �Ѵ�.
'     1. ������Ʈ���� Call�ϴ� ���� 
'           ADF (ADS, ADC, ADF�� �״�� ���)
'           - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
'     2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'            strRetMsg
'######################################################################################################### 

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'    ���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim strFrYear, strFrMonth, strFrDay 
    Dim strToYear, strToMonth, strToDay
    
    FncQuery = False          '��: Processing is NG
    Err.Clear                 '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")            '��: "Will you destory previous data"
        if IntRetCD = vbNo Then
            Exit Function
        End If
    End If
   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    Call InitVariables                              '��: Initializes local global variables
   
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then    '��: This function check indispensable field
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtIssuedDtFr.Text, frm1.txtIssuedDtTo.Text, frm1.txtIssuedDtFr.Alt, frm1.txtIssuedDtTo.Alt, _
                        "970025", frm1.txtIssuedDtFr.UserDefinedFormat, parent.gComDateType, true) = False Then
            frm1.txtBdgYymmFr.focus                                                        '��: GL Date Compare Common Function
            Exit Function
    End if

    Call ExtractDateFrom(frm1.txtIssuedDtFr.Text,frm1.txtIssuedDtFr.UserDefinedFormat,parent.gComDateType,strFrYear,strFrMonth,strFrDay)    
        
    Call ExtractDateFrom(frm1.txtIssuedDtTo.Text,frm1.txtIssuedDtTo.UserDefinedFormat,parent.gComDateType,strToYear,strToMonth,strToDay)
     
   
    Call FncSetToolBar("New")
    '-----------------------
    'Query function call area
    '-----------------------

    Call DbQuery        
                                                                                '��: Query db data
       
    FncQuery = True                                                                '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                  '��: Processing is NG
    Err.Clear                       '��: Protect system from crashing
    'On Error Resume Next            '��: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ' ����� ������ �ִ��� Ȯ���Ѵ�.
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015",parent.VB_YES_NO,"X","X")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")     '��: Clear Condition Field    
    Call InitVariables                         '��: Initializes local global variables
    Call SetDefaultVal
    
    Call FncSetToolBar("New")
    
    'SetGridFocus
    FncNew = True                              '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False            '��: Processing is NG
    Err.Clear                  '��: Protect system from crashing
    'On Error Resume Next       '��: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then                   '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '��: Display Message(There is no changed data.)
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '��: Check required field(Multi area)
       Exit Function
    End If

    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave                                                                  '��: Save db data

     FncSave = True                                                           '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
    If frm1.vspdData.MaxRows < 1 Then Exit Function
    
    ggoSpread.Source = frm1.vspdData    
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
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
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call Parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call Parent.FncExport(parent.C_MULTI)                                                '��: ȭ�� ���� 
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

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
Dim IntRetCD
    
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'    ���� : 
'********************************************************************************************************* 

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Dim strVal

    Call LayerShowHide(1)
    
    DbQuery = False
    Err.Clear                '��: Protect system from crashing
    
    With frm1
       
            If lgIntFlgMode = parent.OPMD_CMODE Then
				IF Trim(.txtVatNo.value) <> "" Then
					lgStrPrevVatKey = Trim(.txtVatNo.value)
				END IF
			End if           
            strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
            strVal = strVal & "&txtIssuedDtFr=" & Trim(.txtIssuedDtFr.text)            '��ȸ ���� ����Ÿ 
            strVal = strVal & "&txtIssuedDtTo=" & Trim(.txtIssuedDtTo.text)                '��ȸ ���� ����Ÿ 
            strVal = strVal & "&txtReportBizArea=" & Trim(.txtReportBizArea.value)                    '��ȸ ���� ����Ÿ 
            strVal = strVal & "&cboIoFg=" & Trim(.cboIoFg.value)                    '��ȸ ���� ����Ÿ 
            strVal = strVal & "&txtVatType=" & Trim(.txtVatType.value)                    '��ȸ ���� ����Ÿ 
            strVal = strVal & "&txtBpCd=" & Trim(.txtBpCd.value)                    '��ȸ ���� ����Ÿ 
            strVal = strVal & "&txtissue_dt_fg_cd=" & Trim(.cboissue_dt_fg.value)
            strVal = strVal & "&txtissue_dt_kind_cd=" & Trim(.cboissue_dt_kind.value)
        
            strVal = strVal & "&lgStrPrevVatKey=" & lgStrPrevVatKey
            strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
            strVal = strVal & "&lgPageNo=" & lgPageNo

        Call RunMyBizASP(MyBizASP, strVal)        '��: �����Ͻ� ASP �� ���� 
                        
    End With
    
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()                                                        '��: ��ȸ ������ ������� 
    
    Call SetSpreadLock()'(-1, "Query")
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE    '��: Indicates that current mode is Update mode
    
    ' ���� Page�� From Element���� ����ڰ� �Է��� ���� ���ϰ� �ϰų� �ʼ��Է»����� ǥ���Ѵ�.    
    Call ggoOper.LockField(Document, "Q")    '��: This function lock the suitable field
    Call FncSetToolBar("Query")
    
    'SetGridFocus        
    Set gActiveElement = document.activeElement 
    
End Function

'========================================================================================
' Function Name : DbSave()
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow
    Dim lGrpCnt
    Dim strVal,strDel
    Dim strYear,strMonth,strDay
    Dim iColSep
    
    'Call LayerShowHide(1)

    DbSave = False                '��: Processing is NG
    'On Error Resume Next        '��: Protect system from crashing
	
    With frm1
        .txtMode.value = parent.UID_M0002
        .txtUpdtUserId.value = parent.gUsrID
        
        '-----------------------
        'Data manipulate area
        '-----------------------
        lGrpCnt = 1
        strVal = ""
        strDel = ""
        iColSep = Parent.gColSep
    
        '-----------------------
        'Data manipulate area
        '-----------------------
        
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            
            Select Case .vspdData.Text
            
                Case ggoSpread.UpdateFlag                                                '��: ���� 
                    strVal = strVal & "U" & iColSep & lRow & iColSep                    '��: U=Update
                    .vspdData.Col = C_VAT_NO
                    strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_ISSUED_DT
                    strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_BP_CD
                    strVal = strVal & Trim(UCase(.vspdData.Text)) & iColSep
                    .vspdData.Col = C_VAT_TYPE
                    strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_CARD_NO
                    strVal = strVal & Trim(UCase(.vspdData.Text)) & iColSep
                    .vspdData.Col = C_REPORT_BIZ_AREA_CD
                    strVal = strVal & Trim(UCase(.vspdData.Text)) & iColSep
                    .vspdData.Col = C_BIZ_AREA_CD
                    strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_NET_LOC_AMT
                    strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
                    .vspdData.Col = C_VAT_LOC_AMT
                    strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep

					.vspdData.Col = C_issue_dt_kind_cd	: strVal = strVal & Trim(.vspdData.Text) & iColSep
                    .vspdData.Col = C_issue_dt_fg_cd	: strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep  
                                        
                                   
                    
                    lGrpCnt = lGrpCnt + 1
            End Select
                        
        Next
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value =  strVal

         Call ExecMyBizASP(frm1, BIZ_PGM_ID)        '��: �����Ͻ� ASP �� ���� 
    
    End With

    DbSave = True                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()                                                    '��: ���� ������ ���� ���� 
    
    Call InitVariables
    'frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

  
    Call MainQuery()

End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete()
    On Error Resume Next
End Function

'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################

Function FncSetToolBar(Cond)
    Select Case UCase(Cond)
    Case "QUERY"
        Call SetToolbar("1100100100111111")
    End Select
End Function

'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()    
   
    Frm1.vspdData.Row = 1
    Frm1.vspdData.Col = 1
    Frm1.vspdData.Action = 1
        
End Sub    


Function AcctApply()
	
	Dim lRow
	Dim str_fg,str_issue_kind
	
	str_fg = frm1.cboissue_dt_fg2.value 
	str_issue_kind = frm1.cboissue_dt_kind2.value 
     ggoSpread.Source = frm1.vspdData

	With Frm1.vspdData
       For lRow = 1 To .MaxRows
			.Row = lRow
			if str_fg<>"" then
				.Col = C_issue_dt_fg_cd : 				.Text = str_fg
				.Col = C_issue_dt_fg_nm : 				.Text = str_fg
				  ggoSpread.UpdateRow lRow
			end if	
			if str_issue_kind<>"" then	
				.Col = C_issue_dt_kind_cd : 				.Text = str_issue_kind
				call vspdData_Change(C_issue_dt_kind_cd,lRow)
				
		     end if
		Next
		
	End With

End Function
 

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->    
</HEAD>
<!-- '#########################################################################################################
'                           6. Tag�� 
'#########################################################################################################  -->

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
                                <td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ΰ�������</font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* ALIGN=RIGHT>&nbsp;</TD>
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
                                    <TD CLASS="TD5" NOWRAP>��������</TD>
                                    <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtIssuedDtFr" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="���۹�������" id=txtIssuedDtFr></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
                                                           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtIssuedDtTo" CLASS=FPDTYYYYMMDD tag="12" Title="FPDATETIME" ALT="�����������" id=txtIssuedDtTo></OBJECT>');</SCRIPT>
                                                           
                                    </TD>
                                    <TD CLASS="TD5" NOWRAP>���ݽŰ�����</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtReportBizArea" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="���ݽŰ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReportBizArea" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenReportBizArea()">&nbsp;
                                                           <INPUT TYPE=TEXT NAME="txtReportBizAreaNm" SIZE=20 tag="14" ALT="���ݽŰ�����"></TD>
                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>���ⱸ��</TD>
                                    <TD CLASS="TD6" NOWRAP><SELECT NAME="cboIoFg" ALT="���ⱸ��" tag="11" STYLE="WIDTH: 100px"  ><OPTION VALUE=""></OPTION></SELECT></TD>                                        
                                                                        </TD>
                                    <TD CLASS="TD5" NOWRAP>��꼭����</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="��꼭����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenVatType()">&nbsp;
                                                           <INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="14" ALT="��꼭����"></TD>

                                    </TD>
                                </TR>
                                <TR>
                                    <TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
                                    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 1)">&nbsp;
                                                           <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14" ALT="�ŷ�ó"></TD>                                                                        
                                    </TD>
									<TD CLASS="TD5" NOWRAP>��꼭��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatNo" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="��꼭��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript: Call OpenVatNoInfo(frm1.txtVatNo.value,'VatNo')"></TD>
                                </TR>
                                 <TR>
                                  
									<TD CLASS="TD5" NOWRAP>���ڼ��ݰ�꼭���࿩��</TD>
									<TD CLASS="TD6" NOWRAP> <SELECT NAME="cboissue_dt_fg" ALT="���ڼ��ݰ�꼭���࿩��" tag="11" STYLE="WIDTH: 100px"  ><OPTION VALUE=""></OPTION></SELECT></TD>
									  <TD CLASS="TD5" NOWRAP>���ڼ��ݰ�꼭����</TD>
                                    <TD CLASS="TD6" NOWRAP>
                                    <SELECT NAME="cboissue_dt_kind" ALT="���ڼ��ݰ�꼭����" tag="11" STYLE="WIDTH: 170px"  ><OPTION VALUE=""></OPTION></SELECT>
                                    </TD>
                                </TR>
                                
                                
                            </TABLE>
                        </FIELDSET>
                    </TD>
                </TR>
                
                <TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>���ڼ��ݰ�꼭���࿩��</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboissue_dt_fg2" ALT="���ڼ��ݰ�꼭���࿩��" tag="11" STYLE="WIDTH: 100px"  ><OPTION VALUE=""></OPTION></SELECT>
									 </TD>
									 
									<TD CLASS=TD5 NOWRAP>���ڼ��ݰ�꼭����</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboissue_dt_kind2" ALT="���ڼ��ݰ�꼭����" tag="11" STYLE="WIDTH: 170px"  ><OPTION VALUE=""></OPTION></SELECT>
														 <BUTTON NAME="btnApply" style="height:20px" CLASS="CLSSBTN" ONCLICK="vbscript:AcctApply()">����</BUTTON> </TD>
								
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
                                <TD HEIGHT="100%">
                                    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="2" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
            <IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

