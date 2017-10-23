<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m5131ma1
'*  4. Program Name         : �����ϰ�ó�� 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2000/05/06
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Lee Eun Hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/05/08,2000/05/11
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc ����   **************************************** -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ====================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!-- '==========================================  1.1.2 ���� Include   ==================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit											'��: indicates that All variables must be declared in advance

<%
Const BIZ_PGM_ID 		= "M5131mb1.asp"												'��: �����Ͻ� ���� ASP�� 
%>
Const BIZ_PGM_ID 		= "M5131mb1.asp"

Dim C_Select	
Dim C_PostFlag	
Dim C_IvNo		
Dim C_SpplCd	
Dim C_SpplNm	
Dim C_IvAmt		
Dim C_VatAmt	
Dim C_Currency	
Dim C_IvDt		
Dim C_GrpCd		
Dim C_GrpNm		
Dim C_BizAreaCd	
Dim C_BizAreaNm	
Dim C_GlType    
Dim C_GlNo		
Dim C_glref_pop 

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)


Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntGrpCount              ' Group View Size�� ������ ���� 
Dim lgIntFlgMode               ' Variable is for Operation Status

Dim lgStrPrevKey
Dim lgLngCurRows
Dim lgSortKey

Dim IsOpenPop          
Dim lblnWinEvent

'================================  initSpreadPosVariables()  ============================================
Sub initSpreadPosVariables()  
	C_Select	= 1
	C_PostFlag	= 2
	C_IvNo		= 3      '���Թ�ȣ 
	C_SpplCd	= 4      '����ó 
	C_SpplNm	= 5      '����ó�� 
	C_IvAmt		= 6      '���Աݾ� 
	C_VatAmt	= 7      'VAT�ݾ� 
	C_Currency	= 8      'ȭ�� 
	C_IvDt		= 9     '���Ե���� 
	C_GrpCd		= 10     '���ű׷� 
	C_GrpNm		= 11     '���ű׷�� 
	C_BizAreaCd	= 12     '���ݽŰ����� 
	C_BizAreaNm	= 13     '���ݽŰ������ 
	C_GlType    = 14     '��ǥ type
	C_GlNo		= 15     '��ǥ��ȣ 
	C_glref_pop = 16     '��ǥ��ȸ �˾� 
End Sub
'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
	frm1.vspdData.MaxRows = 0
	lgSortKey         = 1                                       '��: initializes sort direction
    
End Sub
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()

	frm1.rdoFlg(0).checked = true       '������ 
	frm1.rdoApFlg(1).checked = true     'Ȯ������ 

	frm1.txtFrIvDt.Text = StartDate
	frm1.txtToIvDt.Text = EndDate

	Call SetToolBar("1110000000001111")
	frm1.txtIvTypeCd.focus 
	Set gActiveElement = document.activeElement
	
End Sub

'========================================  LoadInfTB19029()  ======================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
    <% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>

End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables() 
	
	ggoSpread.Source = frm1.vspdData
	    
    ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread  
	With frm1.vspdData

    .ReDraw = false	
    .MaxCols = C_glref_pop+1
    .MaxRows = 0
  
    Call GetSpreadColumnPos("A")

    ggoSpread.SSSetCheck	C_Select	, "����", 10,,,true
	ggoSpread.SSSetEdit		C_PostFlag	, "Ȯ������", 10,,,,2
	ggoSpread.SSSetEdit 	C_IvNo		, "���Թ�ȣ", 20
    ggoSpread.SSSetEdit 	C_SpplCd	, "����ó", 10
    ggoSpread.SSSetEdit 	C_SpplNm	, "����ó��", 20
    '����(2003.03.19)
    ggoSpread.SSSetFloat    C_IvAmt		, "���Աݾ�"	, 15    ,"A"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    ggoSpread.SSSetFloat    C_VatAmt	, "VAT�ݾ�"		, 15    ,"A"   ,ggStrIntegeralPart ,ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
    ggoSpread.SSSetEdit 	C_Currency	, "ȭ��",10
    ggoSpread.SSSetDate		C_IvDt		, "���Ե����", 15,2,parent.gDateFormat
    ggoSpread.SSSetEdit 	C_GrpCd		, "���ű׷�",10
    ggoSpread.SSSetEdit 	C_GrpNm		, "���ű׷��",20
    ggoSpread.SSSetEdit 	C_BizAreaCd	, "���ݽŰ�����",10
    ggoSpread.SSSetEdit 	C_BizAreaNm	, "���ݽŰ������",20  
    ggoSpread.SSSetEdit 	C_GlType	, "C_GlType", 10
    ggoSpread.SSSetEdit 	C_GlNo		, "��ǥ��ȣ",20
    ggoSpread.SSSetButton 	C_glref_pop    
    
    Call ggoSpread.MakePairsColumn(C_GlNo,C_glref_pop)
    Call ggoSpread.SSSetColHidden(C_GlType,C_GlType,True)
    Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
    Call SetSpreadLock 
    
	.ReDraw = true
	
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
		ggoSpread.Source = frm1.vspdData
		
	    ggoSpread.SpreadUnLock		C_Select	,	-1,	C_Select,		-1
		ggoSpread.SpreadLock 		C_PostFlag	,	-1,	C_PostFlag,		-1
		ggoSpread.SpreadLock 		C_IvNo		,	-1,	C_IvNo,			-1
		ggoSpread.SpreadLock 		C_SpplCd	,	-1, C_SpplCd,		-1
		ggoSpread.SpreadLock 		C_SpplNm	,	-1, C_SpplNm,		-1
		ggoSpread.SpreadLock		C_IvAmt		,	-1, C_IvAmt,		-1
		ggoSpread.SpreadLock		C_VatAmt	,	-1, C_VatAmt,		-1 
		ggoSpread.SpreadLock 		C_Currency	,	-1, C_Currency,		-1
		ggoSpread.SpreadLock		C_IvDt		,	-1, C_IvDt,			-1
		ggoSpread.SpreadLock 		C_GrpCd		,	-1, C_GrpCd,		-1
		ggoSpread.SpreadLock 		C_GrpNm		,	-1, C_GrpNm,		-1
		ggoSpread.SpreadLock 		C_BizAreaCd	,	-1, C_BizAreaCd,	-1
		ggoSpread.SpreadLock 		C_BizAreaNm	,	-1, C_BizAreaNm,	-1
		ggoSpread.SpreadLock 		C_GlType	,	-1, C_GlType,		-1
		ggoSpread.SpreadLock 		C_GlNo		,	-1, C_GlNo,			-1
		ggoSpread.SpreadLock 		C_glref_pop ,	-1,	C_glref_pop,	-1   
		ggoSpread.SSSetProtected	C_glref_pop + 1,  -1	
End Sub

'===================================  GetSpreadColumnPos()  =====================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Select	= iCurColumnPos(1)
			C_PostFlag	= iCurColumnPos(2)
			C_IvNo		= iCurColumnPos(3)
			C_SpplCd	= iCurColumnPos(4)
			C_SpplNm	= iCurColumnPos(5)
			C_IvAmt		= iCurColumnPos(6)
			C_VatAmt	= iCurColumnPos(7)
			C_Currency	= iCurColumnPos(8)
			C_IvDt		= iCurColumnPos(9)
			C_GrpCd		= iCurColumnPos(10)
			C_GrpNm		= iCurColumnPos(11)
			C_BizAreaCd	= iCurColumnPos(12)
			C_BizAreaNm	= iCurColumnPos(13)
			C_GlType    = iCurColumnPos(14)
			C_GlNo		= iCurColumnPos(15)
			C_glref_pop = iCurColumnPos(16)

    End Select    
End Sub
'------------------------------------------  OpenGrp()  -------------------------------------------------
Function OpenGrp()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	If lblnWinEvent = True Or UCase(frm1.txtGrpCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	lblnWinEvent = True
	
	arrHeader(0) = "���ű׷�"									' Header��(0)
    arrHeader(1) = "���ű׷��"									' Header��(1)
    
    arrField(0) = "PUR_GRP"											' Field��(0)
    arrField(1) = "PUR_GRP_NM"										' Field��(1)
    
	arrParam(0) = "���ű׷�"									' �˾� ��Ī 
	arrParam(1) = "B_PUR_GRP"										' TABLE ��Ī 
	arrParam(2) = FilterVar(Trim(frm1.txtGrpCd.Value), "", "SNM")	' Code Condition
	arrParam(4) = ""
	arrParam(5) = "���ű׷�"									' TextBox ��Ī 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtGrpCd.focus
		Exit Function
	Else
		frm1.txtGrpCd.Value = arrRet(0)
		frm1.txtGrpNm.Value = arrRet(1)
		frm1.txtGrpCd.focus	
		Set gActiveElement = document.activeElement
    end if
    	
End Function

'------------------------------------------  OpenSppl()  -------------------------------------------------
Function OpenSppl()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	If lblnWinEvent = True Or UCase(frm1.txtSpplCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	lblnWinEvent = True
	
	arrHeader(0) = "����ó"											' Header��(0)
    arrHeader(1) = "����ó��"										' Header��(1)
    'arrHeader(2) = "����ڵ�Ϲ�ȣ"								' Header��(2)
    
    arrField(0) = "BP_Cd"												' Field��(0)%>
    arrField(1) = "BP_Nm"												' Field��(1)%>
    
	arrParam(0) = "����ó"											' �˾� ��Ī %>
	arrParam(1) = "B_BIZ_PARTNER"										' TABLE ��Ī %>
	arrParam(2) = FilterVar(Trim(frm1.txtSpplCd.Value), "", "SNM")							' Code Condition
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "										' Where Condition
	arrParam(5) = "����ó"											' TextBox ��Ī 
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtSpplCd.focus
		Exit Function
	Else
		frm1.txtSpplCd.Value = arrRet(0)
		frm1.txtSpplNm.Value = arrRet(1)
		frm1.txtSpplCd.focus	
		Set gActiveElement = document.activeElement
    end if
    	
End Function

'------------------------------------------  OpenIvType()  -------------------------------------------------
Function OpenIvType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	If lblnWinEvent = True Or UCase(frm1.txtIvTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	lblnWinEvent = True

	arrParam(0) = "��������"						' �˾� ��Ī 
	arrParam(1) = "M_IV_TYPE"							' TABLE ��Ī 
	arrParam(2) = FilterVar(Trim(frm1.txtIvTypeCd.Value), "", "SNM")			' Code Condition
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") 									' Where Condition
	arrParam(5) = "��������"						' TextBox ��Ī 
	
    arrField(0) = "IV_TYPE_CD"							' Field��(0)
    arrField(1) = "IV_TYPE_NM"							' Field��(1)
    
	arrHeader(0) = "��������"						' Header��(0)
    arrHeader(1) = "�������¸�"						' Header��(1)
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtIvTypeCd.focus
		Exit Function
	Else
		frm1.txtIvTypeCd.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
		frm1.txtIvTypeCd.focus	
		Call ChangeIvtype()
		Set gActiveElement = document.activeElement
    end if
    
End Function

'------------------------------------------  OpenBizArea()  -------------------------------------------------
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lblnWinEvent = True Or UCase(frm1.txtBizAreaCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lblnWinEvent = True

	arrParam(0) = "���ݽŰ�����"	
	arrParam(1) = "B_TAX_BIZ_AREA"
	arrParam(2) = FilterVar(Trim(frm1.txtBizAreaCd.Value), "", "SNM")
	arrParam(4) = ""
	arrParam(5) = "���ݽŰ�����"			
	
    arrField(0) = "TAX_BIZ_AREA_CD"
    arrField(1) = "TAX_BIZ_AREA_NM"
    
    arrHeader(0) = "���ݽŰ�����"
    arrHeader(1) = "���ݽŰ������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus	
		Exit Function
	Else
		frm1.txtBizAreaCd.Value = arrRet(0)
		frm1.txtBizAreaNm.Value = arrRet(1)
		frm1.txtBizAreaCd.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function
'------------------------------------------  OpenGLRef()  -------------------------------------------------
Function OpenGLRef()

	Dim strRet
	Dim arrParam(1)
	Dim iCalledAspName
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	
	frm1.vspdData.Col = C_GlNo                      '��ǥ��ȣ  
    arrParam(0) = Trim(frm1.vspdData.Text)
    frm1.vspdData.Col = C_IvNo                      '���Թ�ȣ 
	arrParam(1) = Trim(frm1.vspdData.Text)              
	
   frm1.vspdData.Col = C_GlType                      '��ǥ��ȣ type 
   
   If Trim(frm1.vspdData.Text) = "A" Then               'ȸ����ǥ�˾� 
		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif Trim(frm1.vspdData.Text) = "T" Then          '������ǥ�˾� 
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
			IsOpenPop = False
			Exit Function
		End If
	   strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

    Elseif Trim(frm1.vspdData.Text) = "B" Then
     	Call DisplayMsgBox("205154","X" , "X","X")   '���� ��ǥ�� �������� �ʾҽ��ϴ�. 
    End if

	lblnWinEvent = False
	
End Function

'======================================   Getglno()  =====================================
Sub Getglno()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
    Dim strwhere,strrefno
    Dim strglno
    
    frm1.vspdData.Row = frm1.vspdData.ActiveRow
    
    frm1.vspdData.Col = C_IvNo           '���Թ�ȣ 
    strrefno = Trim(frm1.vspdData.Text)
    
    Err.Clear
    
    strwhere = " ref_no =  " & FilterVar(strrefno , "''", "S") & ""
    Call CommonQueryRs(" gl_no ", " a_gl ",strwhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    if Trim(lgF0) = "" then
        Err.Clear
        Call CommonQueryRs(" temp_gl_no ", " a_temp_gl ",strwhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        if Trim(lgF0) = "" then
            frm1.vspdData.Col = C_GlType  
            frm1.vspdData.Text = "B"
        else
            frm1.vspdData.Col = C_GlType  
            frm1.vspdData.Text = "T"
        end if
        
    else
        frm1.vspdData.Col = C_GlType  
        frm1.vspdData.Text = "A"
    end if

End Sub
'===========================  SetSpreadFloatLocal()   ==============================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , _
                    ByVal dColWidth , ByVal HAlign , _
                    ByVal iFlag )
	        
   Select Case iFlag
        Case 2                                                              '�ݾ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign
        Case 3                                                              '���� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
        Case 4                                                              '�ܰ� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
        Case 5                                                              'ȯ�� 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec, HAlign,,"P"
    End Select
         
End Sub

Sub Changeflg()
	'lgBlnFlgChgValue = True
End Sub

'------------------------------------------  ChangeIvtype()  ---------------------------------------------
Sub ChangeIvtype()

	On Error Resume Next
	Err.Clear
	
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strIvTypeNm , strImportFlg
	Call CommonQueryRs(" IV_TYPE_NM, IMPORT_FLG ", " M_IV_TYPE ", " IV_TYPE_CD =  " & FilterVar(frm1.txtIvTypeCd.Value, "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    if isnull(lgF0) then
		frm1.txtIvTypeNm.Value = ""
		Err.Clear 
		Exit Sub
    end IF
    
    strIvTypeNm		= Split(lgF0, Chr(11))
	strImportFlg	= Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		frm1.txtIvTypeNm.Value = ""
		Err.Clear 
		Exit Sub
	Else 
		frm1.txtIvTypeNm.Value = strIvTypeNm(0)
		frm1.hdnImportFlg.Value = strImportFlg(0)
	End If
		
End sub
'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
        
End Sub
'======================================  vspdData_MouseDown()  =================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'======================================  FncSplitColumn()  =================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
    
End Function
'======================================  vspdData_Click()  =================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    
    Call SetPopupMenuItemInf("0000111111")
    
    gMouseClickStatus = "SPC"  
     
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
		
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If   
    
    
End Sub
'======================================  vspdData_ColWidthChange()  =================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'======================================  vspdData_ScriptDragDropBlock()  =================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    If Col <= C_PostFlag Or NewCol <= C_PostFlag Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'======================================  OCX_EVENT()  =================================
Sub txtFrIvDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrIvDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtFrIvDt.Focus
	End if
End Sub

Sub txtToIvDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToIvDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtToIvDt.Focus
	End if
End Sub

Sub txtApDt_DblClick(Button)
	if Button = 1 then
		frm1.txtApDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtApDt.Focus
	End if
End Sub

Sub txtFrIvDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToIvDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'========================================================================================================
'=	Event Name : rdoflg1_onpropertychange															=
'=	Event Desc :   ���Ե���� Ŭ���� lock																	=
'========================================================================================================
Sub rdoflg1_onpropertychange()	
	if frm1.rdoflg(0).checked = true then
		ggoOper.SetReqAttr	frm1.txtApDt, "Q"
	End if
	
End Sub
'=======================================================================================================
'=	Event Name : rdoflg2_onpropertychange															=
'=	Event Desc :������ Ŭ���� lock�� Ǯ���ش�														=
'========================================================================================================
Sub rdoflg2_onpropertychange()
	
	if frm1.rdoflg(1).checked = true then
		ggoOper.SetReqAttr	frm1.txtApDt, "N"
	End if
	
End Sub
'======================================  vspdData_Change()  =================================
Sub vspdData_Change(ByVal Col , ByVal Row )
		
End Sub
'======================================  vspdData_DblClick()  =================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
	If lgIntFlgMode = parent.OPMD_CMODE Then Exit Sub
	
	 If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Sub
     End If
	Frm1.vspdData.ReDraw = False
	If Col = C_Select And Row > 0 Then
	    Select Case ButtonDown
	    Case 0

			ggoSpread.Source = frm1.vspdData
			ggoSpread.EditUndo Row
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_Currency,C_IvAmt,"A" ,"I","X","X")         
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_Currency,C_VatAmt,"A" ,"I","X","X")         
			lgBlnFlgChgValue = False

	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True

	        frm1.vspdData.Col = C_PostFlag
	        if Trim(frm1.vspdData.Text) = "Y" then
	           frm1.vspdData.Text = "N"
	        else
	           frm1.vspdData.Text = "Y"
	        end if
	    End Select
    elseIF Col = C_glref_pop then
       Call Getglno()
       Call OpenGLRef()
    End If

	Frm1.vspdData.ReDraw = True
End Sub
'======================================  vspdData_TopLeftChange()  =================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	
	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
	    If lgStrPrevKey <> "" Then
		    If CheckRunningBizProcess = True Then
			    Exit Sub
		    End If	
			
		    Call DisableToolBar(parent.TBC_QUERY)
		    If DBQuery = False Then
			    Call RestoreToolBar()
			    Exit Sub
		    End If
	    End if
	End if    
End Sub


'======================================  FncQuery()  =================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = true or ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables
    															'��: Initializes local global variables
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    with frm1
        If CompareDateByFormat(.txtFrIvDt.text,.txtToIvDt.text,.txtFrIvDt.Alt,.txtToIvDt.Alt, _
                   "970025",.txtFrIvDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrIvDt.text) <> "" And Trim(.txtToIvDt.text) <> "" Then
           Call DisplayMsgBox("17a003","X","���Ե����","X")	      
           Exit Function
        End if                  
	End with
        
   If DbQuery = False Then Exit Function
       
    FncQuery = True																'��: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'======================================  FncNew()  =================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
    ggoSpread.Source = frm1.vspdData
    
    Err.Clear                                                               '��: Protect system from crashing
    On Error Resume Next                                                    '��: Protect system from crashing
    
    If lgBlnFlgChgValue = True or ggoSpread.SSCheckChange = True  Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")                                          '��: Lock  Suitable  Field
    Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
    
    FncNew = True                                                           '��: Processing is OK
	Set gActiveElement = document.activeElement
End Function

'======================================  FncDelete()  =================================
Function FncDelete() 
	Dim IntRetCD

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")
    If IntRetCD = vbNo Then Exit Function

    
    FncDelete = False                                                       '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    On Error Resume Next                                                    '��: Protect system from crashing
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then                                                '��: Delete db data
       Exit Function                                                        '��:
    End If

    Call ggoOper.ClearField(Document, "A")                                         '��: Clear Condition Field
    
    FncDelete = True                                                        '��: Processing is OK
    Set gActiveElement = document.activeElement
End Function

'======================================  FncSave()  =================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear    
    
	if CheckRunningBizProcess = true then
		exit function
	end if
    On Error Resume Next                                                    '��: Protect system from crashing
    	
    ggoSpread.Source = frm1.vspdData	
	
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
		'Call MsgBox("No data changed!!", vbInformation)
	    Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                  '��: Check contents area
       Exit Function
    End If
	
    If DbSave  = False Then Exit Function                                   '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
    Set gActiveElement = document.activeElement
End Function



'======================================  FncCancel()  =================================
Function FncCancel() 
	frm1.vspdData.Redraw = False
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                                 '��: Protect system from crashing
    
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_currency,C_vat_doc_amt,"A" ,"I","X","X")         
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_currency,C_pay_doc_amt,"A" ,"I","X","X")         
    
    frm1.vspdData.Redraw = True
    Set gActiveElement = document.activeElement
End Function

'======================================  FncPrint()  =================================
Function FncPrint()
	   Call parent.FncPrint()
	   Set gActiveElement = document.activeElement
End Function

'======================================  FncDeleteRow()  =================================
Function FncDeleteRow() 
    Dim lDelRows
    
    ggoSpread.Source = frm1.vspdData
    
    if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow

    lgBlnFlgChgValue = True
    
    End With
    Set gActiveElement = document.activeElement
End Function

'======================================  FncPrev()  =================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'======================================  FncNext()  =================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'======================================  FncExcel()  =================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
    Set gActiveElement = document.activeElement
End Function

'======================================  FncFind()  =================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
    Set gActiveElement = document.activeElement
End Function
'======================================  PopSaveSpreadColumnInf()  =================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'======================================  PopRestoreSpreadColumnInf()  =================================
Sub PopRestoreSpreadColumnInf()
	Dim index
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	
	With frm1
	
	.vspdData.ReDraw = False
	
    For index = 1 to .vspdData.MaxRows 
        .vspdData.Col = C_PostFlag   'Ȯ������ 
		.vspdData.Row = index
		
		if Trim(.vspdData.Text) = "Y" then
            ggoSpread.spreadUnlock 		C_glref_pop, index,C_glref_pop,index
        else
            ggoSpread.SpreadLock 		C_glref_pop, index,C_glref_pop,index
        end if
	next
	.vspdData.ReDraw = True
	End With
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_Currency,C_IvAmt,"A" ,"I","X","X")         
    Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,-1,-1,C_Currency,C_VatAmt,"A" ,"I","X","X")         
	
End Sub
'======================================  FncExit()  =================================
Function FncExit()
	
	Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")              '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    Set gActiveElement = document.activeElement
End Function

'======================================  DbQuery()  =================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	
	Dim strVal
    
    if LayerShowHide(1) = false then
		exit function
	end if
    
    With frm1
    
	If lgIntFlgMode = parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    strVal = strVal & "&txtFrIvDt=" & .hdnFrDt.Value
	    strVal = strVal & "&txtToIvDt=" & .hdnToDt.Value
	    strVal = strVal & "&txtIvType=" & .hdnIvType.Value
	    strVal = strVal & "&txtSppl=" & .hdnSppl.Value
	    strVal = strVal & "&txtBizArea=" & .hdnBizArea.Value
	    strVal = strVal & "&txtGrp=" & .hdnGrp.Value
	    strVal = strVal & "&txtImportFlg=" & .hdnImportFlg.Value
	    strVal = strVal & "&txtApPost=" & .hdnApFlg.value
	    
	else
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    strVal = strVal & "&txtFrIvDt=" & Trim(.txtFrIvDt.Text)
	    strVal = strVal & "&txtToIvDt=" & Trim(.txtToIvDt.Text)
	    strVal = strVal & "&txtIvType=" & .txtIvTypeCd.Value
	    strVal = strVal & "&txtSppl=" & FilterVar(Trim(.txtSpplCd.Value), "", "SNM")
	    strVal = strVal & "&txtBizArea=" & FilterVar(Trim(.txtBizAreaCd.Value), "", "SNM")
	    strVal = strVal & "&txtGrp=" & FilterVar(Trim(.txtGrpCd.Value), "", "SNM")
	    strVal = strVal & "&txtImportFlg=" & .hdnImportFlg.Value
	    if .rdoApFlg(0).Checked = true then
	    	strVal = strVal & "&txtApPost=" & "Y"
	    else
	    	strVal = strVal & "&txtApPost=" & "N"
	    End if
	    
	end if
	
    Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True

End Function
'======================================  DbQueryOk()  =================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	Dim index

    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
	Call SetToolBar("11101000000111")

    With frm1
	.vspdData.ReDraw = False
	
    For index = 1 to .vspdData.MaxRows 
        .vspdData.Col = C_PostFlag   'Ȯ������ 
		.vspdData.Row = index
		
		if Trim(.vspdData.Text) = "Y" then
            ggoSpread.spreadUnlock 		C_glref_pop, index,C_glref_pop,index
        else
            ggoSpread.SpreadLock 		C_glref_pop, index,C_glref_pop,index
        end if
	next
	.vspdData.ReDraw = True
	End With

	Call RemovedivTextArea

End Function
'======================================  DbSave()  =================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim ColSep, RowSep
	
	Dim iTmpCUBuffer         '������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount    '������ ���� Position
	Dim iTmpCUBufferMaxCount '������ ���� Chunk Size
	Dim strCUTotalvalLen
	
	Dim objTEXTAREA '������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
	
    DbSave = False                                                          '��: Processing is NG
    
    ColSep = Parent.gColSep															
	RowSep = Parent.gRowSep
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '�ѹ��� ������ ������ ũ�� ����[����,�ű�]
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '�ʱ� ������ ����[����,�ű�]

	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0

    On Error Resume Next                                                   '��: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002
    
		lGrpCnt = 1
		strVal = ""
	
		If frm1.rdoFlg(0).checked = True then
			frm1.hdnApDateFlg.value = "IV"
		Else
			frm1.hdnApDateFlg.value = ""
		End If
		'����(2003.06.09)_____________________
		If frm1.rdoApFlg(0).checked = True Then
			frm1.hdnApFlg.value = "N"
		Else
			frm1.hdnApFlg.value = "Y"
		End If '-----------------------------
				
	    For lRow = 1 To .vspdData.MaxRows
	        .vspdData.Row = lRow
	        .vspdData.Col = C_Select

	        If .vspdData.Text = 1 Then
	            '���ȭ�� �����(2003.06.09)
				.vspdData.Col = C_IvNo
			    strVal = strVal & Trim(.vspdData.Text) & ColSep					'0
				'strVal = strVal & Trim(frm1.hdnApDateFlg.value) & ColSep		'1
				if frm1.rdoFlg(0).checked = True then
					.vspdData.Col = C_IvDt
					strVal = strVal & Trim(.vspdData.Text) & ColSep				'2
				else
					strVal = strVal & Trim(frm1.txtApDt.Text) & ColSep			'2
				End if
				'strVal = strVal & Trim(frm1.hdnImportFlg.value) & ColSep		'3
				'if frm1.rdoApFlg(0).checked = True then
				'	strVal = strVal & "N" & ColSep			'4
				'else
			'		strVal = strVal & "Y" & ColSep
			'	End if
		        strVal = strVal & lRow & RowSep				'5					
		        
		        lGrpCnt = lGrpCnt + 1
				'--------------------------------------------------------
				If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
                            
				   Set objTEXTAREA = document.createElement("TEXTAREA")                 '�������� �Ѱ��� form element�� �������� ������ �װ��� ����Ÿ ���� 
				   objTEXTAREA.name = "txtCUSpread"
				   objTEXTAREA.value = Join(iTmpCUBuffer,"")
				   divTextArea.appendChild(objTEXTAREA)     
 
				   iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' �ӽ� ���� ���� �ʱ�ȭ 
				   ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
				   iTmpCUBufferCount = -1
				   strCUTotalvalLen  = 0
				End If
       
				iTmpCUBufferCount = iTmpCUBufferCount + 1
      
				If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '������ ���� ����ġ�� ������ 
				   iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '���� ũ�� ���� 
				   ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
				End If   
				iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
				strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
				'--------------------------------------------------------
				strVal = ""
		     End If 
		     
	    Next
		
		If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
		   Set objTEXTAREA = document.createElement("TEXTAREA")
		   objTEXTAREA.name   = "txtCUSpread"
		   objTEXTAREA.value = Join(iTmpCUBuffer,"")
		   divTextArea.appendChild(objTEXTAREA)     
		End If 

		'.txtMaxRows.value = lGrpCnt-1
		'.txtSpread.value = strVal
	
		if lGrpCnt > 1 then
			if LayerShowHide(1) = false then
				exit function
			end if
			Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'��: �����Ͻ� ASP �� ���� 
		else
			Call DisplayMsgBox("900002","X","X","X")
		end if

	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function
'======================================  DbSaveOk()  =================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
	Call InitVariables	
    Call MainQuery()

End Function
'======================================  DbDelete()  =================================
Function DbDelete() 
End Function

'======================================  RemovedivTextArea()  =================================
Function RemovedivTextArea()
	Dim ii
	
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����ϰ�ó��</font></td>
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
									<TD CLASS="TD5" NOWRAP>���Ե����</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���Ե���� NAME="txtFrIvDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="12N1" Title="FPDATETIME"></OBJECT>');</SCRIPT> ~
														   <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=���Ե���� NAME="txtToIvDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="12N1" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtIvTypeCd" ALT="��������" MAXLENGTH=5 SIZE=10 tag="12NXXU" OnChange="VBScript:ChangeIvType()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">
														   <INPUT TYPE=TEXT NAME="txtIvTypeNm" ALT="��������" SIZE=20 tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSpplCd" ALT="����ó" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSupplier" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl()">
														   <INPUT TYPE=TEXT NAME="txtSpplNm" ALT="����ó" tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>���ݽŰ�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd" ALT="���ݽŰ�����" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayMethod" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizArea()">
														   <INPUT TYPE=TEXT NAME="txtBizAreaNm" ALT="���ݽŰ�����" SIZE=20 tag="14X" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGrpCd" ALT="���ű׷�" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGrp()">
														   <INPUT TYPE=TEXT NAME="txtGrpNm" ALT="���ű׷�" SIZE=20 tag="14X"></TD>
									<TD CLASS="TD5" NOWRAP>Ȯ������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio NAME="rdoApFlg" ALT="Ȯ������" id="rdoApFlg1" Value="Y" CLASS="RADIO" checked tag="11"><label for="rdoApFlg1">&nbsp;Yes</label>
														   <INPUT TYPE=radio NAME="rdoApFlg" ALT="Ȯ������" id="rdoApFlg2" Value="N" CLASS="RADIO" tag="11"><label for="rdoApFlg2">&nbsp;No&nbsp;</label></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>Ȯ��������</TD>
								<TD CLASS="TD6" NOWRAP>
									<Table Cellspacing=0 Cellpadding=0>
										<TR>
											<TD NOWRAP>
												<INPUT TYPE=radio NAME="rdoFlg" id=rdoflg1 ALT="���Ե����" CLASS="RADIO" checked tag="11" ONCLICK="vbscript:Changeflg()">
											</TD>
											<TD NOWRAP>
												<label for="rdoflg1">&nbsp;���Ե����&nbsp;</label>
											</TD>
											<TD NOWRAP>
												<INPUT TYPE=radio NAME="rdoFlg" id=rdoflg2 ALT="������" CLASS="RADIO" tag="11" ONCLICK="vbscript:Changeflg()">
											</TD>
											<TD NOWRAP>
												<label for="rdoflg2">&nbsp;������&nbsp;</label>&nbsp;
											</TD>
											<TD NOWRAP>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=������ NAME="txtApDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="24N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</Table> 	   
								<TD CLASS="TD6" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<P ID="divTextArea"></P>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSppl" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBizArea" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGrp" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnApFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnImportFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnApDateFlg" tag="14">
</FORM>



    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
    </DIV>

</BODY>
</HTML>

