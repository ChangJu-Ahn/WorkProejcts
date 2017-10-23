<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Change
'*  3. Program ID           : a7107ma1
'*  4. Program Name         : �����ڻ� ���⳻�� ��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : AS0031,  
'							  AS0039	
'							  +B19029LookupNumericFormatF	
'*  7. Modified date(First) : 2000/03/18
'*  8. Modified date(Last)  : 2001/02/23
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'=======================================================================================================
'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>


<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit    												'��: indicates that All variables must be declared in advance%>


'========================================================================================================
Const BIZ_PGM_ID  = "a7107mb1_ko441.asp"  
Const COM_BIZ_EXCHRATE_ID = "../../inc/GetExchRate.asp"			'ȯ������ �����Ͻ� ���� ASP�� 

Const C_SHEETMAXROWS = 30	
'****************************************************************************************************

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'@Grid_Column
Dim C_Seq
Dim C_RcptType								            'Spread Sheet �� Columns �ε��� 
Dim C_RcptTypePopup
Dim C_RcptTypeNm								            'Spread Sheet �� Columns �ε��� 
Dim C_Amt
Dim C_LocAmt
Dim C_BankAcct
Dim C_BankAcctPopup
Dim C_NoteNo	
Dim C_NoteNoPopup

'2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
Dim C_BuildAsstNo	
Dim C_BuildAsstNoPopup


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

Const gIsShowLocal = "Y"  
<%
Const gIsShowLocal = "Y"
Dim BaseDate
   BaseDate = GetSvrDate                                  'Company Formate Date   
%>

'@Global_Var
'Dim lgBlnFlgChgValue                                         'Variable is for Dirty flag
'Dim lgIntGrpCount                                            'Group View Size�� ������ ���� 
'Dim lgIntFlgMode                                             'Variable is for Operation Status

'Dim lgStrPrevKey                                             'Previous NextKey

Dim IsOpenPop						                        'Popup
'Dim lgSortKey
Dim gSelframeFlg                                            'Current Tab Page

Dim lgMasterQueryFg                                         ''�ڻ�Master�� query ���� 

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'========================================================================================================
Sub initSpreadPosVariables()
	C_Seq				= 1
	C_RcptType			= 2									            'Spread Sheet �� Columns �ε��� 
	C_RcptTypePopup		= 3
	C_RcptTypeNm		= 4								            'Spread Sheet �� Columns �ε��� 
	C_Amt				= 5
	C_LocAmt			= 6
	C_BankAcct			= 7
	C_BankAcctPopup		= 8
	C_NoteNo			= 9
	C_NoteNoPopup		= 10
	
    '2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
    C_BuildAsstNo	    = 11
    C_BuildAsstNoPopup  = 12
	
End Sub


'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '��: initializes Previous Key Index
    lgSortKey         = 1                                       '��: initializes sort direction
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================

Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 	
	frm1.txtChgDt.text    = UNIDateClientFormat("<%=BaseDate%>")
	frm1.txtDueDt.text    = UNIDateClientFormat("<%=BaseDate%>")
	frm1.txtIssuedDt.text    = UNIDateClientFormat("<%=BaseDate%>")

	frm1.txtDocCur.value	= parent.gCurrency
	
	If gIsShowLocal <> "N" Then
		frm1.txtXchRate.text	= "1"
	else
		frm1.txtXchRate.value	= "1"	
	end if
	
	Call ggoOper.FormatDate(frm1.txtChgDt, gDateFormat, 1)
	Call ggoOper.FormatDate(frm1.txtDueDt, gDateFormat, 1)	
	lgBlnFlgChgValue = False
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "MA") %>  ' check
End Sub

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

   ' ------ Developer Coding part (Start ) --------------------------------------------------------------
    Dim strCodeList
    Dim strNameList
    
'                       Select                 From        Where                Return value list
'    Call CommonQueryRs(" MINOR_CD,MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'A1006' AND MINOR_CD NOT IN ( 'NR', 'PP', 'CR', 'PR') ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)

	Call CommonQueryRs("A.MINOR_CD,A.MINOR_NM","B_MINOR A, B_CONFIGURATION B", _
						"(A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND (A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " ) AND A.MINOR_CD NOT IN ( " & FilterVar("NR", "''", "S") & " , " & FilterVar("PP", "''", "S") & " , " & FilterVar("CR", "''", "S") & " , " & FilterVar("PR", "''", "S") & " ) AND B.SEQ_NO = 4 ", _
	                         lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	'A1006

    strCodeList = Replace(lgF0, Chr(11), vbTab)
    strNameList = Replace(lgF1, Chr(11), vbTab)

    ggoSpread.SetCombo strCodeList, C_RcptType
    ggoSpread.SetCombo strNameList, C_RcptTypeNm
        
    '------ Developer Coding part (End )   --------------------------------------------------------------
End Sub


'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================

Sub InitData()
   '------ Developer Coding part (Start ) --------------------------------------------------------------    
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
			.Col = C_RcptType  :  intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
			.Col = C_RcptTypeNm  :  .Value = intindex					
		Next	
	End With
End Sub
' 
'
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	           
	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread  
		.ReDraw = false	

        '2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
       '.MaxCols = C_NoteNoPopup + 1                               '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		.MaxCols = C_BuildAsstNoPopup + 1                          '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		
		'.MaxRows = 0
		ggoSpread.Source = frm1.vspdData
		ggospread.ClearSpreadData		'Buffer Clear
		
		'Hidden Column ���� 
    	.Col = .MaxCols											'������Ʈ�� ��� Hidden Column
    	.ColHidden = True
    		
'    	.Col = C_RcptType
'    	.ColHidden = True

		Call GetSpreadColumnPos("A")
		
		'Call AppendNumberPlace("6","3","0")
		
		'ggoSpread.SSSetFloat  C_Seq,       "����",       5, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		
		
		ggoSpread.SSSetEdit	  C_Seq,        "����",   5, 2
		ggoSpread.SSSetEdit  C_RcptType,  "��������"       ,10,,,5,2
		ggoSpread.SSSetButton C_RcptTypePopup
		ggoSpread.SSSetEdit  C_RcptTypeNm,"����������"     ,16

		ggoSpread.SSSetFloat  C_Amt,       "�ݾ�",       19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_LocAmt,    "�ݾ�(�ڱ�)", 19, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		If gIsShowLocal = "N" Then
'			.Col = C_LocAmt
'    		.ColHidden = True
    		Call ggoSpread.SSSetColHidden(C_LocAmt,C_LocAmt,True)
		End If
		
		ggoSpread.SSSetEdit	  C_BankAcct,  "�������ڵ�",25,0,-1,30,2
		ggoSpread.SSSetButton C_BankAcctPopup
		ggoSpread.SSSetEdit   C_NoteNo,    "������ȣ",25,0,-1,30,2
		ggoSpread.SSSetButton C_NoteNoPopup


        '2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
		ggoSpread.SSSetEdit   C_BuildAsstNo,    "�Ǽ������ڻ��ȣ",25,0,-1,30,2
		ggoSpread.SSSetButton C_BuildAsstNoPopup        
        Call ggoSpread.MakePairsColumn(C_BuildAsstNo,C_BuildAsstNoPopup,"1")
		
		Call ggoSpread.MakePairsColumn(C_RcptType,C_RcptTypePopup,"1")
		Call ggoSpread.MakePairsColumn(C_BankAcct,C_BankAcctPopup,"1")
		Call ggoSpread.MakePairsColumn(C_NoteNo,C_NoteNoPopup,"1")
		
		'Call ggoSpread.SSSetColHidden(C_RcptType,C_RcptType,True)
		'Call InitComboBox
		.ReDraw = true
		
	End With
    Call SetSpreadLock 
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		
		ggoSpread.SpreadLock C_Seq,        -1, C_Seq
	    'ggoSpread.SpreadLock C_RcptType,   -1, C_RcptType
		''''ggoSpread.SpreadLock C_RcptTypeNm, -1, C_RcptTypeNm
		ggoSpread.SpreadLock C_NoteNoPopup+1,-1, C_NoteNoPopup+1
		.vspdData.ReDraw = True
	End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStarRow, ByVal pvEndRow)
   	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_Seq, pvStarRow, pvEndRow		
		ggoSpread.SSSetRequired	 C_RcptType, pvStarRow, pvEndRow
		ggoSpread.SSSetProtected C_RcptTypeNm, pvStarRow, pvEndRow
		ggoSpread.SSSetRequired	 C_Amt, pvStarRow, pvEndRow
		
		.vspdData.ReDraw = True
	End With
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		C_Seq				= iCurColumnPos(1)
		C_RcptType			= iCurColumnPos(2)
		C_RcptTypePopup		= iCurColumnPos(3)
		C_RcptTypeNm		= iCurColumnPos(4)
		C_Amt				= iCurColumnPos(5)
		C_LocAmt			= iCurColumnPos(6)
		C_BankAcct			= iCurColumnPos(7)
		C_BankAcctPopup		= iCurColumnPos(8)
		C_NoteNo			= iCurColumnPos(9)
		C_NoteNoPopup		= iCurColumnPos(10)

        '2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
        C_BuildAsstNo	    = iCurColumnPos(11)
        C_BuildAsstNoPopup  = iCurColumnPos(12)
		
	End Select
End Sub

'======================================================================================================
'   Function Name : OpenChgNoInfo()
'   Function Desc : 
'=======================================================================================================
Function OpenChgNoInfo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	'2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
	Dim strRequestMenuID
	strRequestMenuID = "a7107ma1"

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a7107ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7107ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	'2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
	'arrRet = window.showModalDialog(iCalledAspName & "?PID=" & gstrRequestMenuID , Array(window.parent,arrParam), _
	'		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	arrRet = window.showModalDialog(iCalledAspName & "?PID=" & strRequestMenuID , Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtChgNo.focus
		Exit Function
	Else
		Call SetChgNoInfo(arrRet)
	End If	


End Function

'======================================================================================================
'   Function Name : SetChgNoInfo(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetChgNoInfo(Byval arrRet)
	frm1.txtChgNo.focus
	frm1.txtChgNo.value  = arrRet(0)
End Function
'=======================================================================================================
'	Name : OpenNoteNo()
'	Description : Note No PopUp
'=======================================================================================================
Function OpenNoteNo(byVal strCode, byVal strCard)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim field_fg
	If IsOpenPop = True  Then Exit Function


	IF UCase(strCard) = "CP"	Then		
		arrParam(0) = "���ұ���ī�� �˾�"				        ' �˾� ��Ī 
		arrParam(1) = "f_note a,b_biz_partner b, b_bank c, b_card_co d"		' TABLE ��Ī 
		arrParam(2) = ""								' Code Condition
		arrParam(3) = ""								' Name Cindition			
		arrParam(4) = "a.note_sts = " & FilterVar("BG", "''", "S") & "  AND a.note_fg = " & FilterVar("CP", "''", "S") & "  and a.bp_cd = b.bp_cd  "			
		arrParam(4) = arrParam(4) & " and a.bank_cd *= c.bank_cd and a.card_co_cd *= d.card_co_cd "
		arrParam(5) = "����ī���ȣ"						' �����ʵ��� �� ��Ī 

		arrField(0) = "a.Note_no"					' Field��(0)
		arrField(1) = "F2" & parent.gColSep & "a.Note_amt"		' Field��(1)
		arrField(2) = "DD" & parent.gColSep & "a.Issue_dt"		' Field��(2)
		arrField(3) = "b.bp_nm"					' Field��(3)
		arrField(4) = "d.card_co_nm"    	    			' Field��(4)

		arrHeader(0) = "����ī���ȣ"				' Header��(0)
		arrHeader(1) = "�ݾ�"				' Header��(1)
		arrHeader(2) = "������"				' Header��(2)	    
		arrHeader(3) = "�ŷ�ó"				' Header��(3)
		arrHeader(4) = "ī���"				' Header��(4)

	ElseIf UCase(strCard) = "NP"	Then
		arrParam(0) = "���޾�����ȣ �˾�"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"				
		arrParam(2) = strCode
		arrParam(3) = ""
		
		arrParam(4) = "A.NOTE_STS = " & FilterVar("BG", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D3", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "���޾�����ȣ"			
		
	    arrField(0) = "A.NOTE_NO"		
	    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
	    arrField(2) = "C.BP_NM"	    
	    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
	    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
	    arrField(5) = "B.BANK_NM"	        
	    
	    arrHeader(0) = "���޾�����ȣ"
	    arrHeader(1) = "�����ݾ�"        		
		arrHeader(2) = "�ŷ�ó"        		        	
		arrHeader(3) = "������"        		        
		arrHeader(4) = "������"        		        
		arrHeader(5) = "����"       

	Else
		arrParam(0) = "�輭������ȣ �˾�"	
		arrParam(1) = "F_NOTE A,B_BANK B,B_BIZ_PARTNER C"				
		arrParam(2) = strCode
		arrParam(3) = ""
		
		arrParam(4) = "A.NOTE_STS = " & FilterVar("ED", "''", "S") & "  AND A.NOTE_FG = " & FilterVar("D1", "''", "S") & "  AND A.BP_CD = C.BP_CD AND A.BANK_CD = B.BANK_CD"				
		arrParam(5) = "�輭������ȣ"			
		
	    arrField(0) = "A.NOTE_NO"		
	    arrField(1) = "F2" & parent.gColSep & "A.NOTE_AMT"
	    arrField(2) = "C.BP_NM"	    
	    arrField(3) = "DD" & parent.gColSep & "A.ISSUE_DT"
	    arrField(4) = "DD" & parent.gColSep & "A.DUE_DT"	
	    arrField(5) = "B.BANK_NM"	        
	    
	    arrHeader(0) = "�輭������ȣ"
	    arrHeader(1) = "�����ݾ�"        		
		arrHeader(2) = "�ŷ�ó"        		        	
		arrHeader(3) = "������"        		        
		arrHeader(4) = "������"        		        
		arrHeader(5) = "����"       
	 		        
	End If	

	IsOpenPop = True	
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	Call GridSetFocus(2)
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopup(arrRet,2)
	End If	

End Function
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrParamAdo(3)

	If IsOpenPop = True Then Exit Function	
	
	Select Case iWhere
		Case 0
			Dim strAsstNo
			Dim IntRetCd

			strAsstNo  = Trim(frm1.txtAsstNo.value)
	
			if strAsstNo = "" then
				IntRetCD = DisplayMsgBox("117326","X","X","X")    '�ڻ��ȣ�� �Է��Ͻʽÿ�.
				IsOpenPop = False
				Exit Function
			end if
				
			arrParam(0) = "ȸ��μ� �˾�"	
			arrParam(1) = "B_ACCT_DEPT"
			arrParam(2) = Trim(frm1.txtDeptCd.value)
			arrParam(3) = ""
			arrParam(4) = " INTERNAL_CD IN (SELECT INTERNAL_CD FROM A_ASSET_INFORM_OF_DEPT WHERE ASST_NO =  " & FilterVar(frm1.txtAsstNo.value, "''", "S") & " )"
			arrParam(5) = "����μ�"			
			
			arrField(0) = "DEPT_CD"	
			arrField(1) = "DEPT_NM"
			arrField(2) = "ORG_CHANGE_ID "
			
			arrHeader(0) = "����μ��ڵ�"
			arrHeader(1) = "����μ���"
			arrHeader(2) = "��������ID"

		Case 1
			arrParam(0) = "�ŷ�ó �˾�"	
			arrParam(1) = "B_BIZ_PARTNER"				
			arrParam(2) = Trim(frm1.txtBpCd.Value)
			arrParam(3) = ""
			arrParam(4) = ""			
			arrParam(5) = "�ŷ�ó�ڵ�"			
	
			arrField(0) = "BP_CD"	
			arrField(1) = "BP_NM"
			arrField(2) = "BP_RGST_NO"
    
			arrHeader(0) = "�ŷ�ó �ڵ�"		
			arrHeader(1) = "�ŷ�ó ��"
			arrHeader(2) = "����ڵ�Ϲ�ȣ"
		Case 3	
			arrParam(0) = "�ŷ���ȭ �˾�"	
			arrParam(1) = "B_CURRENCY"				
			arrParam(2) = Trim(frm1.txtDocCur.Value)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "�ŷ���ȭ"
	
			arrField(0) = "CURRENCY"	
			arrField(1) = "CURRENCY_DESC"	
    
			arrHeader(0) = "�ŷ���ȭ"		
			arrHeader(1) = "�ŷ���ȭ��"
		Case 4    
			arrParam(0) = "�������ڵ� �˾�"	' �˾� ��Ī 
			arrParam(1) = "B_BANK A, F_DPST B"			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "A.BANK_CD = B.BANK_CD "		' Where Condition
			arrParam(5) = "�����ڵ�"				' �����ʵ��� �� ��Ī 
				
			arrField(0) = "A.BANK_NM"					' Field��(1)
			arrField(1) = "B.BANK_ACCT_NO"				' Field��(2)
   
			arrHeader(0) = "�����"						' Header��(1)
			arrHeader(1) = "�������ڵ�"

		Case 5
			arrHeader(0) = "�ΰ�������"						' Header��(0)
			arrHeader(1) = "�ΰ�����"						' Header��(1)
			arrHeader(2) = "�ΰ���Rate"
    
			arrField(0) = "B_Minor.MINOR_CD"							' Field��(0)
			arrField(1) = "B_Minor.MINOR_NM"							' Field��(1)
			arrField(2) = "F2" & parent.gColSep & "b_configuration.REFERENCE"    
'			arrField(2) = "b_configuration.REFERENCE"    
			arrParam(0) = "�ΰ�������"						' �˾� ��Ī 
			arrParam(1) = "B_Minor,b_configuration"				' TABLE ��Ī 
			arrParam(2) = Trim(StrCode)			' Code Condition
			'arrParam(3) = Trim(frm1.txtPayMethNM.Value)			' Name Cindition
			
			arrParam(4) = "B_Minor.Major_Cd=" & FilterVar("B9001", "''", "S") & "  and B_Minor.minor_cd = b_configuration.minor_cd and " & _
			              " b_configuration.SEQ_NO=1 AND b_configuration.major_cd = B_Minor.Major_Cd"	 
			arrParam(5) = "�ΰ�������"						' TextBox ��Ī			
		Case 6    
			arrParam(0) = "�Ա�����"								' �˾� ��Ī 
		 
			arrParam(1) = "B_MINOR A, B_CONFIGURATION B"
			arrParam(2) = Trim(frm1.vspdData.text)
			arrParam(3) = ""											' Name Condition
			arrParam(4) = "(A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND (A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " )" _
			   & " AND A.MINOR_CD NOT IN ( " & FilterVar("NR", "''", "S") & " , " & FilterVar("PP", "''", "S") & " , " & FilterVar("CR", "''", "S") & " , " & FilterVar("PR", "''", "S") & " ) AND B.SEQ_NO = 4 " ' Where Condition        
			arrParam(5) = "�Ա�����"								' TextBox ��Ī 
	 
			arrField(0) = "A.MINOR_CD"							' Field��(0)
			arrField(1) = "A.MINOR_NM"							' Field��(1)
			  
			arrHeader(0) = "�Ա�����"								' Header��(0)
			arrHeader(1) = "�Ա�������"								' Header��(1) 

	End Select
	
	IsOpenPop = True

	if iWhere = 1 then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=650px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")			
	end if

	IsOpenPop = False

	Call GridSetFocus(iWhere)
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If

End Function

'=======================================================================================================
Function GridsetFocus(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtDeptCd.focus
			Case 1
				.txtBpCd.focus
			Case 2
				Call SetActiveCell(.vspdData,C_NoteNo,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				.txtDocCur.focus
			Case 4
				Call SetActiveCell(.vspdData,C_BankAcct,.vspdData.ActiveRow ,"M","X","X")
			Case 5
				.txtVatType.focus
			Case 6
				Call SetActiveCell(.vspdData,C_Rcpttype,.vspdData.ActiveRow ,"M","X","X")
		END Select
	End With
End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
	Select Case iWhere
		Case 0
			.txtDeptCd.Value		= Trim(arrRet(0))
			.txtDeptNm.Value		= arrRet(1)
			.HORGCHANGEID.Value		= Trim(arrRet(2))
'			Call txtDeptCd_OnChange()
		Case 1
			.txtBpCd.Value    = Trim(arrRet(0))
			.txtBpNm.Value    = arrRet(1)
		Case 2
				.vspdData.Col	= C_NoteNo
				.vspdData.Text	= arrRet(0)

				.vspdData.Col	= C_Amt
				.vspdData.Text	= arrRet(1)

				.vspdData.Col	= C_LocAmt
				.vspdData.Text	= arrRet(1)
			    Call vspdData_Change(.vspdData.Col, .vspdData.Row)				 ' ������ dlf��ٰ� �˷��� 				
		Case 3
			.txtDocCur.value    = arrRet(0)
			If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' �ŷ���ȭ�ϰ� Company ��ȭ�� �ٸ��� ȯ���� 0���� ���� 
				If gIsShowLocal <> "N" Then
					.txtXchRate.text	= "0"                         ' ����Ʈ���� 1�� �� ������ ȯ���� �Էµ� ������ �Ǵ��Ͽ� 
				else
					.txtXchRate.value	= "0"
				end if							                          ' ȯ�������� ���� �ʰ� �Էµ� ������ ���. 
			Else 
				If gIsShowLocal <> "N" Then
					.txtXchRate.text	= "1"
				else
					.txtXchRate.value	= "1"
				end if
			End If
			call txtDocCur_OnChangeASP()
		Case 4
			.vspdData.Col = C_BankAcct
			.vspdData.Text = arrRet(1)
			Call vspdData_Change(frm1.vspdData.Col, frm1.vspdData.Row)				 ' ������ �о�ٰ� �˷��� 		
		Case 5
			.txtVatType.value = arrRet(0)
			.txtVatTypeNm.value = arrRet(1)
			If gIsShowLocal <> "N" Then
				.txtVatRate.text = arrRet(2)
			else
				.txtVatRate.value = arrRet(2)
			END IF
			Call txtVatType_OnChange
		Case 6
			.vspdData.Col = C_RcptType
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_RcptTypeNm
			.vspdData.Text = arrRet(1)
			Call vspdData_Change(C_RcptType, frm1.vspdData.Row)				 ' ������ �о�ٰ� �˷��� 		
	END Select
	End With
	lgBlnFlgChgValue = True
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
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :���� S: ���� T: ��ü 
	arrParam(5) = ""									'SUP :����ó PAYTO: ����ó SOL:�ֹ�ó PAYER :����ó INV:���ݰ�� 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		Call SetBpCd(arrRet)
		lgBlnFlgChgValue = True
	End If
		
End Function
'========================================================================================
Function SetBpCd(byval arrRet)
	frm1.txtBpCd.focus
	frm1.txtBpCd.Value    = arrRet(0)		
	frm1.txtBpNm.Value    = arrRet(1)		
	lgBlnFlgChgValue = True
End Function

'2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
Function OpenBuildAsstNoPopUp(Byval strCode, ByVal Row)
	Dim iCalledAspName
	Dim IntRetCD

	If GetSetupMod(Parent.gSetupMod, "p") <> "Y" Then
    	'Call DisplayMsgBox("169936","X", "X", "X")
		'Exit Function
	End if
				
	Dim arrRet
	Dim Param1, Param2, Param3, Param4, Param5, Param6 

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("A7102PA1")     '20080226::HANC
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"A7102PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, strCode), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    

	IsOpenPop = False
	
	If arrRet(0,0) = "" Then
		Call SetActiveCell(frm1.vspdData2,C_BuildAsstNo,frm1.vspdData2.ActiveRow,"M","X","X")
		Exit Function
	Else
		Call SetBuildAsstNoPopup(arrRet)
	End If
	
End Function

'2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
Function SetBuildAsstNoPopup(Byval arrRet)

	With frm1
		.vspdData.Col  = C_BuildAsstNo
		.vspdData.Text = arrRet(0,0)

		Call vspdData_Change(C_BuildAsstNo, frm1.vspdData.Row)				 ' ������ �о�ٰ� �˷��� 		
	End With
	lgBlnFlgChgValue = True
	
End Function


'------------------------------------------  OpenPoRef()  -------------------------------------------------
'	Name : OpenPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function OpenMasterRef()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If lgIntFlgMode = parent.OPMD_UMODE Then 
			Call DisplayMsgBox("200005", "X", "X", "X")
			Exit function
	End If	
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus 
		Exit Function
	Else
		Call SetPoRef(arrRet)
	End If	
	
		
End Function

 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(arrRet)
	
    Dim strVal
    
    lgMasterQueryFg = False
    
	
	Call ggoOper.ClearField(Document, "A")								
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear

    Call SetDefaultVal
    Call InitVariables
    
	frm1.txtAsstNo.value     = arrRet(0)
	frm1.txtAsstNm.value	 = arrRet(1)
	frm1.txtRegDt.text       = arrRet(2)	

	frm1.txtAcctDeptNm.value = arrRet(9)
	
	frm1.txtAcqQty.text     = arrRet(7)	
	frm1.txtInvQty.text     = arrRet(8)
	
	frm1.txtCur.value		 = arrRet(3)

	If gIsShowLocal <> "N" Then
		frm1.txtXchRt.text       = arrRet(4)	
		frm1.txtAcqLocAmt.text  = arrRet(6)
	else
		frm1.txtXchRt.value       = arrRet(4)	
		frm1.txtAcqLocAmt.value  = arrRet(6)	
	end if		
	
	frm1.txtAcqAmt.text     = arrRet(5)
	
	Call ggoOper.LockField(Document, "Q")
	
	frm1.txtDeptCd.focus 
	
	lgMasterQueryFg = True
	
	IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
	
	call txtCur_OnChange()
	
	if frm1.Rb_Rve.checked = true then call Radio2_onChange()
	
	'''lgBlnFlgChgValue = False
	End If
End Sub
'=======================================================================================================
'Description : ������ǥ �������� �˾� 
'=======================================================================================================
Function OpenPopupTempGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5130ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtTempGlNo.value)	'������ǥ��ȣ 
	arrParam(1) = ""							'Reference��ȣ 

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
'Description : ȸ����ǥ �������� �˾� 
'=======================================================================================================
Function OpenPopupGL()

	Dim arrRet
	Dim arrParam(8)	
	Dim iCalledAspName

	iCalledAspName = AskPRAspName("a5120ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	arrParam(0) = Trim(frm1.txtGlNo.value)	'ȸ����ǥ��ȣ 
	arrParam(1) = ""						'Reference��ȣ 

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

'===========================================================================
' Function Name : OpenReportAreaCd
' Function Desc : OpenReportAreaCd Reference Popup
'===========================================================================
Function OpenReportAreaCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)


	arrParam(0) = "�Ű������ �˾�"	
	arrParam(1) = "B_TAX_BIZ_AREA"				
	arrParam(2) = Trim(frm1.txtReportAreaCd.value)
	arrParam(3) = "" 
	arrParam(4) = ""
	arrParam(5) = "�Ű������"			
	
    arrField(0) = "TAX_BIZ_AREA_CD"	
    arrField(1) = "TAX_BIZ_AREA_NM"
    
    arrHeader(0) = "�Ű�������ڵ�"		
    arrHeader(1) = "�Ű�������"		

	IsOpenPop = True
        
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtReportAreaCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReportArea(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetReportArea()
'	Description : Bp Cd Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetReportArea(byval arrRet)
	frm1.txtReportAreaCd.Value		= arrRet(0)		
	frm1.txtReportAreaNm.Value		= arrRet(1)		
	lgBlnFlgChgValue = True
End Function


'===========================================================================
' Function Name : OpenApAcct()
' Function Desc : OpenApAcct Reference Popup
'===========================================================================
Function OpenApAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg
    
	If IsOpenPop = True Or UCase(frm1.txtBpCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����ޱݰ��� �˾�"	
	arrParam(1) = "a_jnl_acct_assn a, a_acct b"
	arrParam(2) = Trim(frm1.txtApAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.trans_type = " & FilterVar("AS002", "''", "S") & "  and A.Acct_cd = B.Acct_cd and Jnl_cd = " & FilterVar("AP", "''", "S") & " "
	arrParam(5) = "�����ޱݰ��� �ڵ�"
	
    arrField(0) = "a.acct_cd"	
    arrField(1) = "b.acct_nm"
    
    arrHeader(0) = "�����ޱݰ��� �ڵ�"		
    arrHeader(1) = "�����ޱݰ�����"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtApAcctCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetApAcct(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetApAcct()
'	Description : 
'=======================================================================================================
Function SetApAcct(byval arrRet)
	frm1.txtApAcctCd.Value		= arrRet(0)
	frm1.txtApAcctNm.Value		= arrRet(1)
	lgBlnFlgChgValue = True
End Function

'===========================================================================
' Function Name : OpenExpAcct()
' Function Desc : OpenExpAcct Reference Popup
'===========================================================================
Function OpenExpAcct()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    dim field_fg
    
	If IsOpenPop = True Or frm1.Rb_CPT.checked = true Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������˾�"	
	arrParam(1) = "a_jnl_acct_assn a, a_acct b"
	arrParam(2) = Trim(frm1.txtExpAcctCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.trans_type = " & FilterVar("AS002", "''", "S") & "  and A.Acct_cd = B.Acct_cd and Jnl_cd = " & FilterVar("COST", "''", "S") & " "
	arrParam(5) = "��������ڵ�"
	
    arrField(0) = "a.acct_cd"	
    arrField(1) = "b.acct_nm"
    
    arrHeader(0) = "��������ڵ�"		
    arrHeader(1) = "���������"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtExpAcctCd.focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetExpAcct(arrRet)
	End If	
End Function

'=======================================================================================================
'	Name : SetApAcct()
'	Description : 
'=======================================================================================================
Function SetExpAcct(byval arrRet)
	frm1.txtExpAcctCd.Value		= arrRet(0)
	frm1.txtExpAcctNm.Value		= arrRet(1)
	lgBlnFlgChgValue = True
End Function



Function OpenDept()

	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	arrParam(0) = Trim(frm1.txtDeptCd.value) 'strCode		            '  Code Condition
   	arrParam(1) = frm1.txtChgDt.Text

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupDtA2.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,"DeptCd")
	End If	
End Function

Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg
		
			case "DeptCd"
				.txtChgDt.text			= arrRet(3)
				.txtDeptCd.value        = arrRet(0)
				.txtDeptNm.value 		= arrRet(1)
				Call txtDeptCd_OnChange()

		End select	

		lgBlnFlgChgValue = True
	End With
End Function


'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================

Function MaxSpreadVal(byval Row)
  Dim iRows
  Dim MaxValue  
  Dim tmpVal

	MAxValue = 0
	with frm1
		For iRows = 1 to  .vspdData.MaxRows
			.vspddata.row = iRows
		        .vspddata.col = C_Seq

			if .vspdData.Text = "" then 
			   tmpVal = 0
			else
  			   tmpVal = cdbl(.vspdData.value)
			end if

			if tmpval > MaxValue   then
			   MaxValue = cdbl(tmpVal)

			end if
		Next

		MaxValue = MaxValue + 1

		.vspddata.row = row
		.vspddata.col = C_Seq
		.vspddata.text = MaxValue
	end with

end Function
 
Sub PopSaveSpreadColumnInf()
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
	Dim iRow
	Dim varData
	
	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
	Call InitSpreadSheet()
	'Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
	
	Call SetSpreadColor(-1,-1)

 	With frm1	
		.vspdData.Redraw = False			 	
		For iRow = 1 To frm1.vspdData.MaxRows
	
			.vspdData.Col = C_RcptType		
			.vspdData.Row = iRow
			
			varData = frm1.vspdData.text
			
			Call subVspdSettingChange(iRow,varData)   ''''Rcpt Type�� �Է��ʼ� �ʵ� ǥ�� 
		Next
		
		.vspdData.Redraw = True			
	End With
	
End Sub



'======================================================================================================
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()
'    Call GetGlobalVar
'    Call ClassLoad                                                          'Load Common DLL
    Call LoadInfTB19029                                                     'Load table , B_numeric_format
    
    Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field                         
                                                                            'Format Numeric Contents Field                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    frm1.txtAcqAmt.AllowNull =false
    frm1.txtTotalAmt.AllowNull =false
    frm1.txtApAmt.AllowNull =false
    frm1.txtAcqQty.AllowNull =false
    frm1.txtInvQty.AllowNull =false
    
    frm1.hORGCHANGEID.value =parent.gChangeOrgId 
       
    Call InitSpreadSheet                                                    'Setup the Spread sheet
    Call InitVariables                                                      'Initializes local global variables
    Call SetDefaultVal

    Call SetToolBar("1110110100101111")										' ó�� �ε�� ǥ�� �� ���� 

    frm1.txtChgNo.focus 

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

'=======================================================================================================
'   Event Name : txtPrpaymDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtChgDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtChgDt.Action = 7
    End If
End Sub
Sub txtDueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDueDt.Action = 7
    End If
End Sub

Sub txtIssuedDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtIssuedDt.Action = 7
    End If
End Sub


'======================================================================================================
'   Event Name :vspddata_EditChange
'   Event Desc :
'=======================================================================================================
Sub vspdData_EditChange(ByVal Col , ByVal Row )
Dim DblEntryQty

    With frm1.vspdData 
        If Col = C_Amt then
            .Col = C_Amt
            If .Text = "" Then
                DblEntryQty = 0
            Else
                DblEntryQty = UNICDbl(.Text)
            End If
        End If
    
    End With
                
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim intIndex
	Dim varData

	With frm1.vspdData
	
		.Row = Row
    
		frm1.vspdData.ReDraw = False
		Select Case Col
			Case  C_RcptType
				.Col = Col
				intIndex = .Value
				.Col = C_RcptType
				.Value = intIndex
				varData = .text
				If Trim(varData) <> "" Then 
					IF CommonQueryRs( " A.MINOR_CD,A.MINOR_NM " , "B_MINOR A, B_CONFIGURATION B  " , "  (A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) AND A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND A.MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						Select Case UCase(lgF0)
							Case "DP" & Chr(11)			' ������ 
								.Row  = Row
								.Col  = C_NoteNo
								.Text = ""
							Case "NO" & Chr(11)
								.Row  = Row
								.Col  = C_BankAcct
								.Text = ""
							Case Else
								.Row  = Row
								.Col  = C_NoteNo
								.Text = ""

								.Row  = Row
								.Col  = C_BankAcct
								.Text = ""
						End Select
						.Col  = C_RcptTypeNm
						.Text = Replace(lgF1, Chr(11), "")
					Else
						Call DisplayMsgBox("179051", "X", "X" ,"x")
						.Col  = C_RcptType
						.Text = ""
						.Col  = C_RcptTypeNm
						.Text = ""
						Call SetActiveCell(frm1.vspdData,C_RcptType,frm1.vspdData.ActiveRow ,"M","X","X")
					End if
				End if

				'.Col  = C_Amt
				'.Text = ""
				'.Col  = C_LocAmt
				'.Text = ""

				call subVspdSettingChange(Row,varData)
				

            '2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
			Case  C_BuildAsstNo

				.Col     = Col
				varData  = .text
				If Trim(varData) <> "" Then 
					IF CommonQueryRs( " COUNT(*) " , "A_GL_ITEM A (NOLOCK), A_GL_DTL B (NOLOCK), A_ACCT C (NOLOCK)  " , "  A.GL_NO = B.GL_NO AND A.ITEM_SEQ = B.ITEM_SEQ AND A.ACCT_CD = C.ACCT_CD AND B.CTRL_CD =  " & FilterVar("CA", "''", "S") & " AND	B.CTRL_VAL = " & FilterVar(varData , "''", "S") , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then

                        if UniCdbl(lgF0) > 0 Then
                        
                        Else
						    .Col  = C_BuildAsstNo
						    .Text = ""
						    Call SetActiveCell(frm1.vspdData,C_BuildAsstNo, frm1.vspdData.ActiveRow ,"M","X","X")                        
                        End If

					Else
						Call DisplayMsgBox("970000", "X", frm1.vspdData.ActiveRow & "���� �Ǽ����� �ڻ��ȣ��ȣ" ,"x")
						.Col  = C_BuildAsstNo
						.Text = ""
						Call SetActiveCell(frm1.vspdData,C_BuildAsstNo, frm1.vspdData.ActiveRow ,"M","X","X")
					End if
				End if
	
		End Select
	End With

	frm1.vspdData.ReDraw = True	


	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If

    gMouseClickStatus = "SPC"	'Split �����ڵ� 

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
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

Sub subVspdSettingChange(ByVal lRow, Byval varData)	

	ggoSpread.Source = frm1.vspdData
	IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)				
				Case "DP" & Chr(11)			' ������	
					ggoSpread.SSSetRequired	 C_BankAcct,		 lRow, lRow			
					ggoSpread.SpreadUnLock   C_BankAcct,      lRow, C_BankAcct
					ggoSpread.SpreadUnLock   C_BankAcctPopUp, lRow, C_BankAcctPopUp
					'ggoSpread.SSSetEdit		 C_BankAcct, "�������ڵ�", 25, 0, lRow, 30,2   
					ggoSpread.SSSetRequired	 C_BankAcct,      lRow, lRow	
					ggoSpread.SpreadLock     C_NoteNo,		 lRow, C_NoteNo,lRow   '������ȣ protect
					ggoSpread.SSSetProtected C_NoteNo,       lRow, lRow						
					ggoSpread.SpreadLock     C_NoteNoPopup,  lRow, C_NoteNoPopup,lRow          
	
				Case "NO" & Chr(11)				
					ggoSpread.SpreadUnLock   C_NoteNo,        lRow, C_NoteNo,       lRow
					ggoSpread.SpreadUnLock   C_NoteNoPopup,   lRow, C_NoteNoPopup,  lRow
					 
					ggoSpread.SpreadLock     C_BankAcct,      lRow, C_BankAcct,     lRow   
					ggoSpread.SpreadLock     C_BankAcctPopup, lRow, C_BankAcctPopup,lRow
		
					ggoSpread.SSSetProtected C_BankAcct,      lRow, lRow								
		
					'ggoSpread.SSSetEdit      C_NoteNo, "������ȣ", 25, 0, lRow, 30,2	
					ggoSpread.SSSetRequired  C_NoteNo,        lRow, lRow
		
				Case Else									
					ggoSpread.SpreadLock     C_BankAcct,      lRow, C_BankAcct,     lRow   			
					ggoSpread.SpreadLock     C_BankAcctPopup, lRow, C_BankAcctPopup,lRow
							
					ggoSpread.SSSetProtected C_BankAcct,      lRow, lRow							
		
					ggoSpread.SpreadLock     C_NoteNo,        lRow, C_NoteNo,     lRow
					ggoSpread.SpreadLock     C_NoteNoPopup,   lRow, C_NoteNoPopup,lRow		
		
					ggoSpread.SSSetProtected C_NoteNo,        lRow, lRow													
			End Select			
		
	end if
	
End Sub	


'==========================================================================================
'   Event Name : txtVatType_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtVatType_OnChange()
	Dim dblVatAmt
	
	lgBlnFlgChgValue = True
	
	if frm1.txtVatAmt.text = "" then
		dblVatAmt = 0
	else
		dblVatAmt = UNICDbl(frm1.txtVatAmt.text)	
	end if
	
	If Trim(frm1.txtVatType.Value) = "" and  dblVatAmt = 0 Then
		ggoOper.SetReqAttr frm1.txtVatAmt, "D"    '�ΰ����ݾ� 
		ggoOper.SetReqAttr frm1.txtVatType, "D"    '�ΰ���Ÿ�� 
	Else
		ggoOper.SetReqAttr frm1.txtVatAmt, "N"    '�ΰ����ݾ� 
		ggoOper.SetReqAttr frm1.txtVatType, "N"    '�ΰ���Ÿ�� 
	End If

End Sub


'==========================================================================================
'   Event Name : txtVatAmt_Change()
'   Event Desc : 
'==========================================================================================
Sub txtVatAmt_Change()
	Dim dblVatAmt

	lgBlnFlgChgValue = True	
	
	if frm1.txtVatAmt.text="" then
		dblVatAmt = 0
	else
		dblVatAmt = UNICDbl(frm1.txtVatAmt.text)	
	end if
		
	If Trim(frm1.txtVatType.Value) = "" and  dblVatAmt = 0 Then
		ggoOper.SetReqAttr frm1.txtVatAmt, "D"    '�ΰ����ݾ� 
		ggoOper.SetReqAttr frm1.txtVatType, "D"    '�ΰ���Ÿ�� 
	Else
		ggoOper.SetReqAttr frm1.txtVatAmt, "N"    '�ΰ����ݾ� 
		ggoOper.SetReqAttr frm1.txtVatType, "N"    '�ΰ���Ÿ�� 
	End If
		
End Sub





'==========================================================================================
'   Event Name : txtDocCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChange()
 
	lgBlnFlgChgValue = True
	If UCase(frm1.txtDocCur.value) <> parent.gCurrency Then               ' �ŷ���ȭ�ϰ� Company ��ȭ�� �ٸ��� ȯ���� 0���� ���� 
		frm1.txtXchRate.text	= "0"                         ' ����Ʈ���� 1�� �� ������ ȯ���� �Էµ� ������ �Ǵ��Ͽ� 
							                                        ' ȯ�������� ���� �ʰ� �Էµ� ������ ���. 
	Else 
		frm1.txtXchRate.text	= "1"
	End If			

    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							

		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    
End Sub


'==========================================================================================
'   Event Name : txtDocCur_OnChangeASP
'   Event Desc : 
'==========================================================================================
Sub txtDocCur_OnChangeASP()
 
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtDocCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							

		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()

	END IF	    
End Sub

'==========================================================================================
'   Event Name : txtCur_OnChange
'   Event Desc : 
'==========================================================================================
Sub txtCur_OnChange()

    lgBlnFlgChgValue = True
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.txtCur.value , "''", "S") & "" , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
		Call CurFormatNumericOCXRef()
	END IF	    
End Sub

Sub txtDeptCd_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If Trim(frm1.txtRegDt.Text = "") Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

		'----------------------------------------------------------------------------------------
		strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
		strFrom		=			 " b_acct_dept(NOLOCK) "		
		strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
		strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
		strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtChgDt.Text, gDateFormat,""), "''", "S") & "))"			

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
			Next	
			
		End If
	
		'----------------------------------------------------------------------------------------

End Sub

Sub txtChgDt_onBlur()
    
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2

	lgBlnFlgChgValue = True
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtChgDt.Text <> "") Then
			'----------------------------------------------------------------------------------------
				strSelect	=			 " Distinct org_change_id "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtChgDt.Text, gDateFormat,""), "''", "S") & "))"			
	
			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hORGCHANGEID.value) Then
					'IntRetCD = DisplayMsgBox("124600","X","X","X") 
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hORGCHANGEID.value = ""
					.txtDeptCd.focus
			End if
		End If
	End With
'----------------------------------------------------------------------------------------

End Sub
'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
    
    Dim strTemp
    Dim intPos1
    Dim bankCode
	Dim intRetCd
	Dim strData
	Dim strCard
	
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		if Row > 0 And Col = C_BankAcctPopup Then

			.Col = C_BankAcct
			.Row = Row
			
			Call OpenPopup(.Text,4)
		
		Elseif Row > 0 And Col = C_NoteNoPopUp Then
			
			.Col = C_NoteNo
			.Row = Row
			strTemp = Trim(.text)				    
			.col = C_RcptType
			strCard = .text

			Call OpenNoteNo(strTemp, strCard)
		Elseif Row > 0 And Col = C_RcptTypePopup Then
			.Col = C_RcptType
			.Row = Row
			Call OpenPopup(.Text, 6)
		
		'2010.05.31...kbs...�Ǽ������ڻ��ȣ �߰�
		Elseif Row > 0 And Col = C_BuildAsstNoPopup Then
				.Col = C_BuildAsstNo
				.Row = Row
				Call OpenBuildAsstNoPopup(.Text, Row)	
							
		End If


	
	End With
End Sub


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

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKey <> "" Then
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End if

End Sub


'==========================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'==========================================================================================

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim varData

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
			Case  C_RcptTypeNm
				.Col = Col
				intIndex = .Value
				.Col = C_RcptType
				.Value = intIndex
				varData = .text
		End Select
	End With	

	frm1.vspdData.ReDraw = False	
	
	 IF CommonQueryRs( "REFERENCE" , "B_CONFIGURATION  " , "MAJOR_CD = " & FilterVar("A1006", "''", "S") & "  AND MINOR_CD =  " & FilterVar(varData , "''", "S") & " AND SEQ_NO = 4 ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then         
		   Select Case UCase(lgF0)					
				Case "DP" & Chr(11)			' ������ 
					frm1.vspdData.Row  = Row
					frm1.vspdData.Col  = C_NoteNo
					frm1.vspdData.Text = ""		
				Case "NO" & Chr(11)															
					frm1.vspdData.Row  = Row
					frm1.vspdData.Col  = C_BankAcct
					frm1.vspdData.Text = ""			
				Case Else
					frm1.vspdData.Row  = Row
					frm1.vspdData.Col  = C_NoteNo
					frm1.vspdData.Text = ""			
							
					frm1.vspdData.Row  = Row
					frm1.vspdData.Col  = C_BankAcct
					frm1.vspdData.Text = ""				
			End Select			
	end if

	frm1.vspdData.Col  = C_Amt
	frm1.vspdData.Text = ""	
	frm1.vspdData.Col  = C_LocAmt
	frm1.vspdData.Text = ""	

	call subVspdSettingChange(Row,varData)

	frm1.vspdData.ReDraw = True	

End Sub

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing

  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
    
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
    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables                                                      'Initializes local global variables
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    'Call InitComboBox

  '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery															'Query db data
       
    FncQuery = True															
    
End Function

'======================================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'=======================================================================================================
Function FncNew() 
	Dim IntRetCD 
	
	FncNew = False                                                          
	
	'-----------------------
	'Check previous data area
	'-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
	'-----------------------
	'Erase condition area
	'Erase contents area
	'-----------------------
	Call ggoOper.ClearField(Document, "A")                                  'Clear Condition Field
	ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData
	
	Call ggoOper.LockField(Document, "N")                                   'Lock  Suitable  Field
	Call InitVariables                                                      'Initializes local global variables

	Call SetDefaultVal
    Call SetToolBar("1110110100100111")										' ó�� �ε�� ǥ�� �� ���� 

	call txtDocCur_OnChangeASP()   

	FncNew = True 
	
	'SetGridFocus                                                          

End Function

'======================================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================

Function FncDelete() 
    Dim IntRetCD
	FncDelete = False
		
	IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")   '�����Ͻðڽ��ϱ�?  
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	'-----------------------
	'Precheck area
	'-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        intRetCD = DisplayMsgBox("900002","x","x","x")                                
    	Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete                                                          '��: Delete db data
    
    FncDelete = True

End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
	Dim IntRetCD 
	
	FncSave = False
	
	Err.Clear                                                               

    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer   

    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then  '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","x","x","x")            '��: Display Message(There is no changed data.)
        Exit Function
    End If
      
    If Not chkField(Document, "2") Then               '��: Check required field(Single area)
       Exit Function
    End If

	if isNull(frm1.txtApAmt.text) then
		frm1.txtApAmt.text = "0"
	end if

	if frm1.txtApAmt.text = "" then
		frm1.txtApAmt.text = "0"
	end if			
	
'	if IsNull(frm1.txtIssuedDt.text) or Trim(frm1.txtIssuedDt.text) = "" then
'		frm1.txtIssuedDt.text = frm1.txtChgDt.text		
'	end if
	
	If CompareDateByFormat(frm1.txtRegDt.text,frm1.txtChgDt.text,frm1.txtRegDt.Alt,frm1.txtChgDt.Alt, _
        	               "970023",frm1.txtRegDt.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtChgDt.focus
	   Exit Function
	End If
    'if frm1.vspdData.MaxRows < 1 then  'fpDoubleSingle8
	'	if  frm1.txtApAmt.text = 0 then
	'		IntRetCD = DisplayMsgBox("117991","X","X","X")  ''�ڻ����� �ݾ��� �Է��Ͻʽÿ�.
	'		Exit Function
	'	end if		
	'end if
'	ȸ��μ��� ���� ��쿡 �޼��� ��� 
'	If Trim(frm1.txtAcctDeptNm) <> Trim(frm1.txtDeptNm)
'		IntRetCD = DisplayMsgBox("127801","X","X")
'		Exit Function
'	End if
    ggoSpread.Source = frm1.vspdData                          '��: Preset spreadsheet pointer 
    If Not ggoSpread.SSDefaultCheck         Then              '��: Check required field(Multi area)
       Exit Function
    End If

	'-----------------------
	'Save function call area
	'-----------------------
	IF  DbSave	= False Then
		   Exit Function
    End If		                                                '��: Save db data	 
	FncSave = True
	
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'=======================================================================================================
Function FncCopy()
   	frm1.vspdData.ReDraw = False
    	
    if frm1.vspdData.MaxRows < 1 then Exit Function
    	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    frm1.vspdData.Col = C_RcptType	
	call subVspdSettingChange(frm1.vspdData.ActiveRow,frm1.vspdData.Text)

'	frm1.vspdData.Col = C_RcptType
'	frm1.vspdData.Text = ""

'	frm1.vspdData.Col = C_RcptTypeNm
'	frm1.vspdData.Text = ""

	MaxSpreadVal frm1.vspdData.ActiveRow
        
    frm1.vspdData.ReDraw = True
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'=======================================================================================================
Function FncCancel() 
	Dim iDx
	
	FncCancel = False

    if frm1.vspdData.MaxRows < 1 then Exit Function
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.EditUndo
	
	Frm1.vspdData.Row = frm1.vspdData.ActiveRow
   	Frm1.vspdData.Col = C_RcptType
    iDx = Frm1.vspdData.Value
    Frm1.vspdData.Col = C_RcptTypeNm
    Frm1.vspdData.value = iDx
     
    Set gActiveElement = document.ActiveElement   
     
    FncCancel = True
	
End Function

'======================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'=======================================================================================================
Function FncInsertRow(ByVal pvRowCnt)

	Dim varMaxRow
	Dim imRow, indx
	FncInsertRow = False

	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowcount()

	If ImRow="" then
		Exit Function
	End If
	End If

	
	With frm1
	
        .vspdData.focus
        ggoSpread.Source = .vspdData
        '.vspdData.EditMode = True
        .vspdData.ReDraw = False
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        

        for indx = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
			'MaxSpreadVal .vspdData.ActiveRow
			call MaxSpreadVal(indx)
		next		
		.vspdData.ReDraw = True		

    End With
    

    Set gActiveElement = document.ActiveElement  
	If Err.number = 0 Then
	   FncInsertRow = True                                                          '��: Processing is OK
	End If       
End Function

'======================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'=======================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    if frm1.vspdData.MaxRows < 1 then Exit Function

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
'		lgBlnFlgChgValue = True
    End With
    
End Function

'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint() 
    Call parent.FncPrint()                                              
End Function

'======================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'=======================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
    '�ϼ��� ���� �ʾ��� 
End Function

'======================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'=======================================================================================================
Function FncNext() 
    On Error Resume Next
    '�ϼ��� ���� �ʾ���                                                    
End Function

'=======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)										
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : ȭ�� �Ӽ�, Tab���� 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLEMULTI , True)                               
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

'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	    	
	
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
		
    End If
    
    FncExit = True
    
End Function

'=======================================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'=======================================================================================================
Function DbDelete() 
    Dim strVal
    
    DbDelete = False														'��: Processing is NG 
 
 	Call LayerShowHide(1)   
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003
    strVal = strVal & "&txtChgNo=" & Trim(frm1.txtChgNo.value)			'��: ���� ���� ����Ÿ 

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
	 
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
    
    DbDelete = True                                                         '��: Processing is NG
End Function

'=======================================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================================
Function DbDeleteOk()		

	lgBlnFlgChgValue = False         '���� ������ ���� ���� 
	Call FncNew()
End Function

'=======================================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================================
Function DbQuery() 
    
	DbQuery = False                                                         
	
	Call LayerShowHide(1)
	
	Dim strVal
	
	With frm1
	
        	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						'��: 
        	strVal = strVal     & "&txtChgNo=" & Trim(.txtChgNo.value)	'��ȸ ���� ����Ÿ 
        	strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey

			' ���Ѱ��� �߰� 
			strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
			strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
			strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
			strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
			    	
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)										'�����Ͻ� ASP �� ���� 
	
	DbQuery = True                                                          
    
End Function

'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================================
Function DbQueryOk()													'��ȸ ������ ������� 
	Dim varData
	Dim iRow
	
	lgIntFlgMode = parent.OPMD_UMODE												'Indicates that current mode is Update mode
	
	Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field	

    if frm1.Rb_CPT.checked = True then 
		Call radio1_onchange
	else
		Call radio2_onchange
	end if

	Call SetToolBar("1111111100111111")									'��ư ���� ���� 
	
	Call SetSpreadColor(-1,-1)

 	With frm1	
 	
		For iRow = 1 To frm1.vspdData.MaxRows
	
			.vspdData.Col = C_RcptType		
			.vspdData.Row = iRow
			
			varData = frm1.vspdData.text
			
			Call subVspdSettingChange(iRow,varData)   ''''Rcpt Type�� �Է��ʼ� �ʵ� ǥ�� 
		Next
		
		.vspdData.Redraw = True			
	End With
	
	call txtDocCur_OnChangeASP()
	call txtCur_OnChange()
	Call txtVatAmt_Change()
	Call txtVatType_OnChange() 
	'SetGridFocus
	lgBlnFlgChgValue = False	
    frm1.txtChgNo.focus 
    
End Function

Sub InitData()
	Dim intRow
	Dim intIndex 
	dim temp
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row  = intRow
			
			.Col	 = C_RcptType
			intIndex = .Value 			

			.Col     = C_RcptTypeNm
			.Value   = intindex					
		Next	
	End With	

End Sub

'=======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
	
	Dim IntRows 
	Dim IntCols 
	
	Dim lGrpcnt 
	Dim strVal
	Dim strDel
	
	DbSave = False                                                          
	
	'On Error Resume Next                                                   
	
	Call LayerShowHide(1)
	
	'Call SetSumItem
	
	strVal = ""
	strDel = ""
	
	With frm1
		.txtMode.value = parent.UID_M0002											'��: ���� ���� 
		.txtFlgMode.value = lgIntFlgMode									'��: �ű��Է�/���� ���� 
	End With
	
	'-----------------------
	'Data manipulate area
	'-----------------------
	' Data ���� ��Ģ 
	' 0: Flag , 1: Row��ġ, 2~N: �� ����Ÿ 
	
	lGrpCnt = 1

	With frm1.vspdData
	    
		For IntRows = 1 To .MaxRows
		
			.Row = IntRows
			.Col = 0

			If .Text = ggoSpread.DeleteFlag Then
				strDel = strDel & "D" & parent.gColSep & IntRows & parent.gColSep		'D=Delete

			ElseIf .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "U" & parent.gColSep & IntRows & parent.gColSep		'U=Update

			Else				
				strVal = strVal & "C" & parent.gColSep & IntRows & parent.gColSep		'C=Create, Sheet�� 2�� �̹Ƿ� ����			
			End If
		
			Select Case .Text
				Case ggoSpread.DeleteFlag
					.Col = C_Seq
					strDel = strDel & Trim(.Text) &parent.gRowSep				    '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 
					
					lGrpcnt = lGrpcnt + 1             
								
				Case Else
					
					.Col = C_Seq
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_RcptType
					strVal = strVal & Trim(.Text) & parent.gColSep
					.Col = C_Amt
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep
					.Col = C_LocAmt
					strVal = strVal & UNIConvNum(Trim(.Text),0) & parent.gColSep					
					.Col = C_BankAcct
					strVal = strVal & Trim(.Text) & parent.gColSep					
					.Col = C_NoteNo
					strVal = strVal & Trim(.Text) & parent.gColSep				    '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 

                    '2010.05.31...kbs..�Ǽ������ڻ��ȣ �߰�
                    .Col = C_BuildAsstNo
                    strVal = strVal & Trim(.Text) & parent.gRowSep				    '������ ����Ÿ�� Row �и���ȣ�� �ִ´� 

					lGrpCnt = lGrpCnt + 1

			End Select
		Next

	End With

	frm1.txtMaxRows.value = lGrpCnt-1										'��: Spread Sheet�� ����� �ִ밹�� 
	frm1.txtSpread.value = strDel & strVal									'��: Spread Sheet ������ ���� 

	'���Ѱ����߰� start
	frm1.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
	frm1.txthInternalCd.value =  lgInternalCd
	frm1.txthSubInternalCd.value = lgSubInternalCd
	frm1.txthAuthUsrID.value = lgAuthUsrID		
	'���Ѱ����߰� end
		
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)								'��: ���� �����Ͻ� ASP �� ���� 

	DbSave = True                                                           
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'=======================================================================================================
Function DbSaveOk()													'��: ���� ������ ���� ���� 

   	lgBlnFlgChgValue = false	

    Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggospread.ClearSpreadData		'Buffer Clear
	
    Call InitVariables                                                      'Initializes local global variables
	frm1.vspdData.MaxRows = 0 ' InitSpreadSheet ��� 
	Call DbQuery	

End Function

'======================================================================================================
' Area Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================
Function NoteValidCheck()
	For iRow = 1 To frm1.vspdData.MaxRows
	
		.vspdData.Col = C_RcptType		
		.vspdData.Row = iRow
			
		varData = frm1.vspdData.text
			
		Call subVspdSettingChange(iRow,varData)   ''''Rcpt Type�� �Է��ʼ� �ʵ� ǥ�� 
	Next

End Function


Function Radio1_onChange()
	
	frm1.txtExpAcctCd.value = ""
	frm1.txtExpAcctNm.value = ""
	
	ggoOper.SetReqAttr frm1.txtExpAcctCd, "Q"    '�̼��ݰ��� 

	lgBlnFlgChgValue = True
End Function

Function Radio2_onChange()
	
	ggoOper.SetReqAttr frm1.txtExpAcctCd, "N"    '�̼��ݰ��� 
	lgBlnFlgChgValue = True
End Function

function txtDeptCd_onblur()
	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
end function

function txtBpCd_onblur()
	if Trim(frm1.txtBpCd.value) = "" then 		
		frm1.txtBpNm.value = ""		
	end if	
End function

Function txtDueDt_Change
	lgBlnFlgChgValue = True
End Function


Function txtIssuedDt_Change
	lgBlnFlgChgValue = True
End Function

Function txtXchRate_Change
	lgBlnFlgChgValue = True
End Function

Function txtChgAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtChgLocAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtTotalAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtTotalLocAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtApAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtApLocAmt_Change
	lgBlnFlgChgValue = True
End Function

Function txtVatRate_Change
	lgBlnFlgChgValue = True
End Function

Function txtVatLocAmt_Change
	lgBlnFlgChgValue = True
End Function
'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1

'		ggoOper.FormatFieldByObjectOfCur .txtAcqAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtTotalAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtApAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		ggoOper.FormatFieldByObjectOfCur .txtVatAmt, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
		

	End With

End Sub
'===================================== CurFormatNumericOCXRef()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCXRef()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtAcqAmt, .txtCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, parent.gComNum1000, parent.gComNumDec
	End With

End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'�ݾ� 
		ggoSpread.SSSetFloatByCellOfCur C_Amt,-1, .txtDocCur.value, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec		
		
	End With

End Sub
    
'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>
<!--
'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
'======================================================================================================= -->
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenMasterRef()">�ڻ긶��������</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupTempGL()">������ǥ</A>&nbsp;|&nbsp;<A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>
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
									<TD CLASS="TD5" NOWRAP>�ڻ꺯����ȣ</TD>
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtChgNo" SIZE=20 MAXLENGTH=18 tag="12XXXU" ALT="�ڻ꺯����ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPrpaymNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenChgNoInfo"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=20 valign=top>
						<TABLE <%=LR_SPACE_TYPE_50%>>
									<TR>
										<TD CLASS="TD5" NOWRAP>�ڻ��ȣ</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstNo" SIZE=20 MAXLENGTH=18 TAG="23XXXU" ALT="�ڻ��ȣ"> <INPUT TYPE="Text" NAME="txtAsstNm" SIZE=25 MAXLENGTH=40 tag="24X" ALT="�ڻ��"></TD>
										<TD CLASS="TD5" NOWRAP>�������</TD>
										<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtRegDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="�������" tag="24X1" id=fpDateTime1></OBJECT>');</SCRIPT></TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>�����μ�</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctDeptNm" SIZE=27 MAXLENGTH=40 tag="24XXXU" ALT="ȸ��μ���"></TD>
										<TD CLASS="TD5" NOWRAP>���-�������</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle0 name="txtAcqQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="�������" tag="24X3"> </OBJECT>');</SCRIPT>&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name="txtInvQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="�������" tag="24X3"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
<%	If gIsShowLocal <> "N" Then	%>
									<TR>
										<TD CLASS=TD5 NOWRAP>�ŷ���ȭ</TD>
										<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCur" ALT="�ŷ���ȭ" TYPE="Text" MAXLENGTH=3 SIZE=10 tag="24" ></TD>
										<TD CLASS=TD5 NOWRAP>ȯ��</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtXchRt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="ȯ��" tag="24X5Z"> </OBJECT>');</SCRIPT>&nbsp;
										</TD>
									</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtCur"><INPUT TYPE=HIDDEN NAME="txtXchRt">
<%	End If %>
									<TR>
										<TD CLASS=TD5 NOWRAP>���ݾ�</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtAcqAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="���ݾ�" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
										</TD>
<%	If gIsShowLocal <> "N" Then	%>
										<TD CLASS=TD5 NOWRAP>���ݾ�(�ڱ�)</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 name=txtAcqLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="���ݾ�(�ڱ�)" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
										</TD>
<%	ELSE %>
<TD CLASS=TD5 NOWRAP></TD><TD CLASS=TD6 NOWRAP><INPUT TYPE=HIDDEN NAME="txtAcqLocAmt"></TD>
<%	End If %>
									</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtChgDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��������" tag="22X1" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
								<TD CLASS="TD5" NOWRAP>ȸ��μ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDeptCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="ȸ��μ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDept()">&nbsp;<INPUT TYPE=TEXT NAME="txtDeptNm" SIZE=20 tag="24" ALT="ȸ��μ���"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>��������</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_CPT Checked tag = 2 value="01" onclick=radio1_onchange()><LABEL FOR=Rb_CPT>�ں�������</LABEL>&nbsp;<INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_Rve tag = 2 value="02" onclick=radio2_onchange()><LABEL FOR=Rb_Rve>����������</LABEL></TD>
								<TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="22XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenBp(frm1.txtBpCd.value, 1)">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="24" ALT="�ŷ�ó��"></TD>
							</TR>
<%	If gIsShowLocal <> "N" Then	%>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ŷ���ȭ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDocCur" ALT="�ŷ���ȭ" TYPE="Text" SIZE=10 tag="22XXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocCur" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.value, 3)"></TD>
								<TD CLASS="TD5" NOWRAP>ȯ��</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtXchRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="ȯ��" tag="24X5Z" id=fpDoubleSingle5></OBJECT>');</SCRIPT></TD>
							</TR>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtDocCur"><INPUT TYPE=HIDDEN NAME="txtXchRate">
<%	End If %>
							<TR>
								<TD CLASS="TD5" NOWRAP>�������</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotalAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�������" tag="22X2" id=fpDoubleSingle6></OBJECT>');</SCRIPT></TD>
<%	If gIsShowLocal <> "N" Then	%>
								<TD CLASS="TD5" NOWRAP>�������(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtTotalLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�������(�ڱ�)" tag="21X2" id=fpDoubleSingle7></OBJECT>');</SCRIPT></TD>
<%	ELSE %>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=HIDDEN NAME="txtTotalLocAmt"></TD>
<%	End If %>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ޱݰ���</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtApAcctCd" SIZE=10 MAXLENGTH=20 tag="21XXXU" ALT="�����ޱݰ���"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApAcctCd" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenApAcct()">&nbsp;<INPUT TYPE=TEXT NAME="txtApAcctNm" SIZE=22 tag="24"  alt = "�����ޱݰ�����"></TD>
								<TD CLASS="TD5" NOWRAP>�������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtExpAcctCd" SIZE=10 MAXLENGTH=20 tag="24XXXU" ALT="�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnExpAcctCd" ALIGN=Top TYPE="BUTTON" ONCLICK="vbscript:OpenExpAcct()">&nbsp;<INPUT TYPE=TEXT NAME="txtExpAcctNm" SIZE=22 tag="24"  alt = "���������"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ޱݾ�</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtApAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�����ޱݾ�" tag="24X2" id=fpDoubleSingle8></OBJECT>');</SCRIPT></TD>
<%	If gIsShowLocal <> "N" Then	%>
								<TD CLASS="TD5" NOWRAP>�����ޱݾ�(�ڱ�)</TD>
	                            <TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtApLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�����ޱݾ�(�ڱ�)" tag="24X2" id=fpDoubleSingle9></OBJECT>');</SCRIPT></TD>
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtApLocAmt">
<%	End If %>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�����ޱ� ��������</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtDueDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="21X1" ALT="�����ޱ� ��������"> </OBJECT>');</SCRIPT>
								</TD>
								<TD CLASS="TD5" NOWRAP>�����ޱ� ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtApNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="�����ޱ� ��ȣ"></TD>
							</TR>
<%	If gIsShowLocal <> "N" Then	%>
							<TR>
								<TD CLASS="TD5" NOWRAP>�ΰ�������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtVatType" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="�ΰ�������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVatType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopup(frm1.txtVatType.value, 5)">&nbsp;<INPUT TYPE=TEXT NAME="txtVatTypeNm" SIZE=20 tag="24" ALT="�ΰ�������"></TD>
								<TD CLASS="TD5" NOWRAP>�ΰ�����</TD>
								<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtVatRate style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="�ΰ�����" tag="21X5Z" id=fpDoubleSingle5></OBJECT>');</SCRIPT>&nbsp;%</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ΰ����ݾ�</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name=txtVatAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="�ΰ����ݾ�" tag="21X2"> </OBJECT>');</SCRIPT>&nbsp;
	                            </TD>
								<TD CLASS=TD5 NOWRAP>�ΰ����ݾ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name=txtVatLocAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="�ΰ����ݾ�(�ڱ�)" tag="21X2"> </OBJECT>');</SCRIPT> &nbsp;
 									</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>�Ű������</TD>
							    <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtReportAreaCd" SIZE=10 MAXLENGTH=10 tag="21XXXU" ALT="�Ű�������ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnReportAreaCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenReportAreaCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtReportAreaNm" SIZE=20 tag="24" ALT="�Ű�������"></TD>
								<TD CLASS="TD5" NOWRAP>��꼭������</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime4 name=txtIssuedDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="21" ALT="��ǥ��������"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							
<%	ELSE %>
<INPUT TYPE=HIDDEN NAME="txtVatAmt"><INPUT TYPE=HIDDEN NAME="txtVatLocAmt"><INPUT TYPE=HIDDEN NAME="txtVatType"><INPUT TYPE=HIDDEN NAME="txtVatTypeNm"><INPUT TYPE=HIDDEN NAME="txtVatRate">
<%	End If %>
							<TR>
								<TD CLASS="TD5" NOWRAP>������ǥ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTempGlNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="������ǥ��ȣ"></TD>
								<TD CLASS="TD5" NOWRAP>ȸ����ǥ��ȣ</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtGlNo" SIZE=20 MAXLENGTH=18 tag="24" ALT="ȸ����ǥ��ȣ"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD656" NOWRAP COLSPAN=3><INPUT TYPE=TEXT NAME="txtChgDesc" SIZE=90 MAXLENGTH=30 tag="2X" ALT="����"></TD>
							</TR>
							<TR>
								<TD WIDTH="80%" HEIGHT=100% COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> id=vspdData NAME=vspdData width="100%" HEIGHT="100%" tag="2" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<TR HEIGHT=10>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
		<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>

		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1" ></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode"         tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtMaxRows"	  tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode"	  tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="hORGCHANGEID"   tag="24"TABINDEX = "-1" >
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
