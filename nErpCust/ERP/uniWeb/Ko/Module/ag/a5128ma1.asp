<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5128ma1
'*  4. Program Name         : ȸ����ǥ ���(Copy)
'*  5. Program Desc         : ȸ����ǥ������ ���, ����, ����, ��ȸ 
'*  6. Component List       : PAGG020.dll
'*  7. ModIfied date(First) : 2003/01/02
'*  8. ModIfied date(Last)  : 2003/06/05
'*  9. ModIfier (First)     : Kim Ho Young
'* 10. ModIfier (Last)      : Lim YOung Woon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="Acctctrl_ko441_1.vbs"></SCRIPT>

<SCRIPT LANGUAGE=vbscript>

Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID      = "a5104mb1.asp"			'��: �����Ͻ� ���� ASP�� 
Const JUMP_PGM_ID_TAX_REP = "a5104ma1"
'                       4.2 Constant variables 
'========================================================================================================
Const C_GLINPUTTYPE = "GL"
Const MENU_NEW	=	"1100000000011111"
Const MENU_CRT	=	"1100111100111111"
Const MENU_UPD	=	"1101111100111111"
Const MENU_PRT	=	"1100000000011111"

'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'=                       4.4 User-defind Variables
'========================================================================================================
'��: Grid Columns
Dim  C_ItemSeq		
Dim  C_deptcd		
Dim  C_deptPopup	
Dim  C_deptnm	   	
Dim  C_AcctCd		
Dim  C_AcctPopup	
Dim  C_AcctNm		
Dim  C_DrCrFg		
Dim  C_DrCrNm		
Dim  C_DocCur		
Dim  C_DocCurPopup	
Dim  C_ExchRate	
Dim  C_ItemAmt		
Dim  C_ItemLocAmt	
Dim  C_IsLAmtChange
Dim  C_ItemDesc	
Dim  C_VatType		
Dim  C_VatNm		
Dim  C_AcctCd2		


Dim lgCurrRow
Dim lgStrPrevKeyDtl
Dim lgStrPrevKey2
Dim lgStrPrevKey3
Dim lgFormLoad
Dim lgstartfnc
Dim intItemCnt

Dim IsOpenPop

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'========================================================================================================
Sub initSpreadPosVariables()
	C_ItemSeq		= 1 
	C_deptcd		= 2 
	C_deptPopup		= 3 
	C_deptnm		= 4	
	C_AcctCd		= 5 
	C_AcctPopup		= 6 
	C_AcctNm		= 7 
	C_DrCrFg		= 8 
	C_DrCrNm		= 9 
	C_DocCur		= 10
	C_DocCurPopup	= 11
	C_ExchRate		= 12
	C_ItemAmt		= 13
	C_ItemLocAmt	= 14
	C_IsLAmtChange	= 15
	C_ItemDesc		= 16
	C_VatType		= 17
	C_VatNm			= 18
	C_AcctCd2		= 19
End Sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0   
    lgStrPrevKey = ""
    lgLngCurRows = 0   
     
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	' ClearField(pDoc, Optional ByVal pStrGrp)
    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
	frm1.txtGLDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
         
    frm1.txtCommandMode.Value = "CREATE"
    frm1.cboGlInputType.Value = C_GLINPUTTYPE
    
	frm1.cboGlType.Value = "03"
	
	frm1.txtDeptCd.Value	= parent.gDepart

    frm1.vspdData3.MaxRows = 0
    frm1.vspdData3.MaxCols = 16    
    
   frm1.hOrgChangeId.Value = parent.gChangeOrgId 
    '���ݰ����� �������´�.
	Call GetCheckAcct	
	    
    frm1.txtGLNo.focus
    lgBlnFlgChgValue = False
End Sub

'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	With frm1.vspdData
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.Spreadinit "V20030218",,parent.gAllowDragDropSpread    
	
		.MaxCols = C_AcctCd2 + 1
		.Col = .MaxCols				'��: ������Ʈ�� ��� Hidden Column
		.ColHidden = True
		.MaxRows = 0
		.ReDraw = False

		Call AppEndNumberPlace("6","3","0")
        Call GetSpreadColumnPos("A")
        ggoSpread.SSSetFloat  C_ItemSeq,    " ", 4, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_deptcd,     "�μ��ڵ�",   10, , , 10, 2
        ggoSpread.SSSetButton C_deptpopup
        ggoSpread.SSSetEdit   C_deptnm,     "�μ���",     17, , , 30
		ggoSpread.SSSetEdit   C_AcctCd,     "�����ڵ�",   15, , , 18
		ggoSpread.SSSetButton C_AcctPopup
		ggoSpread.SSSetEdit   C_AcctNm,     "�����ڵ��", 20, , , 30
		ggoSpread.SSSetCombo  C_DrCrFg,     "", 8
	    ggoSpread.SSSetCombo  C_DrCrNm,     "���뱸��",   11
		ggoSpread.SSSetEdit   C_DocCur,     "�ŷ���ȭ",   10, , , 10, 2
        ggoSpread.SSSetButton C_DocCurPopup
		ggoSpread.SSSetFloat  C_ExchRate,   "ȯ��", 15, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_ItemAmt,    "�ݾ�",       15, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_ItemLocAmt, "�ݾ�(�ڱ�)", 15, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit   C_IsLAmtChange,   "",     30, , , 128
		ggoSpread.SSSetEdit   C_ItemDesc,   "��  ��",     30, , , 128
		ggoSpread.SSSetCombo  C_VATTYPE,     "", 8
	    ggoSpread.SSSetCombo  C_VATNM,     "��꼭����",   20	    		
		ggoSpread.SSSetEdit   C_AcctCd2,   "",     30, , , 128

		Call ggoSpread.MakePairsColumn(C_deptcd,C_deptpopup)
		Call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctPopup)
		Call ggoSpread.MakePairsColumn(C_DrCrFg,C_DrCrNm,"1")
		Call ggoSpread.MakePairsColumn(C_VATTYPE,C_VATNM,"1")

		Call ggoSpread.SSSetColHidden(C_ItemSeq,C_ItemSeq,True)
		Call ggoSpread.SSSetColHidden(C_DrCrFg,C_DrCrFg,True)
		Call ggoSpread.SSSetColHidden(C_VatType,C_VatType,True)
		Call ggoSpread.SSSetColHidden(C_VatNm,C_VatNm,True)
		Call ggoSpread.SSSetColHidden(C_IsLAmtChange,C_IsLAmtChange,True)
		Call ggoSpread.SSSetColHidden(C_AcctCd2,C_AcctCd2,True)

		.ReDraw = True
	End With
         										'������Ʈ�� ��� Hidden Column
    SetSpreadLock "I", 0, 1, ""
End Sub
'=======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    With frm1
        ggoSpread.Source = .vspdData
        lRow2 = .vspdData.MaxRows
        .vspdData.Redraw = False
        
        Select Case Index
			Case 0
			    ggoSpread.SSSetRequired C_deptcd    , -1    , C_deptcd
			    ggoSpread.SpreadLock C_deptnm       , -1    , C_deptnm
			    ggoSpread.SSSetRequired C_AcctCd    , -1    , C_AcctCd
				ggoSpread.SpreadLock C_AcctPopup    , -1    , C_AcctPopup
				ggoSpread.SpreadLock C_AcctNm       , -1    , C_AcctNm
			    ggoSpread.SSSetRequired C_DrCrNm    , -1    , C_DrCrNm
			    ggoSpread.SSSetRequired C_ItemAmt   , -1    , C_ItemAmt
			Case 1
				ggoSpread.SpreadLock C_ItemSeq, -1, C_AcctCd2 'Item Grid ��ü Lock���� 
        End Select
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
        .vspdData.Redraw = True
    End With
End Sub

'=======================================================================================================
Sub SetSpread2Lock(Byval stsFg,Byval Index,ByVal lRow  ,ByVal lRow2 )
    
    With frm1
		ggoSpread.Source = .vspdData2			
		If lRow = "" Then
			lRow = 1
		End If	
		If lRow2 = "" Then
			lRow2 = .vspdData2.MaxRows
		End If
			
		.vspdData2.Redraw = False	
		Select Case Index
			Case 0			
			Case 1
				ggoSpread.SpreadLock 1, lRow, .vspdData2.MaxCols, lRow2		
		End Select		
		.vspdData2.Redraw = True
		    
    End With
End Sub

'========================================================================================
Sub SetSpreadColor(Byval stsFg, Byval Index, ByVal lRow, ByVal lRow2)
    With frm1

		If  lRow2 = "" Then	lRow2 = lRow
		
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_ItemSeq, lRow, lRow2   ' 
		ggoSpread.SSSetProtected C_deptNm,    lRow, lRow2
		ggoSpread.SSSetProtected C_AcctNm, lRow, lRow2   ' �����ڵ��		
		ggoSpread.SSSetRequired  C_deptcd,    lRow, lRow2	   ' �μ��ڵ� 
						
		Select Case stsFg
		Case "I"				
			ggoSpread.SSSetRequired  C_deptcd,    lRow, lRow2	   ' �μ��ڵ� 
			ggoSpread.SSSetRequired C_AcctCd, lRow, lRow2	' �����ڵ� 
		CASE "Q"			
			ggoSpread.SSSetProtected  C_deptcd,    lRow, lRow2	   ' �μ��ڵ� 
			ggoSpread.SSSetProtected C_AcctCd, lRow, lRow2	' �����ڵ�			
		CASE "D"			
			ggoSpread.SSSetRequired  C_deptcd,    lRow, lRow2	   ' �μ��ڵ� 
			ggoSpread.SSSetProtected C_AcctCd, lRow, lRow2	' �����ڵ�							
		End Select	
		
		If  frm1.cboGlType.Value <> "03" Then
			ggoSpread.SSSetProtected C_DrCrNm, lRow, lRow2	' ���뱸�� 
'			ggoSpread.SSSetProtected  C_DocCur, lRow, lRow2	   ' �μ��ڵ� 
		ELSE
			ggoSpread.SSSetRequired C_DrCrNm, lRow, lRow2	' ���뱸�� 
		End If

		ggoSpread.SSSetRequired  C_DocCur, lRow, lRow2	   ' �μ��ڵ�				
		ggoSpread.SSSetRequired C_ItemAmt, lRow, lRow2	' �ݾ� 

		.vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub InitComboBox()
	
	Err.clear
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1013", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlType ,lgF0  ,lgF1  ,Chr(11))
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboGlInputType ,lgF0  ,lgF1  ,Chr(11))
	 
End Sub
'=======================================================================================================
Function InitComboBoxGrid()
    ggoSpread.Source = frm1.vspdData
	
	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1012", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1
	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_DrCrFg
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_DrCrNm
    
    Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("B9001", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    lgF0 = "" & chr(11) & lgF0
	lgF1 = "" & chr(11) & lgF1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_VatType
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_VatNm
	
End Function

'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet				'���Ѱ��� �߰�   							  
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	'//frm1.hOrgChangeId.Value = parent.gChangeOrgId

	Select Case iWhere
		Case 0
		
		Case 1
			If frm1.txtDeptCd.readOnly = true Then
				IsOpenPop = False
				Exit Function
			End If
			
			arrStrRet =  AutorityMakeSql("DEPT",frm1.hORGCHANGEID.Value, "","","","")	'���Ѱ��� �߰�   							  
			
			arrParam(0) = "�μ� �˾�"				' �˾� ��Ī 
			arrParam(1) = arrstrRet(0)										'���Ѱ��� �߰�   							  
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = arrstrRet(1)										'���Ѱ��� �߰�   							  
			arrParam(5) = "�μ��ڵ�"					' �����ʵ��� �� ��Ī 

			arrField(0) = "DEPT_CD"	     				' Field��(0)
			arrField(1) = "DEPT_NM"			    		' Field��(1)
    
			arrHeader(0) = "�μ��ڵ�"					' Header��(0)
			arrHeader(1) = "�μ���"				' Header��(1)
			
		Case 2
			arrParam(0) = "��ȭ�ڵ� �˾�"
			arrParam(1) = "B_Currency"
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "��ȭ�ڵ�"

			arrField(0) = "Currency"	
			arrField(1) = "Currency_desc"	
    
			arrHeader(0) = "��ȭ�ڵ�"
			arrHeader(1) = "��ȭ�ڵ��"		
			
		Case 3
			arrParam(0) = "�����ڵ��˾�"
			arrParam(1) = "A_Acct, A_ACCT_GP" 	
			arrParam(2) = strCode	
			arrParam(3) = ""	
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & " "	
			arrParam(5) = "�����ڵ�"	

			arrField(0) = "A_ACCT.Acct_CD"	
			arrField(1) = "A_ACCT.Acct_NM"	
    		arrField(2) = "A_ACCT_GP.GP_CD"	
			arrField(3) = "A_ACCT_GP.GP_NM"	
			
			arrHeader(0) = "�����ڵ�"	
			arrHeader(1) = "�����ڵ��"	
			arrHeader(2) = "�׷��ڵ�"	
			arrHeader(3) = "�׷��"	

		Case 4
			arrStrRet =  AutorityMakeSql("DEPT_ITEM",frm1.hORGCHANGEID.Value, frm1.txtDeptCd.Value,"","","")'���Ѱ��� �߰� 
			
			arrParam(0) = "�μ� �˾�"							' �˾� ��Ī 
			arrParam(1) = arrstrRet(0)								'���Ѱ��� �߰� 
			arrParam(2) = strCode									' Code Condition
			arrParam(3) = ""										' Name Cindition
			arrParam(4) = arrstrRet(1)								'���Ѱ��� �߰�									   
																	' Where Condition
			arrParam(5) = "�μ��ڵ�"							' �����ʵ��� �� ��Ī 

			arrField(0) = "A.DEPT_CD"	     									' Field��(0)
			arrField(1) = "A.DEPT_NM"			    							' Field��(1)
    
			arrHeader(0) = "�μ��ڵ�"									' Header��(0)
			arrHeader(1) = "�μ���"										' Header��(1)

	End Select
    
   	If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function
'========================================================================================================= 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
	
		Select Case iWhere
			Case 0
				.txtGlNo.Value = UCase(Trim(arrRet(0)))
				
			Case 1
				.txtDeptCd.Value = UCase(Trim(arrRet(0)))
				.txtDeptNm.Value = arrRet(1)
				Call txtDeptCd_OnChange()
			Case 2
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
				.vspdData.Col  = C_ItemLocAmt
				.vspdData.Text = ""
				.vspdData.Col  = C_DocCur 
				.vspdData.Text = UCase(Trim(arrRet(0)))
				If Trim(.vspdData.Text) = parent.gCurrency Then
					.vspdData.Col  = C_ExchRate
					.vspdData.Text = 1
				Else
					call FindExchRate(UniConvDateToYYYYMMDD(frm1.txtGLDt.text,parent.gDateFormat,""), UCase(Trim(arrRet(0))),frm1.vspdData.ActiveRow)
				End If
				
				call DocCur_OnChange(frm1.vspdData.ActiveRow,frm1.vspdData.ActiveRow)
			Case 3
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 
				    
				.vspdData.Col  = C_AcctCD
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_AcctNm
				.vspdData.Text = arrRet(1)
                Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)
			Case 4
				frm1.vspdData.Row = frm1.vspdData.ActiveRow 
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
				
				.vspdData.Col  = C_deptcd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_deptnm
				.vspdData.Text = arrRet(1)				
		End Select

	End With	
End Function
'========================================================================================================= 

Function OpenRefGL()
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(8)	                           '���Ѱ��� �߰� (3 -> 4)
	
	If IsOpenPop = True Then Exit Function
	iCalledAspName = AskPRAspName("a5104ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5104ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	IsOpenPop = True
	Call CookiePage("GL_POPUP")
	arrParam(4)	= lgAuthorityFlag 

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = ""  Then			
		Exit Function
	Else		
		Call SetRefGL(arrRet)
	End If
	
End Function
'========================================================================================================= 

Function SetRefGL(Byval arrRet)
	Dim intRtnCnt, strData
	Dim TempRow, I
	Dim j	
	
	With frm1
		.txtGlNo.Value = UCase(Trim(arrRet(0)))
    End With    
   
	frm1.txtGLNo.focus 
End Function

'========================================================================================
' Function Name : FncBtnCalc
' Function Desc : This function calculate local amt from amt of multi
'========================================================================================
Function FncBtnCalc() 

	Dim ii
	Dim tempAmt, tempLocAmt, tempExch, TempSep, tempDoc
	Dim strSelect, strFrom, strWhere
	Dim arrTemp
	Dim strDate
	Dim strExchFg
	Dim IntRetCD
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6		

	With frm1

		strSelect	= "b.minor_cd"
		strFrom		= "b_company a, b_minor b"
		strWhere	= "b.major_cd = " & FilterVar("a1004", "''", "S") & "  and	a.xch_rate_fg = b.minor_cd"
		If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
			arrTemp = Split(lgF0, chr(11))
			strExchFg =  arrTemp(0)
		End If	

		strDate = UniConvDateToYYYYMMDD(frm1.txtGLDt.text,parent.gDateFormat,"")
		If .vspdData.MaxRows <> 0 Then
			For ii = 1 To .vspdData.MaxRows
				.vspdData.Row	=	ii
				.vspdData.Col	=	C_DocCur			
				tempDoc			=	UCase(Trim(.vspdData.text))
				.vspdData.Col	=	C_ItemAmt
				tempAmt			=	UNICDbl(.vspdData.text)
				.vspdData.Col	=	C_ExchRate
				tempExch		=	UNICDbl(.vspdData.text)

				If tempDoc	<> "" and tempDoc <> parent.gCurrency Then
					If UCase(strExchFg) <> "D" Then 	' Fixed Exchange Rate
						strDate = Mid(strDate, 1, 6)
						strSelect	= "multi_divide"
						strFrom		= "b_monthly_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_yrmnth  =  " & FilterVar(strDate , "''", "S") & ""

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121600", "X", "X", "X")
						End If
					Else					' Floating Exchange Rate
						strSelect	= "top 1 multi_divide"
						strFrom		= "b_daily_exchange_rate (noLock) "
						strWhere	= "from_currency =  " & FilterVar(tempDoc , "''", "S") & ""
						strWhere	= strWhere & " And to_currency   =  " & FilterVar(parent.gCurrency , "''", "S") & ""
						strWhere	= strWhere & " And apprl_dt  <= convert(char(21), " & FilterVar(strDate, "''", "S") & ", 20) order by apprl_dt"

						If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then				
							arrTemp = Split(lgF0, chr(11))
							TempSep =  arrTemp(0)
						Else
							IntRetCD = DisplayMsgBox("121500", "X", "X", "X")
						End If
					End If
					If RTrim(LTrim(TempSep)) <> "/" Then
						tempLocAmt		=	tempAmt * TempExch
					Else
						tempLocAmt		=	tempAmt / TempExch
					End If
					.vspdData.Col	=	C_ItemLocAmt
					.vspdData.text	=	UNIConvNumPCToCompanyByCurrency(tempLocAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")

				ElseIf tempDoc = parent.gCurrency Then
					.vspdData.Col	=	C_ItemLocAmt
					.vspdData.text	=	UNIConvNumPCToCompanyByCurrency(tempAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
				End If
			Next		
		End If
	End With
	Call SetSumItem	
End Function

'========================================================================================
Function ExchRateCheck()
	Call FncBtnCalc()
End Function 

'========================================================================================
Function gfRealRound(ByVal x, ByVal Factor )
    Dim lcSwitch, iCurResult
    If x < 0 Then lcSwitch = -1 Else lcSwitch = 1
    x = x * lcSwitch
    iCurResult = Int(x * 10 ^ Factor + 0.5) / 10 ^ Factor
    gfRealRound = iCurResult * lcSwitch
End Function
'========================================================================================================= 
Function OpEndept(Byval strCode, Byval iWhere)
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(8)
	
	If IsOpenPop = True Then Exit Function
	If frm1.txtDeptCd.readOnly = true Then
		IsOpenPop = False
		Exit Function
	End If
	iCalledAspName = AskPRAspName("DeptPopupDtA2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "DeptPopupDtA2", "X")
		IsOpenPop = False
		Exit Function
	End If
	IsOpenPop = True

	arrParam(0) = strCode		            '  Code Condition
   	arrParam(1) = frm1.txtGLDt.Text
	arrParam(2) = lgUsrIntCd								' �ڷ���� Condition  
	If lgIntFlgMode = parent.OPMD_UMODE Then
		arrParam(3) = "T"									' �������� ���� Condition  
	Else
		arrParam(3) = "F"									' �������� ���� Condition  
	End If

	' ���Ѱ��� �߰� 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else		
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'========================================================================================================= 
Function OpenUnderDept(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim field_fg   	

	IsOpenPop = True


	If RTrim(LTrim(frm1.txtDeptCd.Value)) <> "" 	Then
		arrParam(0) = "�μ� �˾�"	
		arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B"				
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.Value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND B.BIZ_AREA_CD = ( SELECT B.BIZ_AREA_CD"
		arrParam(4) = arrParam(4) & " FROM B_ACCT_DEPT A, B_COST_CENTER B WHERE A.DEPT_CD =  " & FilterVar(frm1.txtDeptCd.Value , "''", "S") & ""
		arrParam(4) = arrParam(4) & " AND A.COST_CD = B.COST_CD AND A.ORG_CHANGE_ID =  " & FilterVar(frm1.hOrgChangeId.Value , "''", "S") & ")"
		arrParam(5) = "�μ��ڵ�"			
		arrField(0) = "A.DEPT_CD"	
		arrField(1) = "A.DEPT_Nm"
		arrField(2) = "B.BIZ_AREA_CD"
		arrHeader(0) = "�μ��ڵ�"		
		arrHeader(1) = "�μ��ڵ��"
		arrHeader(2) = "������ڵ�"				
	Else
		arrParam(0) = "�μ� �˾�"	
		arrParam(1) = "B_ACCT_DEPT A"				
		arrParam(2) = Trim(strCode)
		arrParam(3) = "" 
		arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
		arrParam(4) = arrParam(4) & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		arrParam(5) = "�μ��ڵ�"			
		arrField(0) = "A.DEPT_CD"	
		arrField(1) = "A.DEPT_Nm"
		arrHeader(0) = "�μ��ڵ�"		
		arrHeader(1) = "�μ��ڵ��"
	End If



	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
End Function
'========================================================================================================= 
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtDeptCd.Value = arrRet(0)
               .txtDeptNm.Value = arrRet(1)
               .txtInternalCd.Value = arrRet(2)
				.txtGLDt.text = arrRet(3)

				call txtDeptCd_OnChange()  
				
             Case "1"  

				frm1.vspdData.Row = frm1.vspdData.ActiveRow 
				
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow frm1.vspdData.ActiveRow 
				
				.vspdData.Col  = C_deptcd
				.vspdData.Text = arrRet(0)
				.vspdData.Col  = C_deptnm
				.vspdData.Text = arrRet(1)
				
				Call deptCd_underChange(arrRet(0))
				
             Case Else
         '      .vspdData.Col = C_Dept_cd                         'spread
         '      .vspdData.Text = arrRet(1)
        End Select
	End With
End Function       


'======================================================================================================
' Function Name : SetSumItem
' Function Desc :
'=======================================================================================================
Function SetSumItem()
    Dim DblTotDrAmt 
    Dim DblTotLocDrAmt 
    Dim DblTotCrAmt 
    Dim DblTotLocCrAmt 
        
    Dim lngRows 

	ggoSpread.Source = frm1.vspdData
	
    With frm1.vspdData 
		If .MaxRows > 0 Then
	        For lngRows = 1 To .MaxRows
	            .Row = lngRows
                    .Col = 0
                If .text <> ggoSpread.DeleteFlag Then
		            .col = C_DrCrFg
			    
		            If .text = "DR" Then		
			            .Col = C_ItemAmt	'6
			            If .Text = "" Then
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + 0
			            Else
			                DblTotDrAmt = UNICDbl(DblTotDrAmt) + UNICDbl(.Text)
			            End If
			            
			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + 0
			            Else
			                DblTotLocDrAmt = UNICDbl(DblTotLocDrAmt) + UNICDbl(.Text)
			            End If
		            ElseIf .text = "CR" Then
			            .Col = C_ItemAmt	'6
			            If .Text = "" Then
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + 0
			            Else
			                DblTotCrAmt = UNICDbl(DblTotCrAmt) + UNICDbl(.Text)
			            End If
			            
			            .Col = C_ItemLocAmt	'7
			            If .Text = "" Then
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + 0
			            Else
			                DblTotLocCrAmt = UNICDbl(DblTotLocCrAmt) + UNICDbl(.Text)
			            End If
					End If	
				End If	            
	        Next 
       End If                
        
'       If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY = " & FilterVar(UCase(Trim(frm1.txtDocCur.Value)),"''","S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
'			frm1.txtDrAmt.Text	= UNIConvNumPCToCompanyByCurrency(DblTotDrAmt,frm1.txtDocCur.Value,parent.ggAmtOfMoneyNo, "X", "X")
'			frm1.txtCrAmt.Text	= UNIConvNumPCToCompanyByCurrency(DblTotCrAmt,frm1.txtDocCur.Value,parent.ggAmtOfMoneyNo, "X", "X")
'		End If	
        frm1.txtDrLocAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotLocDrAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
        frm1.txtCrLocAmt.Text = UNIConvNumPCToCompanyByCurrency(DblTotLocCrAmt,parent.gCurrency,parent.ggAmtOfMoneyNo, parent.gLocRndPolicyNo, "X")
	End With

'    If frm1.cboGlType.value = "01" Then
'		frm1.txtDrLocAmt.text = frm1.txtCrLocAmt.text
'			frm1.txtDrAmt.text = frm1.txtCrAmt.text		
'	ElseIf frm1.cboGlType.value = "02" Then
'		frm1.txtCrLocAmt.text = frm1.txtDrLocAmt.text
'			frm1.txtCrAmt.text = frm1.txtDrAmt.text
'	End If
	
End Function

'========================================================================================================
'	Desc : Cookie Setting
'========================================================================================================
Function CookiePage(ByVal Kubun)


End Function

'========================================================================================================
'	Desc : ȭ���̵� 
'========================================================================================================
Function PgmJumpChk(strPgmId)
	Dim IntRetCD
	'-----------------------
	'Check previous data area
	'------------------------ 
	ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("990027", "X", "X", "X")                          'No data changed!!
        Exit Function
    End If  
	
	If Trim(frm1.txtGlNo.value) <> "" Then

		WriteCookie "GL_NO", Trim(frm1.txtGlNo.value)
	Else

		WriteCookie "GL_NO", Trim(frm1.htxtGlNo.value)
	End If
	
	'  Call CookiePage(1)
	Call PgmJump(strPgmId)
End Function
'========================================================================================================
'	Desc : ����� ȭ�鿡 ���� Grid�� Protect��ȯ 
'========================================================================================================
Sub CboGLType_ProtectGrid(Byval GlType)
	ggoSpread.Source = frm1.vspdData
	Select Case GlType		
		case "01"			
'			ggoSpread.SSSetProtected C_DocCur, 1, frm1.vspddata.maxrows	' ���뱸�� 
'			ggoSpread.SSSetProtected C_DocCurPopup, 1, frm1.vspddata.maxrows	' ���뱸�� 
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows	' ���뱸�� 
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows	' ���뱸�� 
		Case "02"			
'			ggoSpread.SSSetProtected C_DocCur, 1, frm1.vspddata.maxrows	' ���뱸�� 
'			ggoSpread.SSSetProtected C_DocCurPopup, 1, frm1.vspddata.maxrows	' ���뱸�� 
			ggoSpread.SSSetProtected C_DrCrfg, 1, frm1.vspddata.maxrows	' ���뱸�� 
			ggoSpread.SSSetProtected C_DrCrNm, 1, frm1.vspddata.maxrows	' ���뱸�� 
		Case "03"			
			ggoSpread.SSSetRequired C_DocCur, 1, frm1.vspddata.maxrows	' ���뱸�� 
			ggoSpread.SpreadUnLock C_DocCurPopup, 1, frm1.vspddata.maxrows	' ���뱸�� 
			ggoSpread.SpreadUnLock C_DrCrfg, 1, C_DrCrNm, frm1.vspddata.maxrows
			ggoSpread.SSSetRequired C_DrCrfg, 1, frm1.vspddata.maxrows	' ���뱸�� 
			ggoSpread.SSSetRequired C_DrCrNm, 1, frm1.vspddata.maxrows	' ���뱸�� 
	End Select 				
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
			 C_deptcd			= iCurColumnPos(2)
			 C_deptPopup		= iCurColumnPos(3)
			 C_deptnm	   		= iCurColumnPos(4)
			 C_AcctCd			= iCurColumnPos(5)
			 C_AcctPopup		= iCurColumnPos(6)
			 C_AcctNm			= iCurColumnPos(7)
			 C_DrCrFg			= iCurColumnPos(8)
			 C_DrCrNm			= iCurColumnPos(9)
			 C_DocCur			= iCurColumnPos(10)
			 C_DocCurPopup		= iCurColumnPos(11)
			 C_ExchRate			= iCurColumnPos(12)
			 C_ItemAmt			= iCurColumnPos(13)
			 C_ItemLocAmt		= iCurColumnPos(14)
			 C_IsLAmtChange		= iCurColumnPos(15)
			 C_ItemDesc			= iCurColumnPos(16)
			 C_VatType			= iCurColumnPos(17)
			 C_VatNm			= iCurColumnPos(18)
			 C_AcctCd2			= iCurColumnPos(19)
    End Select    
End Sub

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029    
    Call ggoOper.LockField(Document, "N")
		
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call InitSpreadSheet    
    Call InitCtrlSpread()
    Call InitCtrlHSpread()
    Call InitComboBox
    Call InitComboBoxGrid           
    Call SetAuthorityFlag   
    Call SetToolbar(MENU_NEW)
    Call SetDefaultVal
	Call InitVariables       
	Call CookiePage("FORM_LOAD")	

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

'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'=======================================================================================================
Sub vspdData_onfocus()
	lgCurrRow = frm1.vspdData.ActiveRow
	If frm1.vspdData.MaxRows > 0 Then
		Call SetToolbar(MENU_CRT)		        
	End If 

End Sub


'=======================================================================================================
Sub txtGLDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtGLDt.Action = 7
    End If
End Sub


'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Dim i
    Dim tmpDrCrFG
        
    Call SetPopUpMenuItemInf("1101111111")
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col,lgSortKey
            lgSortKey = 1
        End If    
        Exit Sub

    End If
	
	ggoSpread.Source = frm1.vspdData
	frm1.vspddata.row = frm1.vspddata.ActiveRow	

 	frm1.vspdData.Col = C_AcctCd
	
    If Len(frm1.vspdData.Text) < 1 Then
        'frm1.vspdData2.MaxRows = 0
        ggoSpread.Source = frm1.vspdData2
        ggoSpread.ClearSpreadData
	End If
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, ShIft, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)

    If Row <> NewRow And NewRow > 0 Then
	
        With frm1        
            .vspdData.Row = NewRow
            .vspdData.Col = C_ItemSeq
            
            .hItemSeq.value = .vspdData.Text
        ggoSpread.Source = frm1.vspdData2
        ggoSpread.ClearSpreadData
        End With

		frm1.vspddata.Col = 0
        If frm1.vspddata.Text = ggoSpread.DeleteFlag Then      
			Exit Sub
		End If

		lgCurrRow = NewRow
       
'        If CopyFromData(frm1.hItemSeq.value)  = true Then
'			SetSpread2Color
'		End If	

        Call DbQuery2(lgCurrRow)
    End If

    Call SetSumItem
End Sub

'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
Dim Ifld1 
Dim Ifld2
Dim iTable
Dim istrCode

	'---------- Coding part -------------------------------------------------------------
	With frm1.vspdData
		If Row > 0 And Col = C_AcctPopUp Then
			.Col = Col - 1
			.Row = Row
									
			Call OpenPopUp(.Text, 3)
		End If
		
		If Row > 0 And Col = C_deptPopup Then
			.Col = Col - 1
			.Row = Row							
			Call OpenUnderDept(.Text, 1)
			'//Call OpenPopUp(.Text, 4 )
    	End If    	
		If Row > 0 And Col = C_DocCurPopup Then
			.Col = Col - 1
			.Row = Row
			Call OpenPopUp(.Text, 2)
		End If

	End With
End Sub


'=======================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	Dim tmpDrCrFG
	Dim IntRetCD
	Dim TempExchRate
	Dim TempAmt
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    
    Select Case Col
		 Case   C_DeptCd
			frm1.vspdData.Col = C_DeptCd
			Call DeptCd_underChange(frm1.vspdData.text)
	    Case   C_AcctCd
		    frm1.vspdData.Col = 0
			If  frm1.vspdData.Text = ggoSpread.InsertFlag Then
				frm1.vspdData.Col = C_ItemSeq
				frm1.hItemSeq.Value = frm1.vspdData.Text
				frm1.vspdData.Col = C_AcctCd			
				If Len(frm1.vspdData.Text) > 0 Then
					frm1.vspdData.Row = Row
					frm1.vspdData.Col = C_ItemSeq	 
					DeleteHsheet frm1.vspdData.Text
					
					frm1.vspdData.Col = C_DrCrFg
					tmpDrCrFG = frm1.vspdData.text
					frm1.vspdData.Col = C_AcctCd

					If AcctCheck(frm1.vspdData.text,frm1.cboGlType.value, tmpDrCrFG) = True Then					
						Call Dbquery3(Row)
						Call InputCtrlVal(Row)
					End If	
				Else
					frm1.vspdData.Col = C_AcctNm
					frm1.vspdData.Text = ""
				End If   
			End If
    	Case   C_DrCrFg
    		Call SetSumItem
    	Case   C_DrCrNm  
    		Call vspdData_ComboSelChange(Col,Row)
			Call SetSumItem	
    	Case   C_ItemAmt
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_ItemLocAmt
			frm1.vspdData.Text = ""
    		Call SetSumItem	
		Case   C_ItemLocAmt
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_IsLAmtChange
			frm1.vspdData.Text = "Y"
			Call SetSumItem	
		Case	C_ExchRate
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_DocCur
			If UCase(Trim(frm1.vspdData.Text)) = parent.gCurrency Then
				frm1.vspdData.Row = Row
				frm1.vspdData.Col = C_ExchRate
				frm1.vspdData.Text = 1
			End If
		Case	C_DocCur
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = C_ItemLocAmt
			frm1.vspdData.Text = ""		
			frm1.vspdData.Col = C_DocCur
			If UCase(Trim(frm1.vspdData.Text)) = parent.gCurrency Then
				frm1.vspdData.Col = C_ExchRate
				frm1.vspdData.Text = 1
			Else
				Call FindExchRate(UniConvDateToYYYYMMDD(frm1.txtGLDt.text,parent.gDateFormat,""), UCase(Trim(frm1.vspdData.Text)),frm1.vspdData.ActiveRow)				
			End If
			
			Call DocCur_OnChange(Row,Row)			
    End Select	
    
	lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	Dim tmpDrCrFg
	Dim ii
	Dim iChkAcctForVat
	
	
	With frm1
		.vspddata.Row = Row

		Select Case Col
			Case C_DrCrNm

				.vspddata.Col = Col
				
				intIndex = .vspddata.Value
				.vspddata.Col = C_DrCrFg
				.vspddata.Value = intIndex				
				'SetSpread2Color 
				tmpDrCrFg = .vspddata.text			
				
				.vspddata.Col = C_AcctCd
				If AcctCheck(frm1.vspdData.text,frm1.cboGlType.value, tmpDrCrFg) = True Then					
					Call SetSpread2Color 					
				End If	
				
			Case C_VatNm
						
				.vspddata.Col = Col		       			
			    intIndex = .vspddata.Value
				.vspddata.Col = C_VatType				
				.vspddata.Value = intIndex		
				Call InputCtrlVal(Row)'
		End Select
		
	End With

End Sub
'==========================================================================================
Sub txtGlNo_OnKeyPress()	
	If window.event.keycode = 39 Then	'Single quotation mark �ԷºҰ� 
		window.event.keycode = 0	
	End If
End Sub

'==========================================================================================
Sub txtGlNo_OnKeyUp()	
	If Instr(1,frm1.txtGlNo.Value,"'") > 0 Then
		frm1.txtGlNo.Value = Replace(frm1.txtGlNo.Value, "'", "")		
	End If
End Sub

'==========================================================================================
Sub txtGlNo_onpaste()	
	Dim iStrGlNo 	
	iStrGlNo = window.clipboardData.getData("Text")
	iStrGlNo = RePlace(iStrGlNo, "'", "")
	Call window.clipboardData.setData("text",iStrGlNo)		
End Sub

'==========================================================================================
Sub DocCur_OnChange(FromRow, ToRow)

	Dim ii
    lgBlnFlgChgValue = True

	For ii = FromRow	to	ToRow
		frm1.vspdData.Row	= ii
		frm1.vspdData.Col	= C_DocCur

		If CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY =  " & FilterVar(frm1.vspdData.Text, "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
			Call CurFormatNumSprSheet(ii)
			Call SetSumItem
		End If	  
	Next  
End Sub

'==========================================================================================

Sub txtDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtGLDt.Text) = "" Or Trim(frm1.txtDeptCd.value) = "" Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S") 
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
	
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.Value = ""
		frm1.txtDeptNm.Value = ""
		frm1.hOrgChangeId.Value = ""
	Else 
		
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
					
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.hOrgChangeId.Value = Trim(arrVal2(2))
		Next	
			
	End If
	

End Sub
'==========================================================================================

Sub QueryDeptCd_OnChange()
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtGLDt.Text) = "" Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S") 
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
	
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
'			IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.Value = ""
		frm1.txtDeptNm.Value = ""
		frm1.hOrgChangeId.Value = ""
	Else 
		
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)
					
		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.hOrgChangeId.Value = Trim(arrVal2(2))
		Next	
			
	End If
	
End Sub

'==========================================================================================
Sub DeptCd_underChange(Byval strCode)
        
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 

    If Trim(frm1.txtGLDt.Text) = "" Or Trim(strCode) = ""  Then    
		Exit sub
    End If
    lgBlnFlgChgValue = True
	'----------------------------------------------------------------------------------------
	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(strCode)), "''", "S") 
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
	
		IntRetCD = DisplayMsgBox("124600","X","X","X")  

		frm1.vspdData.Col = C_deptcd			
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.text = ""
		frm1.vspdData.Col = C_deptnm		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow	
		frm1.vspdData.text = ""
	
	End If 
	
End Sub


'==========================================================================================
Sub txtGLDt_Change()

	If lgstartfnc = False Then
	    If lgFormLoad = True Then
			Dim strSelect
			Dim strFrom
			Dim strWhere 	
			Dim IntRetCD 
			Dim ii
			Dim arrVal1
			Dim arrVal2
			Dim jj


			lgBlnFlgChgValue = True
			With frm1
			
			If Trim(.txtDeptCd.Value) <> "" and Trim(.txtGLDt.Text ) <> "" Then
				'----------------------------------------------------------------------------------------
				strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
				strFrom		=			 " b_acct_dept(NOLOCK) "		
				strWhere	=			 " dept_Cd = " & FilterVar(LTrim(RTrim(.txtDeptCd.Value)), "''", "S") 
				strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
				strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
				strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
					If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
						IntRetCD = DisplayMsgBox("124600","X","X","X")
						.txtDeptCd.Value = ""
						.txtDeptNm.Value = ""
						.hOrgChangeId.Value = ""
						If .vspdData.MaxRows <> 0 Then
							For ii = 1 To .vspdData.MaxRows
							.vspdData.Col = C_deptcd			
						    .vspdData.Row = ii
						    .vspdData.text = ""
						    .vspdData.Col = C_deptnm	
						    .vspdData.text = ""
							Next		
						End If
					Else
						arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
						jj = Ubound(arrVal1,1)
								
						For ii = 0 to jj - 1
							arrVal2 = Split(arrVal1(ii), chr(11))			
							frm1.hOrgChangeId.Value = Trim(arrVal2(2))
						Next	
					End If 
				End If
			End With
		End If
	End If
End Sub

'==========================================================================================
Sub cboGLType_OnChange()
	
	Dim	i		
	Dim IntRetCD	
	
	ggoSpread.Source = frm1.vspdData
	
	SELECT CASE frm1.cboGlType.Value 
		CASE "01"			
			'�Ա���ǥ�� �ٲٸ� ������ �Էµǰų� ���ݰ����� �ԷµǾ����� check�Ѵ�.
			FOR i = 1 TO  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				If  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.Value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.Value )		
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")					
					Exit sub
				End If
																			
				frm1.vspddata.col = C_DrCrFg
				If  Trim(frm1.vspddata.Value) = "2" Then
					frm1.cboGlType.Value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.Value )		
					IntRetCD = DisplayMsgBox("113104", "X", "X", "X")					
					Exit sub
				End If	
			Next				
			
			FOR i = 1 TO  frm1.vspdData.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				If Trim(frm1.vspddata.Value) <> "1"  Then					
					frm1.vspdData.Value	= "1"							
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.Value	= "1"							
				End If
				frm1.vspddata.col = C_DocCur
				frm1.vspddata.text = parent.gCurrency
			Next
			Call CboGLType_ProtectGrid(frm1.cboGlType.Value )		
		CASE "02"
			'�����ǥ�� �ٲٸ� �뺯�� �Էµǰų� ���ݰ����� �ԷµǾ����� check�Ѵ�.	
			FOR i = 1 TO  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_Acctcd
				If  frm1.vspddata.text = lgCashAcct Then
					frm1.cboGlType.Value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.Value )		
					IntRetCD = DisplayMsgBox("113106", "X", "X", "X")					
					Exit sub
				End If								
				
				frm1.vspddata.col = C_DrCrFg
				If  Trim(frm1.vspddata.Value) = "1" Then
					frm1.cboGlType.Value = "03"
					Call CboGLType_ProtectGrid(frm1.cboGlType.Value )		
					IntRetCD = DisplayMsgBox("113105", "X", "X", "X")					
					Exit sub				
				End If											
			Next
				
			FOR i = 1 TO  frm1.vspddata.maxrows
				frm1.vspddata.Row = i
				frm1.vspddata.col = C_DrCrFg
				If Trim(frm1.vspddata.Value) <> "2"  Then					
					frm1.vspdData.Value	= "2"							
					frm1.vspddata.col = C_DrCrNm
					frm1.vspdData.Value	= "2"							
				End If
				frm1.vspddata.col = C_DocCur
				frm1.vspddata.text = parent.gCurrency
			Next
			Call CboGLType_ProtectGrid(frm1.cboGlType.Value )		
		CASE "03"
		'��ü�� �ٲٸ� Protect�� Ǯ���ش�.		
			Call CboGLType_ProtectGrid(frm1.cboGlType.Value )		
		
	End SELECT	
	
	lgBlnFlgChgValue = True
End Sub


'========================================================================================

Function FncQuery() 
    Dim IntRetCD
    Dim RetFlag
    lgstartfnc = True
    FncQuery = False
    Err.Clear

    '-----------------------
    'Check previous data area
    '-----------------------
    	
    ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "1") Then	'��: This function check indispensable field
       Exit Function
    End If
    
    If lgBlnFlgChgValue = True  OR ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
      	Exit Function
     	End If
    End If
    
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")      '��: Condition field clear
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
    Call InitVariables		
    '-----------------------
    'Check condition area
    '-----------------------
    
    If frm1.txtDeptCd.Value = "" Then
		frm1.txtDeptNm.Value = ""
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then																'��: Query db data
		Exit Function
    End If
    
    If frm1.vspddata.maxrows = 0 Then	
       frm1.txtGlNo.Value = ""
    End If
   
    FncQuery = True																'��: Processing is OK
    lgstartfnc = False
End Function

'========================================================================================
Function FncNew() 
	Dim IntRetCD 
    
    Dim var1, var2
    
    lgstartfnc = True
    FncNew = False  
    Err.Clear       
    On Error Resume Next
    
    ggoSpread.Source = frm1.vspdData
    var1 = ggoSpread.SSCheckChange
    ggoSpread.Source = frm1.vspdData2
    var2 = ggoSpread.SSCheckChange

    '-----------------------
    'Check previous data area
    '-----------------------
    ' ����� ������ �ִ��� Ȯ���Ѵ�.
    If (lgBlnFlgChgValue = True Or var1 = True Or var2 = True) And lgBlnExecDelete <> True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgBlnExecDelete = False

'    Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
'    Call ggoOper.ClearField(Document, "2")                                  '��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData3
    ggoSpread.ClearSpreadData
'    Call ggoOper.LockField(Document,  "N")                                  '��: Lock  Suitable  Field

    SetGridFocus()
    SetGridFocus2()
    
	Call SetDefaultVal
    Call InitVariables                                                      '��: Initializes local global variables
	Call SetSumItem()
        
    Call SetToolbar(MENU_NEW)										'��ư ���� ���� 
	
	Call ggoOper.SetReqAttr(frm1.txtGlDt,   "N")
	Call ggoOper.SetReqAttr(frm1.txtdesc,	"D") 
    
    lgBlnFlgChgValue = False

    FncNew = True                              '��: Processing is OK
    lgFormLoad = True							' gldt read
    lgstartfnc = False
End Function


'========================================================================================

Function FncDelete() 
    
End Function


'=======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'=======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim lngItemCnt
    Dim lngRows

    FncSave = False                                                         
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                   '��: Protect system from crashing
    
	'-----------------------
    'Precheck area
    '----------------------- 
	
    If CheckSpread3 = False Then
	IntRetCD = DisplayMsgBox("110420", "X", "X", "X")                           '�ʼ��Է� check!!
        Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
	lngItemCnt = 0    
		For lngRows = 1 To .MaxRows    
			.Row = lngRows
			.Col = 0
			If .Text <> ggoSpread.DeleteFlag Then
				lngItemCnt = lngItemCnt + 1
			End If		
		Next
    End With
	
	If lngItemCnt < 1 Then															'������ǥ�׸� �����Ͱ� �������� �ʽ��ϴ�.
		IntRetCD = DisplayMsgBox("113200", "X", "X", "X")
		Exit Function
	End If
	

  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") Then                                   '��: Check contents area
       Exit Function
    End If
    
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData3
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    
	'-----------------------
    'Save function call area
    '----------------------- 
	'Call ExchRateCheck()
    If DbSave = False Then				                                                '��: Save db data
		Exit Function
    End If

    FncSave = True                                                          
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 

	Dim  IntRetCD	 
	frm1.vspdData.ReDraw = False	
	If frm1.vspdData.MaxRows < 1 Then Exit Function	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow    
    SetSpreadColor "I", 0, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    MaxSpreadVal frm1.vspdData, C_ItemSeq, frm1.vspdData.ActiveRow

      
	Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DocCur,C_ItemAmt,   "A" ,"I","X","X")
	Call vspdData_Change(C_AcctCd, frm1.vspddata.activerow)    
    Call SetSumItem()
End Function

'========================================================================================================
Function FncCancel() 
    Dim iItemSeq
    Dim RowDocCur

	If frm1.vspdData.MaxRows < 1 Then 	Exit Function	
	
	If  frm1.vspdData.MaxRows = 1 Then  Call ggoOper.SetReqAttr(frm1.cboGlType,   "N")
	
    With frm1.vspdData
        .Row = .ActiveRow
        .Col = 0
        If .Text = ggoSpread.InsertFlag Then
			.Col = C_AcctCd
			If len(Trim(.text)) > 0 Then 
				.Col = C_ItemSeq
				DeleteHSheet(.Text)
			End If	
        End If        

        ggoSpread.Source = frm1.vspdData	
        ggoSpread.EditUndo
        
        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Frm1.vspdData.ActiveRow,Frm1.vspdData.ActiveRow,C_DocCur,C_ItemAmt,   "A" ,"I","X","X")

		If .MaxRows = 0 Then
'			Call SetToolbar(MENU_NEW)
			Exit Function
		End If
	
        InitData
        
        .Row = .ActiveRow
        .Col = 0
		If .row = 0 Then 
			Exit Function
		End If

        If .Text = ggoSpread.InsertFlag Then
                       
            .Col = C_AcctCd
            If Len(.Text) > 0 Then
				.Col = C_ItemSeq
				frm1.hItemSeq.Value = .Text
	            frm1.vspdData2.MaxRows = 0
		        Call DbQuery3(.ActiveRow)
            End If
        Else
            .Col = C_ItemSeq
            frm1.hItemSeq.Value = .Text
            frm1.vspdData2.MaxRows = 0
            If CopyFromData(frm1.hItemSeq.value)  = true Then
				SetSpread2Color
			End If	           
'		    Call DbQuery2(.ActiveRow)            
        End If
        
    End With        
    Call SetSumItem()
    
End Function

'========================================================================================

Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim imRow2
	Dim iCurRowPos

    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False         
   
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else

        imRow = AskSpdSheetAddRowCount()
       If imRow = "" Then
            Exit Function
        End If
    End If
            
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
		iCurRowPos = .vspdData.ActiveRow
       ' ggoSpread.InsertRow ,imRow

        for imRow2 = 1 to imRow 
            ggoSpread.InsertRow ,1
            .vspdData.row = .vspdData.ActiveRow
           .vspdData.col = C_deptcd
            .vspddata.text	= UCase(.txtDeptCd.Value)
            
            .vspdData.col = C_deptnm
            .vspddata.text	= .txtDeptNm.Value

            .vspdData.col = C_DocCur
            .vspddata.text	= parent.gCurrency

            .vspdData.col = C_ExchRate
            .vspddata.text	= "1"
            
            .vspdData.col = C_ItemDesc
            .vspddata.text	= .txtDesc.Value
            If  frm1.cboGlType.value = "01" Then
                .vspdData.col = C_DrCrNm
                .vspdData.value	= 1					
                .vspdData.col = C_DrCrFg
                .vspdData.value	= 1					
            ELSEIf frm1.cboGlType.value = "02" Then		
                .vspdData.col = C_DrCrNm
                .vspdData.value	= 2				
                .vspdData.col = C_DrCrFg
                .vspdData.value	= 2			
            End If	
            SetSpreadColor "I", 0, frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
            MaxSpreadVal frm1.vspdData, C_ItemSeq, frm1.vspdData.ActiveRow

        Next
  
        Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,iCurRowPos + 1,iCurRowPos + imRow,C_DocCur,C_ItemAmt,"A" ,"I","X","X")
      
        .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================

Function FncDeleteRow() 
	Dim lDelRows
	Dim iDelRowCnt, i
    Dim DelItemSeq

	With frm1.vspdData 

    ggoSpread.Source = frm1.vspdData 

	.Row = .ActiveRow
	.Col = 0 
		
	If frm1.vspdData.MaxRows < 1 Or .Text = ggoSpread.InsertFlag Then Exit Function

        .Col = 1 
        DelItemSeq = .Text
    	
    	lDelRows = ggoSpread.DeleteRow
    
    End With
        
    DeleteHsheet DelItemSeq
    Call SetSumItem()
End Function


'========================================================================================

Function FncPrev() 
    On Error Resume Next  
End Function
'========================================================================================

Function FncNext() 
    On Error Resume Next   
End Function

'========================================================================================
Function FncPrint() 
    On Error Resume Next    
    
    parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
    On Error Resume Next  
    Call parent.FncExport(parent.C_MULTI)												'��: ȭ�� ���� 
End Function

'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                         '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim indx

	On Error Resume Next
	Err.Clear 		

	ggoSpread.Source = gActiveSpdSheet

    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call PrevspdDataRestore(gActiveSpdSheet)
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet()
            Call InitComboBoxGrid
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			
			Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,1,-1,C_DocCur,C_ItemAmt,   "A" ,"I","X","X")
			
            Call SetSpreadColor("Q", 0,1, .vspddata.MaxRows)                 
            Call SetSpread2Color()               

		Case "VSPDDATA2"
			Call PrevspdData2Restore(gActiveSpdSheet)						
			Call ggoSpread.RestoreSpreadInf()							
			Call InitCtrlSpread()			'�����׸� �׸��� �ʱ�ȭ			
			Call ggoSpread.ReOrderingSpreadData()			
			Call InitData()
            Call SetSpread2Color()
	End Select
	
	If frm1.vspdData2.MaxRows <= 0 Then
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If
	
	Call SetSumItem()
End Sub


'=======================================================================================================
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
									Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtGLNo.Value)
								End If
						End Select
					Next
					'ggoSpread.Source = frm1.vspdData					
					'ggoSpread.EditUndo
					
				Case ggoSpread.DeleteFlag
					Call fncRestoreDbQuery2(indx, frm1.vspdData.ActiveRow, frm1.htxtGLNo.Value)
					'ggoSpread.Source = frm1.vspdData
					'ggoSpread.EditUndo
			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName
End Sub

'=======================================================================================================
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
							Call fncRestoreDbQuery2(indx1, frm1.vspdData.ActiveRow, frm1.htxtGLNo.Value)
						End If
					Next

				Case ggoSpread.DeleteFlag

			End Select
		End If
	Next
	
	ggoSpread.Source = pActiveSheetName

End Sub

'========================================================================================================
Function fncRestoreDbQuery2(Row, CurrRow, Byval pInvalue1)
	Dim strItemSeq
	Dim strSelect, strFrom, strWhere
	Dim arrTempRow, arrTempCol
	Dim Indx1
	Dim strTableid, strColid, strColNm, strMajorCd
	Dim strNmwhere
	Dim arrVal
	Dim strVal
'	Dim tmpDrCrFG

	On Error Resume Next
	Err.Clear

	fncRestoreDbQuery2 = False

	Call DisableToolBar(parent.TBC_QUERY)
	Call LayerShowHide(1)
	
	With frm1
		.vspdData.row = Row
	    .vspdData.col = C_ItemSeq
		strItemSeq    = .vspdData.Text
		
'	    .vspdData.Col = C_DrCrFg
'		frm1.vspdData.Col = C_DrCrFg
'		tmpDrCrFG = frm1.vspdData.text		
		
	    If Trim(strItemSeq) = "" Then
	        Exit Function
	    End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
'		strSelect = strSelect & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'		strSelect = strSelect & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  		
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " End	, " & strItemSeq & ", "
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

'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False

	ggoSpread.Source = frm1.vspdData    
    If lgBlnFlgChgValue = True OR ggoSpread.SSCheckChange = True Then  
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"	
		If IntRetCD = vbNo Then
			Exit Function
		End If		
    End If       
    
    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim RetFlag

    DbQuery = False

    Call LayerShowHide(1)
    frm1.vspdData3.MaxRows = 0 

    Err.Clear                '��: Protect system from crashing
    
    With frm1
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001 
		strVal = strVal & "&txtGlNo=" & UCase(Trim(.txtGlNo.Value))	'��ȸ ���� ����Ÿ 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		strVal = strVal & "&lgAuthorityFlag=" & lgAuthorityFlag

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
		
		Call RunMyBizASP(MyBizASP, strVal)		'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True
End Function

'=======================================================================================================
Function DbQueryOk()
	Dim ii

	With frm1
	   .vspdData.Col = 1
	   intItemCnt = .vspddata.MaxRows
		
		Call SetSpreadColor("D", 0, 1, intItemCnt)
        '-----------------------
        'Reset variables area
        '-----------------------   
        
        If lgIntFlgMode = parent.OPMD_CMODE Then
			frm1.txtGlNo.value = ""
			frm1.txtGLDt.text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
			Call SetToolbar(MENU_CRT)
		Else
			Call CancelRestoreToolBar()
			Call SetToolbar(MENU_NEW)			
			lgBlnFlgChgValue = False
		End If
        
        lgIntFlgMode = parent.OPMD_CMODE
        Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field
        .txtCommandMode.value = "CREATE"
		.cboGlInputType.value	=	C_GLINPUTTYPE
		
        InitData
        Call SetSumItem

        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = C_ItemSeq
            .hItemSeq.Value = .vspdData.Text 
            Call DbQueryAll(1)
            Call DbQuery2(1)
        End If
    End With
    
    lgIntFlgMode = parent.OPMD_UMODE
    Call txtDeptCd_OnChange()

	SetGridFocus()
    SetGridFocus2()
    lgBlnFlgChgValue = False
End Function

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
	Dim i
	Dim arrVal
	Dim arrTemp
	Dim Indx1		
'	Dim tmpDrCrFG	
	
	Err.Clear
	
	With frm1
	    .vspdData.row = Row
	    .vspdData.col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text

'	    .vspdData.Col = C_DrCrFg
'		frm1.vspdData.Col = C_DrCrFg
'		tmpDrCrFG = frm1.vspdData.text

	    If Trim(.hItemSeq.Value) = "" Then
	        Exit Function
	    End If
	    
	    frm1.vspdData2.ReDraw = false
	    
        If CopyFromData(.hItemSeq.Value) = True Then
			If frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
				Call SetSpread2Lock("",1,1,"")
			Else
				Call SetSpread2Color()
			End  If
			frm1.vspdData2.ReDraw = True
            Exit Function
        End If
    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
'		strSelect = strSelect & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'		strSelect = strSelect & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  		
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " End	, " & .hItemSeq.Value & ", "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')) "
    		
		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_GL_DTL C (NOLOCK), A_GL_ITEM D (NOLOCK) "
		
		strWhere =			  " D.GL_NO = " & FilterVar(.htxtGlNo.value, "''", "S") 	
		strWhere = strWhere & " And D.ITEM_SEQ = " & .hItemSeq.Value & " "				
		strWhere = strWhere & " AND D.GL_NO  =  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
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

					strNmwhere = strColid & " =  " & FilterVar(UCase(frm1.vspddata2.text), "''", "S")
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " And MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") 
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						frm1.vspddata2.col = C_CtrlValNm
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata2.text = arrVal(0)
					End If
				End If

				strVal = strVal & Chr(11) & .hItemSeq.Value

				.vspdData2.Col = C_DtlSeq
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlCd
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlNm
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlVal
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlPB
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_CtrlValNm
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_Seq
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_Tableid
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_Colid
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_ColNm
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_Datatype
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_DataLen
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_DRFg
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_MajorCd
				strVal = strVal & Chr(11) & .vspdData2.Text

				.vspdData2.Col = C_MajorCd + 1

				strVal = strVal & Chr(11) & lngRows

				strVal = strVal & Chr(11) & Chr(12)
			Next

			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSShowData strVal
		End If
						
		intItemCnt = .vspddata.MaxRows
                
		If frm1.cboGlInputType.Value <> C_GLINPUTTYPE Then
			Call SetSpread2Lock("",1,1,"")
		Else
			Call SetSpread2Color()
		End  If
	End With
	
	Call LayerShowHide(0)
	
	frm1.vspdData2.ReDraw = True
	
	DbQuery2 = True
End Function
'=======================================================================================================
'=======================================================================================================
' Function Name : DbQueryAll
' Function Desc : Copy�� �����׸��� ��� call�Ѵ�. 
'=======================================================================================================
Function DbQueryAll(ByVal Row)
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
	Dim i
	Dim arrVal
	Dim arrTemp
	Dim Indx1		
'	Dim tmpDrCrFG	
	
	Err.Clear
	
	With frm1
	    .vspdData.row = Row
	    .vspdData.col = C_ItemSeq
	    .hItemSeq.Value = .vspdData.Text
		.vspddata3.maxrows = 0
		
		Call LayerShowHide(1)
	
		DbQueryAll = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_ItemSeq
		
		strSelect =				" D.ITEM_SEQ, C.DTL_SEQ,  A.CTRL_CD, A.CTRL_NM , C.CTRL_VAL, '',"
		strSelect = strSelect & " CASE  WHEN A.COLM_DATA_TYPE = " & FilterVar("D", "''", "S") & "  Then " & FilterVar("(Format : YYYY-MM-DD)", "''", "S") & "  End , D.ITEM_SEQ, LTrim(ISNULL(A.TBL_ID,'')), LTrim(ISNULL(A.DATA_COLM_ID,'')), "
		strSelect = strSelect & " LTrim(ISNULL(A.DATA_COLM_NM,'')),  LTrim(ISNULL(A.COLM_DATA_TYPE,'')), LTrim(ISNULL(A.DATA_LEN,'')), "
		strSelect = strSelect & " CASE WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("DC", "''", "S") & "  "
'		strSelect = strSelect & " WHEN B.DR_FG = 'Y' AND 'DR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  
'		strSelect = strSelect & " WHEN B.CR_FG = 'Y' AND 'CR'='" & Trim(tmpDrCrFG) & "' THEN 'Y' "  		
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("Y", "''", "S") & "  AND  B.CR_FG = " & FilterVar("N", "''", "S") & "  Then " & FilterVar("D", "''", "S") & "  "
		strSelect = strSelect & " WHEN B.DR_FG = " & FilterVar("N", "''", "S") & "  AND  B.CR_FG = " & FilterVar("Y", "''", "S") & "  Then " & FilterVar("C", "''", "S") & "  "
		strSelect = strSelect & " End	,  "
		strSelect = strSelect & " LTrim(ISNULL(A.MAJOR_CD,'')), "
		strSelect = strSelect & " D.ITEM_SEQ  "  		

		strFrom = " A_CTRL_ITEM	A (NOLOCK), A_ACCT_CTRL_ASSN B (NOLOCK), A_GL_DTL C (NOLOCK), A_GL_ITEM D (NOLOCK) "
		
		strWhere =			  " D.GL_NO = " & FilterVar(.htxtGlNo.value, "''", "S") 	
'		strWhere = strWhere & " And D.ITEM_SEQ = " & .hItemSeq.Value & " "				
		strWhere = strWhere & " AND D.GL_NO  =  C.GL_NO  "
		strWhere = strWhere & " AND D.ITEM_SEQ  =  C.ITEM_SEQ "
		strWhere = strWhere & "	AND D.ACCT_CD *= B.ACCT_CD "
		strWhere = strWhere & " AND C.CTRL_CD *= B.CTRL_CD "		
		strWhere = strWhere & " AND C.CTRL_CD = A.CTRL_CD "
		strWhere = strWhere & " ORDER BY C.DTL_SEQ "
			
		frm1.vspdData3.ReDraw = False
			
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then   
			ggoSpread.Source = frm1.vspdData3
			arrTemp =  Split(lgF2By2,Chr(12))

			For Indx1 = 0 To Ubound(arrTemp) - 1
				arrTemp(indx1) = Replace(arrTemp(indx1), Chr(8), indx1 + 1)
			Next

			lgF2By2 = Join(arrTemp,Chr(12))
			ggoSpread.SSShowData lgF2By2

			For lngRows = 1 To frm1.vspdData3.Maxrows
				frm1.vspddata3.row = lngRows	
				frm1.vspddata3.col = C_Tableid +1
				
				If Trim(frm1.vspddata3.text) <> "" Then
					frm1.vspddata3.col = C_Tableid +1
					strTableid = frm1.vspddata3.text
					frm1.vspddata3.col = C_Colid +1
					strColid = frm1.vspddata3.text 
					frm1.vspddata3.col = C_ColNm + 1
					strColNm = frm1.vspddata3.text	
					frm1.vspddata3.col = C_MajorCd 
					strMajorCd = frm1.vspddata3.text

					frm1.vspddata3.col = C_CtrlVal + 1

					strNmwhere = strColid & " =  " & FilterVar(UCase(frm1.vspddata3.text), "''", "S")
					
					If Trim(strMajorCd) <> "" Then
						strNmwhere = strNmwhere & " And MAJOR_CD =  " & FilterVar(strMajorCd , "''", "S") 
					End If

					If CommonQueryRs( strColNm , strTableid ,  strNmwhere , lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
						frm1.vspddata3.col = C_CtrlValNm + 1
						arrVal = Split(lgF0, Chr(11))  
						frm1.vspddata3.text = arrVal(0)
					End If
				End If
				
			Next
			
		End If
						
		intItemCnt = .vspddata.MaxRows                
		
	End With
	
	Call LayerShowHide(0)
	
	frm1.vspdData3.ReDraw = True
	
	DbQueryAll = True
End Function
'=======================================================================================================

Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim intIndex2 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
					
			.Row = intRow			
			.Col = C_DrCrFg
			intIndex = .Value
			.col = C_DrCrNm
			.Value = intindex
									
			.Col = C_VatType
			intIndex2 = .Value
			.col = C_VatNm
			.Value = intIndex2		
		Next	
	End With
End Sub

'========================================================================================================
Function DbSave() 
    Dim pAP010M 
    Dim lngRows , itemRows
    Dim lGrpcnt
    Dim strVal 
    Dim tempItemSeq
    Dim	intRetCd	
    Dim strNote
    Dim strItemDesc

    DbSave = False                                                          
    Call LayerShowHide(1)

    On Error Resume Next                                                   
	Err.Clear 
    
    Call SetSumItem

	With frm1
		.txtFlgMode.Value = lgIntFlgMode									
		.txtUpdtUserId.Value = parent.gUsrID
		.txtInsrtUserId.Value  = parent.gUsrID
		.txtMode.Value = parent.UID_M0002
		.txtAuthorityFlag.Value     = lgAuthorityFlag               '���Ѱ��� �߰� 
		
		'//	.hOrgChangeId.Value = parent.gChangeOrgId
	End With

    '-----------------------
    'Data manipulate area
    '-----------------------
    ' Data ���� ��Ģ 
    ' 0: Sheet��, 1: Flag , 2: Row��ġ, 3~N: �� ����Ÿ 

    lGrpCnt = 1
    strVal = ""
    
    ggoSpread.Source = frm1.vspdData
    With frm1.vspdData
		For lngRows = 1 To .MaxRows
			.Row = lngRows
			.Col = 0

			If .Text <> ggoSpread.DeleteFlag Then
				strVal = strVal & "C" & parent.gColSep & lngRows & parent.gColSep				'C=Create, Sheet�� 2�� �̹Ƿ� ���� 
			        
			    .Col = C_ItemSeq	'1
			    strVal = strVal & Trim(.Text) & parent.gColSep
			            
			    .Col = C_deptcd	    '2
			    strVal = strVal & Trim(.Text) & parent.gColSep
			        
			    .Col = C_AcctCd		'3
			    strVal = strVal & Trim(.Text) & parent.gColSep

				.Col = C_DrCrFG		'4
			    strVal = strVal & Trim(.Text) & parent.gColSep
			        
			    .Col = C_ItemAmt		'5
			    strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep
			        
			    '.Col = C_IsLAmtChange
				
				'Local �ݾ��� ����� �Է½� �Է±ݾ��� ���� 
				'If .Text = "Y" Then
   					.Col = C_ItemLocAmt	'6
					strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep
				'Else
				'	strVal = strVal & "0" & parent.gColSep
				'End If		
			        
			    .Col = C_ItemDesc	'7
				strItemDesc = Trim(.Text)
			    
				If Trim(strItemDesc) = "" Or isnull(strItemDesc) Then
					 ggoSpread.Source = frm1.vspdData3
					 frm1.vspdData.Col = C_ItemSeq
					 tempItemSeq = frm1.vspdData.Text  
					 strNote = ""
					 With frm1.vspdData3
							For itemRows = 1 to frm1.vspdData3.MaxRows
								.Row = itemRows
								.Col = 1
									
								If .Text =  tempItemSeq Then 
									.Col= 9 'C_Tableid	+ 1				
									If 	.Text = "B_BIZ_PARTNER" OR .Text = "B_BANK" OR .Text = "F_DPST" Then
										.Col = 7 'C_CtrlValNm + 1 
									ELSE
										.Col = 5 'C_CtrlVal + 1 
									End If											
									strNote = strNote & C_NoteSep & Trim(.Text)
								End If		    
							Next
							strNote = Mid(strNote,2)
					 End With
						 
					 strVal = strVal & strNote & parent.gColSep
					 ggoSpread.Source = frm1.vspdData
				Else
					strVal = strVal & strItemDesc & parent.gColSep		'8
				End If	
			    
				.Col = C_ExchRate	'9
			    strVal = strVal & UNICDbl(Trim(.Text)) & parent.gColSep
			        
			    .Col = C_VatType	'10
			    strVal = strVal & Trim(.Text) & parent.gColSep

			    .Col = C_DocCur		'11
			    strVal = strVal & UCase(Trim(.Text)) & parent.gRowSep

			    lGrpCnt = lGrpCnt + 1
			End If
		Next
    End With
	
    frm1.txtMaxRows.Value = lGrpCnt-1										'Spread Sheet�� ����� �ִ밹�� 
    frm1.txtSpread.Value =  strVal									'Spread Sheet ������ ���� 

	If frm1.txtSpread.Value = "" Then	
		intRetCd = DisplayMsgBox("990008", parent.VB_YES_NO, "X", "X")   '�� �ٲ�κ� 
		If intRetCd = VBNO Then
			Exit Function
		End If	
		If  DbDelete = False Then
			Exit Function
		End If			
		ggoSpread.Source = frm1.vspdData
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
		ggoSpread.Source = frm1.vspdData3
		ggoSpread.ClearSpreadData
		Call InitVariables	    
		Exit Function
	End If
	
    lGrpCnt = 1
    strVal = ""
    
    ggoSpread.Source = frm1.vspdData3

    With frm1.vspdData3      ' Dtl ���� 
		For itemRows = 1 To frm1.vspdData.MaxRows 
 		    frm1.vspdData.Row = itemRows
		    frm1.vspdData.Col = 0

			If frm1.vspdData.Text <> ggoSpread.DeleteFlag Then	
		        frm1.vspdData.Col = C_ItemSeq
			    tempItemSeq = frm1.vspdData.Text  

			    For lngRows = 1 To .MaxRows
					.Row = lngRows
					.Col = 1
					
					If .text = tempitemseq Then
		                .Col = 0 

						strVal = strVal & "C" & parent.gColSep
								
						.Col = 1 		 					'ItemSEQ	
						strVal = strVal & tempitemseq & parent.gColSep
						            
						.Col =  2 'C_DtlSeq + 1   				'Dtl SEQ
						strVal = strVal & Trim(.Text) & parent.gColSep
					
						.Col =  3 'C_CtrlCd + 1		 		'�����׸��ڵ� 
						strVal = strVal & Trim(.Text) & parent.gColSep
						        
						.Col = 5 'C_CtrlVal + 1				'�����׸� Value 
						strVal = strVal & UCase(Trim(.Text)) & parent.gRowSep	
				
						lGrpCnt = lGrpCnt + 1
					End If			
		    	Next
			End If
		Next
    End With

	With frm1
		.txtMaxRows3.Value = lGrpCnt-1										'Spread Sheet�� ����� �ִ밹�� 
		.txtSpread3.Value  = strVal											'Spread Sheet ������ ���� 

		'���Ѱ����߰� start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'���Ѱ����߰� end
    End With

    Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'���� �����Ͻ� ASP �� ���� 
        
    DbSave = True                                                           
End Function

'========================================================================================
Function DbSaveOk(Byval GlNo)					'��: ���� ������ ���� ���� 
	frm1.txtGlNo.Value = UCase(Trim(GlNo))
    frm1.txtCommandMode.Value = "UPDATE"

	Call ggoOper.ClearField(Document, "2")      '��: Condition field clear    
    Call InitVariables							'��: Initializes local global variables
	lgIntFlgMode = parent.OPMD_UMODE
	
	DbQuery
End Function

'========================================================================================
Function DbDelete()
	Dim strVal
	
    Err.Clear
    Call LayerShowHide(1)    
	DbDelete = False														'��: Processing is NG

	'//frm1.hOrgChangeId.Value = parent.gChangeOrgId

	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003						'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtGlNo=" & UCase(Trim(frm1.txtGlNo.Value))	'��: ���� ���� ����Ÿ 
    strVal = strVal & "&txtGlDt=" & ggoOper.RetFormat(frm1.txtGLDt.Text, "yyyy-MM-dd")
    strVal = strVal & "&txtDeptCd=" & UCase(Trim(frm1.txtDeptCd.Value))
	strVal = strVal & "&txtOrgChangeId=" & Trim(frm1.hOrgChangeId.Value)
    strVal = strVal & "&txtGlinputType=" & Trim(frm1.txtGlinputType.Value)

	' ���Ѱ��� �߰� 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 
    
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True                                                         '��: Processing is NG

End Function

'=======================================================================================================
Function DbDeleteOk()													'���� ������ ���� ���� 
	Call FncNew()
End Function
'====================================================================================================
Sub CurFormatNumSprSheet(Row)

	With frm1
		ggoSpread.Source = frm1.vspdData
		.vspdData.Row	= Row
		'�ݾ� 
'			Call FixDecimalPlaceByCurrency2(frm1.vspdData,Row,C_DocCur,C_ItemAmt,"A" ,"X","X")
       Call ReFormatSpreadCellByCellByCurrency(Frm1.vspdData,Row,Row,C_DocCur ,C_ItemAmt,"A" ,"I","X","X")         
		'ggoSpread.SSSetFloatByCellOfCur C_ItemAmt, Row, .vspdData.Text, parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, parent.gComNum1000, parent.gComNumDec
		
	End With

End Sub    

'=======================================================================================================    
Sub SetGridFocus()	
    
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1	

End Sub

'=======================================================================================================  
Sub InputCtrlVal(ByVal Row)
			
		Dim strAcctCd		
		Dim ii
			
		lgBlnFlgChgValue = True
		
		ggoSpread.Source = frm1.vspdData
		frm1.vspdData.Col = C_AcctCd
		frm1.vspdData.Row = Row		
		strAcctCd	= Trim(frm1.vspdData.text)		
		
		frm1.vspdData.Col = C_deptcd
		frm1.vspdData.Row = Row			
		
		Call AutoInputDetail(strAcctCd, Trim(frm1.vspdData.text), frm1.txtGLDt.text, Row)
		For ii = 1 To frm1.vspdData2.MaxRows
			frm1.vspddata2.col = C_CtrlVal
			frm1.vspddata2.row = ii
					
			If Trim(frm1.vspddata2.text) <> "" Then
				Call CopyToHSheet2(frm1.vspdData.ActiveRow,ii)			 			
			End If
		Next
		
End Sub

'=======================================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo,varOrgChangeId)
	Dim intRetCd

	StrEbrFile = "a5121ma1"

	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtGlDt.Text, parent.gDateFormat,"")	
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtGlDt.Text, parent.gDateFormat,"")	
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	varGlNoFr = Trim(frm1.txtGlNo.Value)
	varGlNoTo = Trim(frm1.txtGlNo.Value)
	varOrgChangeId = Trim(frm1.hOrgChangeId.Value)	
End Sub

'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
	Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId
    Dim StrEbrFile
    Dim intRetCd
	Dim ObjName
	
    If Not chkFieldByCell(frm1.txtGlNo,"A",1) Then Exit Function
    	
'    If Not chkField(Document, "1") Then	
'       Exit Function
'    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId)

    lngPos = 0

	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|GlNoFr|" & varGlNoFr
	StrUrl = StrUrl & "|GlNoTo|" & varGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)

End Function

'========================================================================================
Function FncBtnPreview() 
    Dim VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId
    Dim StrUrl
    Dim arrParam, arrField, arrHeader
    Dim StrEbrFile
    Dim intRetCD
	Dim ObjName

    If Not chkFieldByCell(frm1.txtGlNo,"A",1) Then Exit Function

'    If Not chkField(Document, "1") Then
'       Exit Function
'    End If

	Call SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo, varOrgChangeId)

    StrUrl = StrUrl & "DateFr|" & VarDateFr
	StrUrl = StrUrl & "|DateTo|" & VarDateTo
	StrUrl = StrUrl & "|DeptCd|" & VarDeptCd
	StrUrl = StrUrl & "|BizAreaCd|" & VarBizAreaCd
	StrUrl = StrUrl & "|GlNoFr|" & varGlNoFr
	StrUrl = StrUrl & "|GlNoTo|" & varGlNoTo
	StrUrl = StrUrl & "|OrgChangeId|" & varOrgChangeId
	StrUrl = StrUrl & "|GlPutType|" & "%"

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf"><IMG src="../../../CShared/image/table/seltab_up_left.gIf" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="center" CLASS="CLSMTABP"><font color=white>ȸ����ǥ��� (Copy)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gIf" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gIf" width="10" height="23" ></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%> </TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǥ��ȣ</TD>
									<TD CLASS=TD656 NOWRAP><INPUT NAME="txtGlNo" ALT="��ǥ��ȣ" MAXLENGTH="18" SIZE=20 STYLE="TEXT-ALIGN: left" tag  ="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gIf" NAME="btnTempGlNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRefGL()"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%" ></TD>
				</TR>
				<TR>		
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP >
						<TABLE <%=LR_SPACE_TYPE_60%>>					
							<TR>
								<TD CLASS=TD5 NOWRAP>��ǥ����</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtGLDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="22" ALT="ȸ������" id=OBJECT7></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>��ǥ����</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlType" tag="23" STYLE="WIDTH:82px:" ALT="��ǥ����"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�μ�</TD>								
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="�μ��ڵ�" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gIf" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpEndept(frm1.txtDeptCd.Value, 0)">&nbsp;
													 <INPUT NAME="txtDeptNm" ALT="�μ���"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="24X"></TD>
													 <INPUT NAME="txtInternalCd" ALT="���κμ��ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"  TABINDEX="-1">
								<TD CLASS=TD5 NOWRAP>��ǥ�Է°��</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="cboGlInputType" tag="24" STYLE="WIDTH:82px:" ALT="��ǥ�Է°��"><OPTION VALUE="" selected></OPTION></SELECT></TD>								
			    
							</TR>							
							<TR>
								<TD CLASS=TD5 NOWRAP>���</TD>
								<TD CLASS=TD656 NOWRAP COLSPAN=3><INPUT NAME="txtDesc" ALT="���" MAXLENGTH="128" SIZE="70" tag="22N" ></TD>
							</TR>							
							<TR> 
								<TD HEIGHT="60%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>								
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD656 WIDTH=* align=right COLSPAN=2><BUTTON NAME="btnCalc" CLASS="CLSSBTNCALC" ONCLICK="vbscript:FncBtnCalc()" Flag=1>�ڱ��ݾװ��</BUTTON>&nbsp;
								<TD CLASS=TD5 NOWRAP>�����հ�(�ڱ�)</TD>
								<TD CLASS=TD6 NOWRAP>&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtDrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="�����հ�(�ڱ�)" id=OBJECT3></OBJECT>');</SCRIPT>
									&nbsp;<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name=txtCrLocAmt style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE tag="24X2" ALT="�뺯�հ�(�ڱ�)" id=OBJECT4></OBJECT>');</SCRIPT></TD>							
							</TR>
							<TR>						                 
								<TD HEIGHT="40%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT5> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>									
			  			  
								</TD>
							</TR>
							<!--<TR>						                 
								<TD HEIGHT="40%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=OBJECT5> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>									
			  			  
								</TD>
							</TR>-->
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTToN NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>�̸�����</BUTToN>&nbsp;
						<BUTToN NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTToN>&nbsp;
					</TD>										
					<TD WIDTH=* ALIGN=RIGHT>					
					<TD WIDTH="*" ALIGN=RIGHT><A HREF = "VBSCRIPT:PgmJumpChk(JUMP_PGM_ID_TAX_REP)">ȸ����ǥ���</A></TD>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IfRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IfRAME></TD>
		<!--<TD WIDTH="100%" HEIGHT=30%><IfRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IfRAME></TD>-->
	</TR>

</TABLE>
<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 width="100%" tag="23" TITLE="SPREAD" id=OBJECT6 TABINDEX="-1"><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
<TEXTAREA class=hidden name=txtSpread		tag="24" TABINDEX="-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3		tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtGlNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtCommandMode"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hInternalCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hOrgChangeId"		tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtAuthorityFlag"  tag="24" TABINDEX="-1"><!--���Ѱ����߰� -->
<INPUT TYPE=HIDDEN NAME="txtGlinputType"	tag="24" TABINDEX="-1">

<INPUT TYPE=hidden NAME="hItemSeq"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hAcctCd"			tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows3"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<Iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></Iframe>
</DIV>
</BODY>
</HTML