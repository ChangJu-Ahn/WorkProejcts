<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5123MA1
'*  4. Program Name         : ȸ����ǥ����� 
'*  5. Program Desc         : �� ���쿡�� ������ �ڷḦ ���� �ϰ������� ��ǥó��.
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/01/23
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     : Kim Ho Young 
'* 10. Modifier (Last)      : Lim YOung Woon
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

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit  

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const BIZ_PGM_ID1 = "a5465mb1.asp"												'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_ID2 = "a5465mb2.asp"												'��: �����Ͻ� ���� ASP�� 
'==========================================================================================================
Const GRID_POPUP_MENU_NEW	=	"0000111111"
Const GRID_POPUP_MENU_CRT	=	"0000111111"
Const GRID_POPUP_MENU_UPD	=	"0001111111"
Const GRID_POPUP_MENU_PRT	=	"0000111111"		

'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================= 
Dim  C_Confirm       
Dim  C_BatchDt       
Dim  C_BatchNo       
Dim  C_BizCD         
Dim  C_BizNm         
Dim  C_Refno         
Dim  C_GLInputType   
Dim  C_GLInputTypeNm 
Dim  C_GlDt          
Dim  C_GlNo          
Dim  C_TEMP_Gl_FG    
Dim  C_BP_CD3
Dim  C_Bp_CD3_NM

'========================================================================================================= 

Dim lgStrPrevKeyTempGlDt
Dim lgStrPrevKeyBatchNo

Dim lgQueryFlag					' �ű���ȸ �� �߰���ȸ ���� Flag
Dim lgAllSelect


Dim  IsOpenPop          

Dim lgGridPoupMenu              ' Grid Popup Menu Setting


'========================================================================================================
Sub InitSpreadPosVariables()
    C_Confirm         = 1															'��: Spread Sheet�� Column�� ��� 
    C_BatchDt         = 2														'��: Spread Sheet�� Column�� ���  
	C_BP_CD3		= 3	
	C_BP_CD3_NM		= 4	
    C_BizCD           = 5
    C_BizNm           = 6
    C_Refno           = 7
    C_GLInputType     = 8
    C_GLInputTypeNm   = 9
    C_GlDt            = 10
    C_GlNo            = 11
    C_BatchNo         = 12
    C_TEMP_Gl_FG      = 13

End Sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode            = parent.OPMD_CMODE 
    lgBlnFlgChgValue        = False  
    lgIntGrpCount           = 0   
    
    lgStrPrevKeyTempGlDt    = ""              
    lgStrPrevKeyBatchNo     = ""                       'initializes Previous Key
    lgLngCurRows            = 0                            'initializes Deleted Rows Count
    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
    Dim StartDate
    Dim EndDate
    Dim strYear
    Dim strMonth
    Dim strDay

	Call ExtractDateFrom("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	StartDate	= UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")		'��: �ʱ�ȭ�鿡 �ѷ����� ���� ��¥ 
	EndDate		= UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)		'��: �ʱ�ȭ�鿡 �ѷ����� ������ ��¥ 

	frm1.txtFromReqDt.text =  StartDate
	frm1.txtToReqDt.text   =  EndDate
	frm1.cboConfFg.value	=	"U"
	Call cboConfFg_OnChange()
	lgGridPoupMenu          = GRID_POPUP_MENU_PRT
	frm1.txtGlInputType.focus	
End Sub
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>

End Sub
'========================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables()

	With frm1.vspdData
	
    .MaxCols = C_TEMP_Gl_FG+1									'��: �ִ� Columns�� �׻� 1�� ������Ŵ 
	.Col = .MaxCols										'��: ����� �� Hidden Column
    .ColHidden = True
           
    .MaxRows = 0
    ggoSpread.Source = frm1.vspdData

	.ReDraw = false
	
    ggoSpread.Spreadinit "V20021127",, parent.gAllowDragDropSpread
    .ReDraw = false

    Call GetSpreadColumnPos("A")
	'SSSetEdit(Col, Header, ColWidth , HAlign , Row , Length)    
    ggoSpread.SSSetCheck C_Confirm       ,   ""              ,     8,  -10, "", True, -1 
    ggoSpread.SSSetDate C_BatchDt        ,   "�߻���"     , 10,,parent.gDateFormat
    ggoSpread.SSSetEdit C_BatchNo        ,   "��ġ��ȣ"    , 15,,,20
    ggoSpread.SSSetEdit C_BizCD          ,   "�����"     , 10,,,10
    ggoSpread.SSSetEdit C_BizNm          ,   "������"   , 15,,,20
    ggoSpread.SSSetEdit C_Refno          ,   "������ȣ"   , 20,,,20                                
    
	ggoSpread.SSSetEdit C_GLInputType    ,   "�Է°��"   ,       10,,,3
	ggoSpread.SSSetEdit C_GLInputTypeNm  ,   "�Է°�θ�" , 15,,,30
    
    ggoSpread.SSSetDate C_GlDt           ,   "��ǥ��"    ,   10,,parent.gDateFormat
    ggoSpread.SSSetEdit C_GlNo           ,   "��ǥ��ȣ"   , 20,,,20
    ggoSpread.SSSetEdit C_Bp_CD3          ,   "�ŷ�ó"     , 10,,,10
    ggoSpread.SSSetEdit C_Bp_CD3_NM          ,   "�ŷ�ó��"     , 10,,,30

    Call ggoSpread.SSSetColHidden(C_TEMP_Gl_FG,C_TEMP_Gl_FG,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

'========================================================================================

Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock C_BatchDt          , -1    ,C_BatchDt
        ggoSpread.spreadlock C_BatchNo          , -1    ,C_BatchNo
        ggoSpread.spreadlock C_BizCD            , -1    ,C_BizCD
        ggoSpread.spreadlock C_BizNm            , -1    ,C_BizNm
        ggoSpread.spreadlock C_Refno            , -1    ,C_Refno
        ggoSpread.spreadlock C_GLInputType      , -1    ,C_GLInputType
        ggoSpread.spreadlock C_GLInputTypeNm    , -1    ,C_GLInputTypeNm
        ggoSpread.spreadlock C_GlDt             , -1    ,C_GlDt
        ggoSpread.spreadlock C_GlNo             , -1    , C_GlNo   
		ggoSpread.spreadlock C_Bp_CD3			, -1    , C_Bp_CD3   
		ggoSpread.spreadlock C_Bp_CD3_NM			, -1    , C_Bp_CD3_NM   
        ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
   
        .vspdData.ReDraw = True
    End With
End Sub

'========================================================================================

Sub SetSpreadColor(ByVal lRow)
    With frm1
    
    .vspdData.ReDraw = False    
    ggoSpread.SSSetProtected	C_BatchDt       , lRow, lRow
    'ggoSpread.SSSetProtected	C_BatchNo   , lRow, lRow
    ggoSpread.SSSetProtected	C_BizCD, lRow   , lRow
    ggoSpread.SSSetProtected	C_BizNm, lRow   , lRow
    ggoSpread.SSSetProtected	C_Refno, lRow   , lRow
    ggoSpread.SSSetProtected	C_GLInputType   , lRow, lRow
    ggoSpread.SSSetProtected	C_GLInputTypeNm , lRow, lRow
    ggoSpread.SSSetProtected	C_GlDt, lRow    , lRow
    ggoSpread.SSSetProtected	C_GlNo, lRow    , lRow
	ggoSpread.SSSetProtected	C_Bp_CD3, lRow    , lRow
	ggoSpread.SSSetProtected	C_Bp_CD3_NM, lRow    , lRow
    .vspdData.ReDraw = True
    
    End With
End Sub
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Confirm         = iCurColumnPos(1)
            C_BatchDt         = iCurColumnPos(2)
            C_Bp_CD3            = iCurColumnPos(3)
            C_Bp_CD3_NM      = iCurColumnPos(4)
            C_BizCD           = iCurColumnPos(5) 
            C_BizNm           = iCurColumnPos(6) 
            C_Refno           = iCurColumnPos(7) 
            C_GLInputType     = iCurColumnPos(8) 
            C_GLInputTypeNm   = iCurColumnPos(9) 
            C_GlDt            = iCurColumnPos(10) 
            C_GlNo            = iCurColumnPos(11)
            C_BatchNo         = iCurColumnPos(12) 
            C_TEMP_Gl_FG      = iCurColumnPos(13)

       End Select    
End Sub

 '========================================================================================
'                       InitComboBox_cond()
' ========================================================================================  
Sub InitComboBox_cond()
	Dim intRetCd,intLoopCnt
	Dim ArrayTemp1
	Dim ArrayTemp2
	IntRetCd = CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("A1007", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 		    
	
	If IntRetCD=False  Then
	    Call DisplayMsgBox("122300","X","X","X")                         '�� : Minor�ڵ������� �����ϴ�.
	Else
		ArrayTemp1 = Split(lgF0,Chr(11))
		ArrayTemp2 = Split(lgF1,Chr(11))

		For intLoopCnt = 0 To UBound(ArrayTemp1,1) -1
			Call SetCombo(frm1.cboConfFg, ArrayTemp1(intLoopCnt), ArrayTemp2(intLoopCnt))
		Next  

	End If
End Sub

'=======================================================================================================
Sub txtFromReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromReqDt.focus    
    End If
End Sub
'========================================================================================================= 
Sub txtToReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToReqDt.focus    
    End If
End Sub

'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0			
			arrParam(0) = "��ǥ�������"					<%' �˾� ��Ī %>
			arrParam(1) = "B_MINOR" 				<%' TABLE ��Ī %>
			arrParam(2) = strCode						<%' Code Condition%>
			arrParam(3) = ""							<%' Name Cindition%>
			arrParam(4) = " MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  "							<%' Where Condition%> 
			arrParam(5) = "��ǥ�Է°��"						<%' �����ʵ��� �� ��Ī %>

			arrField(0) = "MINOR_CD"						<%' Field��(0)%>
			arrField(1) = "MINOR_NM"						<%' Field��(1)%>
    
			arrHeader(0) = "��ǥ�Է°��"					<%' Header��(0)%>
			arrHeader(1) = "��ǥ�Է°�θ�"					<%' Header��(1)%>
		Case 1
			arrParam(0) = "������˾�"  				' �˾� ��Ī 
			arrParam(1) = "B_BIZ_AREA"	 			' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "�����"	    				' �����ʵ��� �� ��Ī 

			arrField(0) = "BIZ_AREA_CD"						' Field��(0)
			arrField(1) = "BIZ_AREA_NM"						' Field��(1)
    
			arrHeader(0) = "�����"	     				' Header��(0)
			arrHeader(1) = "������"					' Header��(1)
		Case 2			
			arrParam(0) = "�ŷ�ó�˾�"						' �˾� ��Ī 
			arrParam(1) = "b_biz_partner"						' TABLE ��Ī 
			arrParam(2) = strCode						 	    ' Code Condition
			arrParam(3) = ""									' Name Cindition
			arrParam(4) = "BP_TYPE<>" & FilterVar("S", "''", "S") & " "									' Where Condition
			arrParam(5) = "�ŷ�ó"			
	
			arrField(0) = "BP_CD"								' Field��(0)
			arrField(1) = "BP_NM"								' Field��(1)
    
			arrHeader(0) = "�ŷ�ó"							' Header��(0)
			arrHeader(1) = "�ŷ�ó��"						' Header��(1)
		Case 3
			arrParam(0) = "�ŷ����� �˾�"    ' �˾� ��Ī 
			arrParam(1) = "A_ACCT_TRANS_TYPE"    ' TABLE ��Ī 
			arrParam(2) = strCode      ' Code Condition
			arrParam(3) = ""       ' Name Cindition
			arrParam(4) = " MO_CD NOT IN (" & FilterVar("A", "''", "S") & " ," & FilterVar("F", "''", "S") & " ) "       ' Where Condition
			arrParam(5) = "�ŷ�����"     ' �����ʵ��� �� ��Ī 

			arrField(0) = "TRANS_TYPE"     ' Field��(0)
			arrField(1) = "TRANS_NM"     ' Field��(1)

			arrHeader(0) = "�ŷ������ڵ�"   ' Header��(0)
			arrHeader(1) = "�ŷ�������"    ' Header��(1)		
	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If	

	Call EscPopUp(iWhere)
End Function
'========================================================================================================= 

Function OpenPopupGL()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	
    
	On Error Resume Next

	If IsOpenPop = True Then Exit Function
    frm1.vspdData.Col =  C_TEMP_Gl_FG
    if Trim(frm1.vspdData.Text) = "G" THEN

		iCalledAspName = AskPRAspName("a5120ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
	ELSEIF	Trim(frm1.vspdData.Text) = "T" THEN	     
		iCalledAspName = AskPRAspName("a5130ra1")
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
			IsOpenPop = False
			Exit Function
		End If
	END IF		

	With frm1.vspdData
		.Row = .ActiveRow
		.Col =  C_GlNo
		arrParam(0) = Trim(.Text)	'������ǥ��ȣ 
		arrParam(1) = ""			'Reference��ȣ 

		if arrParam(0) = "" THEN Exit Function
			
	End With

	IsOpenPop = True
   
    frm1.vspdData.Col =  C_TEMP_Gl_FG
    if Trim(frm1.vspdData.Text) = "G" THEN
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	ELSEIF	Trim(frm1.vspdData.Text) = "T" THEN	     
		arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
			     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	END IF		
	IsOpenPop = False
	
End Function
'=======================================================================================

Function EscPopUp(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtGlInputType.focus
			Case 1
				.txtBizCd.focus
			Case 2
				.txtBpCd.focus

		End Select
	End With
	
End Function

'========================================================================================================= 

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				frm1.txtGlInputType.value = arrRet(0)
				frm1.txtGlInputTypeNm.value = arrRet(1)								
			Case 1
				frm1.txtBizCd.value  = arrRet(0)
				frm1.txtBizNm.value  = arrRet(1)			    
			Case 2
				frm1.txtBpCd.value  = arrRet(0)
				frm1.txtBpNm.value  = arrRet(1)			    
			Case 3
				.txtTransType.value  = arrRet(0)
				.txtTransNm.value  = arrRet(1)			    
				.txtTransType.focus
		End Select

	End With
	
End Function
'========================================================================================================= 

Sub txtBizCd_onBlur()
	
	if frm1.txtBizCd.value = "" then
		frm1.txtBizNm.value = ""
	end if
End Sub	
'========================================================================================================= 

Sub txtGlInputType_onBlur()
	
	if frm1.txtGlInputType.value = "" then
		frm1.txtGlInputTypeNm.value = ""
	end if
End Sub	
		
'========================================================================================================= 
Function fnBttnConf()	
	Dim IntRetCd
	
	IntRetCD = DisplayMsgBox("112190", parent.VB_YES_NO,"x","x")
	
	If IntRetCD = vbNo Then
		Exit Function
	End if	
      
	fnBttnConf = False                                                          '��: Processing is NG
	
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value		  = parent.UID_M0002
		.htxtWorkFg.value	  = "CONF"		
		.txtUpdtUserId.value  = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID    				
    END With
    
    Call ExecMyBizASP(frm1, BIZ_PGM_ID2)									'��: �����Ͻ� ASP �� ���� 
    
    fnBttnConf = True             

End Function
'========================================================================================================= 
Function fnBttnUnConf()
	Dim IntRetCd
	
	IntRetCD = DisplayMsgBox("112191", parent.VB_YES_NO,"x","x")
	If IntRetCD = vbNo Then
		Exit Function
	End if	
      
	fnBttnUnConf = False                                                          '��: Processing is NG
	
	Call LayerShowHide(1)
	
	With frm1
		.txtMode.value		  = parent.UID_M0002
		.htxtWorkFg.value	  = "UNCONF"
		.txtUpdtUserId.value  = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID    				
    END With
    
    Call ExecMyBizASP(frm1, BIZ_PGM_ID2)									'��: �����Ͻ� ASP �� ���� 
    
    fnBttnUnConf = True             

End Function
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
    Call InitSpreadSheet                                                    '��: Setup the Spread sheet
    Call InitVariables                                                      '��: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox_Cond
    Call SetDefaultVal
    Call SetToolbar("110000000000111")
    
    frm1.btnConf.disabled	=	True
    frm1.btnUnCon.disabled	=	True

End Sub

'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
Sub  cboConfFg_OnChange()
    lgBlnFlgChgValue = True

	'IF frm1.cboConfFg.value = "C" Then
	'	frm1.btnConf.disabled	=	True
	'	frm1.btnUnCon.disabled	=	False
	'ELSE
	'	frm1.btnConf.disabled	=	False
	'	frm1.btnUnCon.disabled	=	True
	'END IF	
	
End Sub

'========================================================================================================= 
Sub txtFromReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtToReqDt.focus
		Call FncQuery
	End If
End Sub
'========================================================================================================= 

Sub txtToReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt.focus
		Call FncQuery
	End If
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    Call SetPopupMenuItemInf(lgGridPoupMenu)
    gMouseClickStatus = "SPC"   
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Set gActiveSpdSheet = frm1.vspdData
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

	Select Case Col
	
		Case C_Confirm 							
			ggoSpread.Source = frm1.vspdData
'			ggoSpread.UpdateRow Row	
			lgBlnFlgChgValue = True						
	End Select 	
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 

End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'==========================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
	ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row	
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================

Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'==========================================================================================

Sub vspdData_KeyPress(index , KeyAscii )
     lgBinFlgChgValue = True                                                 '��: Indicates that value changed
End Sub
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    If Col <= C_Confirm Or NewCol <= C_Confirm Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'==========================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow         
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyBatchNo <> "" Then                         
      	   Call DbQuery
    	End If
    End if
        
    
End Sub

'========================================================================================

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtFromReqDt.text,frm1.txtToReqDt.text,frm1.txtFromReqDt.Alt,frm1.txtToReqDt.Alt, _
                        "970025",frm1.txtFromReqDt.UserDefinedFormat,parent.gComDateType,True) = False Then	
		frm1.txtFromReqDt.focus
		Exit Function
	End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	'Call InitSpreadSheet    
    Call InitVariables 															'��: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    
    
	lgQueryFlag = "New"		' �ű���ȸ �� �߰���ȸ ���� Flag (����� �ű���)
	
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'��: Query db data
       
    FncQuery = True																'��: Processing is OK
    
End Function
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    'On Error Resume Next                                                    '��: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X") '�� �ٲ�κ�    
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
   
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")                                  '��: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                  '��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    
    FncNew = True                                                           '��: Processing is OK

End Function

'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    On Error Resume Next                                                    '��: Protect system from crashing
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False  Then  '��: Check If data is chaged
        IntRetCD = DisplayMsgBox("900001","X","X","X")            '��: Display Message(There is no changed data.)
        Exit Function
    End If

    If Not chkField(Document, "1") Then               '��: Check required field(Single area)
       Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtFromReqDt.text,frm1.txtToReqDt.text,frm1.txtFromReqDt.Alt,frm1.txtToReqDt.Alt, _
                        "970025",frm1.txtFromReqDt.UserDefinedFormat,parent.gComDateType,True) = False Then		
		frm1.txtFromReqDt.focus
		Exit Function
	End If

  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If
    
    '-----------------------
    'Save function call area
    '-----------------------
    IF DbSave	= False Then			                                                  '��: Save db data
		 Exit Function
    End If
    
   	
    FncSave = True                                                          '��: Processing is OK
    
End Function
'========================================================================================

Function FncCancel() 

    if frm1.vspdData.MaxRows < 1 then Exit Function

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '��: Protect system from crashing
    
End Function

'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 '��: ȭ�� ���� 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      '��:ȭ�� ����, Tab ���� 
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
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'=======================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================

Function DbQuery() 
	Dim strVal

    DbQuery = False
    Call LayerShowHide(1)
    frm1.btnConf.disabled	=	True
	frm1.btnUnCon.disabled	=	True
    
    Err.Clear                                                               '��: Protect system from crashing
    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID1 & "?txtMode=" & parent.UID_M0001						'��:��ȸǥ�� 			
			strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
			strVal = strVal & "&lgStrPrevKeyBatchNo=" & lgStrPrevKeyBatchNo
			strVal = strVal & "&txtBizCd="         & Trim(.hBizCd.value)	 			    '��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtGlInputType="   & Trim(.hGlInputType.value)

			strVal = strVal & "&txtfrRefNo="   & Trim(.txtfrRefNo.value)
			strVal = strVal & "&txttoRefNo="   & Trim(.txttoRefNo.value)

			strVal = strVal & "&cboConfFg="        & Trim(.hcboConfFg.value)
			strVal = strVal & "&txtFromReqDt="     & (.txtFromReqDt.Text)
			strVal = strVal & "&txtToReqDt="       & (.txtToReqDt.Text)
			strVal = strVal & "&txtBpcd="       & (.hBpcd.value)
			strVal = strVal & "&txtTransType="       & (.hTransType.value)

			strVal = strVal & "&txtMaxRows="       & .vspdData.MaxRows
		Else
			strVal = BIZ_PGM_ID1 & "?txtMode="     & parent.UID_M0001						'��:��ȸǥ�� 			
			strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
			strVal = strVal & "&lgStrPrevKeyBatchNo=" & lgStrPrevKeyBatchNo
			strVal = strVal & "&txtBizCd="         & Trim(.txtBizCd.value)	 			    '��: ��ȸ ���� ����Ÿ 
			strVal = strVal & "&txtGlInputType="   & Trim(.txtGlInputType.value)
			strVal = strVal & "&txtfrRefNo="   & Trim(.txtfrRefNo.value)
			strVal = strVal & "&txttoRefNo="   & Trim(.txttoRefNo.value)

			strVal = strVal & "&cboConfFg="        & Trim(.cboConfFg.value)
			strVal = strVal & "&txtFromReqDt="     & (.txtFromReqDt.Text)		
			strVal = strVal & "&txtToReqDt="       & (.txtToReqDt.Text)
			strVal = strVal & "&txtBpcd="       & (.txtBpcd.value)
			strVal = strVal & "&txtTransType="       & (.txtTransType.value)

			strVal = strVal & "&txtMaxRows="       & .vspdData.MaxRows
		End If
		
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ����        
    End With    
    DbQuery = True
End Function
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
    Call LayerShowHide(0)

    Call SetToolbar("110010000001111")
    lgGridPoupMenu  =   GRID_POPUP_MENU_UPD
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.cboConfFg.value = "C" Then
			frm1.btnConf.disabled	=	True
			frm1.btnUnCon.disabled	=	False
		Else
			frm1.btnConf.disabled	=	False
			frm1.btnUnCon.disabled	=	True
		End If	
	End If
		
End Function
'========================================================================================
Function SetGridFocus()
	with frm1 
		.vspdData.Col = 1
		.vspdData.Row = 1
		.vspdData.Action = 1
	end with 
End Function 

'========================================================================================

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim iSelectCnt
	
    DbSave = False                                                          '��: Processing is NG
    Call LayerShowHide(1)
    
    'On Error Resume Next                                                   '��: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		iSelectCnt = 0
		lgAllSelect = False
		
		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = C_Confirm
        
			If frm1.vspdData.text = "1" THEN

					strVal = strVal & "U" & parent.gColSep				'��: U=Update
					.vspdData.Col = C_BatchNo		'4
					strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
					.vspdData.Col = C_Confirm	'1
					if Trim(.cboConfFg.value)  = "U" THEN
						strVal = strVal & "Y" & parent.gRowSep
					ELSE
						strVal = strVal & "N" & parent.gRowSep
					END IF	
					lGrpCnt = lGrpCnt + 1
					iSelectCnt = iSelectCnt + 1	  
			End if
		         
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal	
		If iSelectCnt = .vspdData.MaxRows Then
			lgAllSelect = True
		End If		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID1)									'��: �����Ͻ� ASP �� ���� 
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function


'========================================================================================

Function DbSaveOk()													'��: ���� ������ ���� ���� 
   
    Call LayerShowHide(0)

	ggoSpread.Source = frm1.vspdData
	frm1.vspdData.ReDraw = False
	ggoSpread.SSDeleteFlag 1 , frm1.vspdData.MaxRows
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call SetSpreadLock
	frm1.vspdData.ReDraw = True
	
	Call InitVariables	
	If lgAllSelect = True Then
		IF frm1.cboConfFg.value = "C" Then
			frm1.cboConfFg.value = "U"
		Else
			frm1.cboConfFg.value = "C"
		End If
	End If
	Call DBQuery()		
End Function


'=======================================================================================================
Function FncExit()
Dim IntRetCD

	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function
'=======================================================================================================
Sub txtTransType_onchange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtTransType.value = "" Then frm1.txtTransNm.value = "":	Exit Sub

	If CommonQueryRs("TRANS_NM", "A_ACCT_TRANS_TYPE ", " TRANS_TYPE=  " & FilterVar(frm1.txtTransType.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtTransNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtTransType.alt,"X")  	
		frm1.txtTransType.value = ""
		frm1.txtTransNm.value = ""
		frm1.txtTransType.focus
	End If
End Sub	
'=======================================================================================================
Sub txtGlInputType_onchange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtGlInputType.value = "" Then frm1.txtGlInputTypeNm.value = "":	Exit Sub

	If CommonQueryRs("MINOR_NM", "B_MINOR ", " MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  AND MINOR_CD=  " & FilterVar(frm1.txtGlInputType.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtGlInputTypeNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtGlInputType.alt,"X")  	
		frm1.txtGlInputType.value = ""
		frm1.txtGlInputTypeNm.value = ""
		frm1.txtGlInputType.focus
	End If
	if frm1.txtGlInputType.value = "" then
		frm1.txtGlInputTypeNm.value = ""
	end if
End Sub	
'=======================================================================================================
Sub txtBizCd_onChange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtBizCd.value = "" Then frm1.txtBizNm.value = "":	Exit Sub

	If CommonQueryRs("BIZ_AREA_NM", "B_BIZ_AREA ", " BIZ_AREA_CD=  " & FilterVar(frm1.txtBizCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBizNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtBizCd.alt,"X")  	
		frm1.txtBizCd.value = ""
		frm1.txtBizNm.value = ""
		frm1.txtBizCd.focus
	End If
End Sub	
'=======================================================================================================
Sub txtBpCd_onChange()	
	Dim IntRetCD
	Dim arrVal
	If frm1.txtBpCd.value = "" Then frm1.txtBpNm.value = "":	Exit Sub

	If CommonQueryRs("BP_NM", "B_BIZ_PARTNER ", " BP_CD=  " & FilterVar(frm1.txtBPCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtBpNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("970000","X",frm1.txtBpCd.alt,"X")  	
		frm1.txtBpCd.value = ""
		frm1.txtBpNm.value = ""
		frm1.txtBpCd.focus
	End If
End Sub	
'=======================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ȸ����ǥ�����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupGL()">ȸ����ǥ</A></TD>					
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD  <%=HEIGHT_TYPE_02%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5"NOWRAP>�߻�����</TD>
									<TD CLASS="TD6"NOWRAP>
										<script language =javascript src='./js/a5465ma1_fpDateTime1_txtFromReqDt.js'></script>
~ 
										<script language =javascript src='./js/a5465ma1_fpDateTime2_txtToReqDt.js'></script>										
									</TD>
									<TD CLASS="TD5"NOWRAP>��ǥ�Է°��</TD>
									<TD CLASS="TD6"NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtGlInputType" SIZE=10  MAXLENGTH=10 tag="11XXXU" ALT="��ǥ�Է°��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtGlInputType.Value, 0)">
										 <INPUT TYPE=TEXT ID="txtGlInputTypeNm" NAME="txtGlInputTypeNm" SIZE=20 tag="14X" ALT="��ǥ�Է°�θ�">
									</TD>
						
								</TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>�����</TD>
									<TD CLASS="TD6"NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizCd.Value, 1)">
										 <INPUT TYPE=TEXT ID="txtBizNm" NAME="txtBizNm" SIZE=20 tag="14X" ALT="������">
									</TD>
									<TD CLASS="TD5"NOWRAP>��������</TD>
									<TD CLASS="TD6"NOWRAP><SELECT NAME="cboConfFg" tag="12" STYLE="WIDTH:82px:" Alt="��������"><OPTION VALUE="" selected></OPTION></SELECT>
								</TR>
								<TR>
									<TD CLASS="TD5"NOWRAP>�ŷ����� </TD>
									<TD CLASS="TD6"NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtTransType" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�ŷ�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtTransType.Value, 3)">
										 <INPUT TYPE=TEXT ID="txtTransNm" NAME="txtTransNm" SIZE=20 tag="14X" ALT="�ŷ�������">
									</TD>
									<TD CLASS="TD5"NOWRAP>�ŷ�ó</TD>
									<TD CLASS="TD6"NOWRAP><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="�ŷ�ó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBpCd.Value, 2)">
										 <INPUT TYPE=TEXT ID="txtBpNm" NAME="txtBpNm" SIZE=20 tag="14X" ALT="�ŷ�ó��">
									</TD>

								</TR>

								<TR>
									<TD CLASS=TD5 NOWRAP>������ȣ</TD>				
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtfrRefNo" SIZE=18 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="������ȣ">&nbsp;~&nbsp;
														 <INPUT TYPE="Text" NAME="txttoRefNo" SIZE=18 MAXLENGTH=25 STYLE="TEXT-ALIGN: left" tag="1XXXXU" ALT="������ȣ"></TD>

									<TD CLASS="TD5"NOWRAP></TD>
									<TD CLASS="TD6"NOWRAP></TD>
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>				
							<TR>
								<TD HEIGHT="100%"><script language =javascript src='./js/a5465ma1_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>							
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE01%>></TD>
	</TR>			
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>				
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnConf" CLASS="CLSMBTN" OnClick="VBScript:Call fnBttnConf()" >�ϰ�����</BUTTON>&nbsp;<BUTTON NAME="btnUnCon" CLASS="CLSMBTN" OnClick="VBScript:Call fnBttnUnConf()">�ϰ����</BUTTON></TD>		        					
					<TD WIDTH=10>&nbsp;</TD>
				</TR>	
			</TABLE>	
		</TD>						
	</TR>
	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH="100%" HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
		<!--<TD WIDTH=100% HEIGHT=30%><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>-->
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"   tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode"           tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtUpdtUserId"     tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtInsrtUserId"    tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txtMaxRows"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hGlInputType"      tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hBizCd"            tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hBpCd"            tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hTransType"            tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="hcboConfFg"        tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="htxtWorkFg"        tag="24" TABINDEX="-1">
<script language =javascript src='./js/a5465ma1_fpDateTime1_hFromReqDt.js'></script>
<script language =javascript src='./js/a5465ma1_fpDateTime2_hToReqDt.js'></script>										
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


