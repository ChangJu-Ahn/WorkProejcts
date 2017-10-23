<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : ȸ�����
*  2. Function Name        : �ΰ�������
*  3. Program ID           : A6119MA1_KO441
*  4. Program Name         : ���Լ��� �Ұ����� ���ٰŵ��
*  5. Program Desc         : ���Լ��� �Ұ����� ���ٰŵ��
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/10
*  8. Modified date(Last)  : 2004/12/27
*  9. Modifier (First)     : SHH
* 10. Modifier (Last)      : SHH
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "A6119MB1_KO441.asp"                                      'Biz Logic ASP
'Const BIZ_PGM_JUMP_ID = "H2001ma1" 
Const C_SHEETMAXROWS    = 500	                                      '�� ȭ�鿡 �������� �ִ밹��*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim	iDBSYSDate
Dim	EndDate, StartDate

	'------	��:	�ʱ�ȭ�鿡 �ѷ�����	������ ��¥	------
	EndDate	=	"<%=GetSvrDate%>"
	'------	��:	�ʱ�ȭ�鿡 �ѷ�����	���� ��¥	------
	StartDate	=	UNIDateAdd("m",	-1,	EndDate, Parent.gServerDateFormat)
	EndDate	=	UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
	StartDate	=	UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)



Dim lsConcd
Dim IsOpenPop          

DIM	 C_BIZ_AREA_CD
DIM  C_BIZ_AREA_POP
DIM	 C_BIZ_AREA_NM
DIM	 C_deduction_type
DIM	 C_deduction_type_POP
DIM	 C_deduction_type_NM
DIM	 C_YMD
DIM	 C_VAT_AMT
DIM	 C_TAX_CNT
DIM	 C_deduction_amt
DIM	 C_deduction_DESC

Dim  gSelframeFlg
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

	C_YMD					= 1	 
	C_BIZ_AREA_CD			= 2	 
	C_BIZ_AREA_POP			= 3 
	C_BIZ_AREA_NM			= 4	 
	C_deduction_type		= 5	 
	C_deduction_type_POP	= 6	 
	C_deduction_type_NM		= 7	 
	C_TAX_CNT				= 8	 
	C_deduction_amt			= 9	 
	C_VAT_AMT				= 10	 
	C_deduction_DESC		= 11 

End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================

Sub SetDefaultVal()
	' ���� Page�� Form Element���� Clear�Ѵ�. 
	Call ggoOper.ClearField(Document, "1")        '��: Condition field clear

	frm1.txtFromDt.text = EndDate 
	frm1.txtToDt.text = EndDate
	
	Call ggoOper.FormatDate(frm1.txtFromDt, parent.gDateFormat,2)
	Call ggoOper.FormatDate(frm1.txtToDt, parent.gDateFormat, 2)
	
	
End Sub


Function CookiePage(ByVal flgs)

	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		WriteCookie CookieSplit , frm1.txtEmp_no.Value
	ElseIf flgs = 0 Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		frm1.txtEmp_no.value =  strTemp

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
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream       = Frm1.txtFromDt.text & parent.gColSep                                           'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtToDt.text & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtdeduction.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtbizareacd.Value & parent.gColSep

End Sub        
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
'version
	'IntRetCD = CommonQueryRs(" minor_nm "," b_minor "," major_cd='H0130' And minor_cd = '" & iDx & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr  
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================

Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			.Row = intRow

			
		Next
	End With

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
	    
	    Dim strMaskYM
	
		strMaskYM = parent.gDateFormatYYYYMM
	
		strMaskYM = Replace(strMaskYM,"YYYY"      ,"9999")
		strMaskYM = Replace(strMaskYM,"YY"        ,"99")
		strMaskYM = Replace(strMaskYM,"MM"        ,"99")
		strMaskYM = Replace(strMaskYM,parent.gComDateType,"X")
	    
	    .ReDraw = false
        .MaxCols = C_deduction_DESC + 1												'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>
	    .Col = .MaxCols															'������Ʈ�� ��� Hidden Column%>
        .ColHidden = True
        .MaxRows = 0        
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData     
  	   
   		Call GetSpreadColumnPos("A")  	
	    Call AppendNumberPlace("6","2","0")
	    
	
		ggoSpread.SSSetMask     C_YMD,				"�Ű���",				8,2, strMaskYM      
	    ggoSpread.SSSetEdit     C_BIZ_AREA_CD,		"���ݽŰ�����",		12,,, 20,1
	    ggoSpread.SSSetButton   C_BIZ_AREA_POP
	    ggoSpread.SSSetEdit     C_BIZ_AREA_NM,		"���ݽŰ������",		18,,, 30,1 
	    ggoSpread.SSSetEdit     C_deduction_type,	"�Ұ�������",			10,,, 20,1
	    ggoSpread.SSSetButton   C_deduction_type_POP
	    ggoSpread.SSSetEdit     C_deduction_type_NM,"�Ұ���������",			15,,, 30,1  
		ggoSpread.SSSetFloat    C_TAX_CNT,			"�ż�",					6, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec,,,"Z" 
		ggoSpread.SSSetFloat    C_deduction_amt,	"���ް���",				12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_VAT_AMT,			"���Լ���",				12, parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,      parent.gComNum1000,     parent.gComNumDec,,,"Z"
        ggoSpread.SSSetEdit     C_deduction_DESC,	"���",					25,,, 30,1 
	    
	    
	    .ReDraw = true
	  
	  ' Call ggoSpread.SSSetColHidden(C_P_CNSLT_CD,C_P_CNSLT_CD,True)
        
       Call SetSpreadLock 
    
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
			 
			C_YMD					= iCurColumnPos(1)	 
			C_BIZ_AREA_CD			= iCurColumnPos(2)
			C_BIZ_AREA_POP			= iCurColumnPos(3)
			C_BIZ_AREA_NM			= iCurColumnPos(4)
			C_deduction_type		= iCurColumnPos(5)
			C_deduction_type_POP	= iCurColumnPos(6)
			C_deduction_type_NM		= iCurColumnPos(7)
			C_TAX_CNT				= iCurColumnPos(8)
			C_deduction_amt			= iCurColumnPos(9)
			C_VAT_AMT				= iCurColumnPos(10)
			C_deduction_DESC		= iCurColumnPos(11)
		 
    End Select
  
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
       
      'ggoSpread.SSSetRequired	  C_YMD,				-1, -1
      'ggoSpread.SSSetRequired	  C_BIZ_AREA_CD,		-1, -1 
      'ggoSpread.SSSetRequired	  C_deduction_type,		-1, -1 
      
      
      ggoSpread.SSSetRequired	  C_TAX_CNT,			-1, -1 
      ggoSpread.SSSetRequired	  C_deduction_amt,		-1, -1 
      ggoSpread.SSSetRequired	  C_VAT_AMT,			-1, -1 
       
      ggoSpread.SSSetProtected	  C_YMD,				-1, -1
      ggoSpread.SSSetProtected	  C_BIZ_AREA_CD,		-1, -1
      ggoSpread.SSSetProtected	  C_deduction_type,		-1, -1
      ggoSpread.SSSetProtected	  C_BIZ_AREA_POP,		-1, -1
      ggoSpread.SSSetProtected	  C_deduction_type_POP,	-1, -1
      ggoSpread.SSSetProtected	  C_BIZ_AREA_NM,		-1, -1
      ggoSpread.SSSetProtected    C_deduction_type_NM,	-1, -1

       
		'ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1            
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
        
		ggoSpread.SSSetProtected		C_BIZ_AREA_NM			, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected		C_deduction_type_NM		, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected		C_YMD					, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected		C_BIZ_AREA_CD			, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected		C_BIZ_AREA_POP			, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected		C_deduction_type		, pvStartRow, pvEndRow
		'ggoSpread.SSSetProtected		C_deduction_type_POP	, pvStartRow, pvEndRow     
		ggoSpread.SSSetRequired			C_TAX_CNT				, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired			C_deduction_amt			, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired			C_VAT_AMT				, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired			C_YMD					, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired			C_BIZ_AREA_CD			, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired			C_deduction_type		, pvStartRow, pvEndRow
		
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
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
	Call ggoOper.LockField(Document, "N")											'��: Lock Field
       
    Call InitSpreadSheet                                                            'Setup the Spread sheet

    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")   

    Call SetToolbar("1100110100001111")										        '��ư ���� ���� 

    Call SetDefaultVal()
	

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

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '��: Initializes local global variables
    Call MakeKeyStream("X")

	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True 
                                                             '��: Processing is OK
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    
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
    
    FncDelete = True                                                             '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim strstartDt
    Dim strendDt
    dim strEntr_dt
    Dim lRow
    Dim strTmp
    
    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '��:There is no changed data. 
        Exit Function
    End If
      
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If


    Call MakeKeyStream("X")
 
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then    
		Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True                                                              '��: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()


End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(Byval pvRowcnt) 
    Dim IntRetCD
    Dim imRow

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
   
	With frm1
		.vspdData.focus
		.vspdData.ReDraw = False

		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow,imRow
		
		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow)
		
		.vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
		FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
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
    Call InitComboBox()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)                                         '��: ȭ�� ���� 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
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

    DbQuery = False
    
    Err.Clear                                                                        '��: Clear err status

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey                 '��: Next key tag
    End With
		
    If lgIntFlgMode = parent.OPMD_UMODE Then
    Else
    End If

	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel

	Dim strRes_no

    DbSave = False                                                          
   
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text
			
			
               Case ggoSpread.InsertFlag                                      '��: Insert

														  strVal = strVal & "C"						& parent.gColSep
														  strVal = strVal & lRow					& parent.gColSep
					.vspdData.Col = C_YMD				: strVal = strVal & Trim(.vspdData.Text)	& parent.gColSep			                                                    
                    .vspdData.Col = C_BIZ_AREA_CD		: strVal = strVal & Trim(.vspdData.Text)	& parent.gColSep
                    .vspdData.Col = C_deduction_type	: strVal = strVal & Trim(.vspdData.Text)	& parent.gColSep
                    .vspdData.Col = C_TAX_CNT			: strVal = strVal & Trim(.vspdData.value)	& parent.gColSep
                    .vspdData.Col = C_deduction_amt		: strVal = strVal & Trim(.vspdData.value)	& parent.gColSep
                    .vspdData.Col = C_VAT_AMT			: strVal = strVal & Trim(.vspdData.value)	& parent.gColSep
                    .vspdData.Col = C_deduction_DESC	: strVal = strVal & Trim(.vspdData.Text)	& parent.gRowSep                                      			                                        
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '��: Update
                strVal = strVal & "U"  & parent.gColSep
                                                    strVal = strVal & lRow							& parent.gColSep                                                    
                    .vspdData.Col = C_YMD				: strVal = strVal & Trim(.vspdData.Text)	& parent.gColSep			                                                    
                    .vspdData.Col = C_BIZ_AREA_CD		: strVal = strVal & Trim(.vspdData.Text)	& parent.gColSep
                    .vspdData.Col = C_deduction_type	: strVal = strVal & Trim(.vspdData.Text)	& parent.gColSep
                    .vspdData.Col = C_TAX_CNT			: strVal = strVal & Trim(.vspdData.value)	& parent.gColSep
                    .vspdData.Col = C_deduction_amt		: strVal = strVal & Trim(.vspdData.value)	& parent.gColSep
                    .vspdData.Col = C_VAT_AMT			: strVal = strVal & Trim(.vspdData.value)	& parent.gColSep
                    .vspdData.Col = C_deduction_DESC	: strVal = strVal & Trim(.vspdData.Text)	& parent.gRowSep         
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '��: Delete

														  strDel = strDel & "D"						& parent.gColSep
														  strDel = strDel & lRow					& parent.gColSep                    
					.vspdData.Col = C_YMD				: strDel = strDel & Trim(.vspdData.Text)	& parent.gColSep			                                                    
                    .vspdData.Col = C_BIZ_AREA_CD		: strDel = strDel & Trim(.vspdData.Text)	& parent.gColSep
                    .vspdData.Col = C_deduction_type	: strDel = strDel & Trim(.vspdData.Text)	& parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

       .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal  	   

	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	

	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
 '   Dim IntRetCd
    
'    FncDelete = False                                                      '��: Processing is NG
    
'    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
 '       Call DisplayMsgBox("900002","X","X","X")                                '��:
  '      Exit Function
   ' End If
    '
'    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
'	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
'		Exit Function	
'	End If
 '   
	'Call DisableToolBar(parent.TBC_DELETE)
  '  If DbDelete = False Then
	'	Call RestoreToolBar()
     '   Exit Function
    'End If
    
    'FncDelete = True                                                        '��: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
    Call InitData()
	'Call SetToolbar("110011110011111")
	Call SetToolbar("110011110001111")

	Frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     
    Call InitVariables															'��: Initializes local global variables
	call MainQuery()
	
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
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

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("1101111111")       

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
	frm1.vspdData.Row = Row
	
End Sub
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim intIndex
    
    DIM strTemp
    
    DIM MinorCd,biz_area_cd,StrWhere
    
    
   	 	
	SELECT CASE COL

		   case C_deduction_type
				frm1.vspdData.Row = frm1.vspdData.ActiveRow
                Frm1.vspdData.Col = C_deduction_type
                MinorCd = frm1.vspddata.value
				Call CommonQueryRs("  MINOR_NM "," b_minor "," MAJOR_CD = 'a3001' and MINOR_Cd = " & FilterVar(Ucase(Trim(MinorCd)),"''","S")  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                frm1.vspdData.Col = C_deduction_type_nm
                frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
         
                If Trim(Replace(lgF0,Chr(11),"")) = "" Then
					frm1.vspdData.Row = frm1.vspdData.ActiveRow
					Frm1.vspdData.Col = C_deduction_type
                    frm1.vspdData.Value = ""
					frm1.vspdData.Row = frm1.vspdData.ActiveRow
					Frm1.vspdData.Col = C_deduction_type_nm
                    frm1.vspdData.Value = ""
				'Call  DisplayMsgBox("800054","X","X","X")	
                End If	
			
			
			
		   case C_biz_area_cd
				frm1.vspdData.Row = frm1.vspdData.ActiveRow
                Frm1.vspdData.Col = C_biz_area_cd
                biz_area_cd = frm1.vspddata.value
				Call CommonQueryRs(" tax_biz_area_nm "," b_tax_biz_area "," tax_biz_area_cd = " & FilterVar(Ucase(Trim(biz_area_cd)),"''","S")  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                frm1.vspdData.Col = C_biz_area_nm
                frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
         
                If Trim(Replace(lgF0,Chr(11),"")) = "" Then
					frm1.vspdData.Row = frm1.vspdData.ActiveRow
					Frm1.vspdData.Col = C_biz_area_cd
                    frm1.vspdData.Value = ""
					frm1.vspdData.Row = frm1.vspdData.ActiveRow
					Frm1.vspdData.Col = C_biz_area_nm
                    frm1.vspdData.Value = ""
				'Call  DisplayMsgBox("800054","X","X","X")	
                End If
			
    End SELECT
    
 	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
		If UNICDbl(frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
		     Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
		End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
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

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
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

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
'Private Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
'End Sub

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

Function OpenPopUp(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strSelect,strFrom,strWhere
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6	
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    Select Case iWhere
        Case 0
			arrParam(0) = "���ݽŰ�����"									' �˾� ��Ī 
			arrParam(1) = "b_tax_biz_area" 
			arrParam(2) = Trim(frm1.txtbizareacd.Value)						' Code Condition
			arrParam(3) = ""												' Where Condition
			arrParam(4) = ""
			arrParam(5) = "���ݽŰ������ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "tax_biz_area_CD"									' Field��(0)
			arrField(1) = "tax_biz_area_NM"	

			arrHeader(0) = "���ݽŰ������ڵ�"									' Header��(0)
			arrHeader(1) = "���ݽŰ������"									' Header��(1)
	           
        Case 1

	        arrParam(0) = "�Ұ�������"									' �˾� ��Ī 
			arrParam(1) = " b_minor " 
			arrParam(2) = Trim(frm1.txtdeduction.Value)						' Code Condition
			arrParam(3) = ""												' Where Condition
			arrParam(4) = " major_cd = 'a3001' "
			arrParam(5) = "�����ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "minor_cd"									' Field��(0)
			arrField(1) = "minor_nm"	

			arrHeader(0) = "�Ұ��������ڵ�"									' Header��(0)
			arrHeader(1) = "�Ұ���������"	
			
			
		Case 2

	        arrParam(0) = "���ݽŰ�����"									' �˾� ��Ī 
			arrParam(1) = " b_tax_biz_area " 
			frm1.vspddata.row = frm1.vspddata.activerow
			frm1.vspddata.col = C_biz_area_CD
			
			arrParam(2) = Trim(frm1.vspddata.Value)						' Code Condition
			arrParam(3) = ""												' Where Condition
			arrParam(4) = ""
			arrParam(5) = "���ݽŰ������ڵ�"									' �����ʵ��� �� ��Ī 

			arrField(0) = "tax_biz_area_CD"									' Field��(0)
			arrField(1) = "tax_biz_area_NM"	

			arrHeader(0) = "���ݽŰ������ڵ�"									' Header��(0)
			arrHeader(1) = "���ݽŰ������"		
		
		Case 3

	        arrParam(0) = "�Ұ�������"									' �˾� ��Ī 
			arrParam(1) = " b_minor " 
			frm1.vspddata.row = frm1.vspddata.activerow
			frm1.vspddata.col = C_deduction_type
			
			arrParam(2) = Trim(frm1.vspddata.Value)						' Code Condition
			arrParam(3) = ""												' Where Condition
			arrParam(4) = " major_cd = 'a3001' "
			arrParam(5) = "�Ұ�������"									' �����ʵ��� �� ��Ī 

			arrField(0) = "minor_cd"									' Field��(0)
			arrField(1) = "minor_NM"	

			arrHeader(0) = "�Ұ�������"									' Header��(0)
			arrHeader(1) = "�Ұ���������"			
	   
	    			    	
	End Select         

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If	

End Function

'------------------------------------------  SetReturnVal()  ---------------------------------------------
'	Name : SetReturnVal()
'	Description : Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetReturnVal(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere 
		    
		    Case 0
			  .txtbizareacd.value = arrRet(0)
			  .txtbizareanm.value = arrRet(1)  
			
		    Case 1
			  .txtdeduction.value = arrRet(0)
			  .txtdeductionnm.value = arrRet(1) 
			  
			Case 2
			  .vspddata.row = .vspddata.ActiveRow
			  .vspddata.col = C_biz_area_cd
			  .vspddata.value = arrRet(0)
			  .vspddata.col = C_biz_area_NM
			  .vspddata.value = arrRet(1) 
			
			Case 3
			  .vspddata.row = .vspddata.ActiveRow
			  .vspddata.col = C_deduction_type
			  .vspddata.value = arrRet(0)
			  .vspddata.col = C_deduction_type_nm
			  .vspddata.value = arrRet(1)      
			
		End Select
	End With
End Function

Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtFromDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtFromDt.Focus       
    End If
End Sub

'========================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
       frm1.txtToDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtToDt.Focus       
    End If
End Sub

'========================================================================================
Sub txtFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtToDt.focus
	   Call FncQuery
	End If   
End Sub

'========================================================================================
Sub txtToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
		frm1.txtFromDt.focus
	   Call FncQuery
	End If   
End Sub

Function txtbizareaCd_Onchange()
    Dim IntRetCd, strWhere
    
    IF Trim(frm1.txtbizareaCd.value) = "" Then
        frm1.txtbizareaCd.value        = ""
        frm1.txtbizareanm.value    = ""
    Else
        
        strWhere =  "  A.TAX_BIZ_AREA_CD = " & FilterVar(Ucase(Trim(frm1.txtbizareaCd.value)),"''","S")
        
        IntRetCd = CommonQueryRs( " DISTINCT A.TAX_BIZ_AREA_CD, A.TAX_BIZ_AREA_NM  " ," B_TAX_BIZ_AREA A ", strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        If IntRetCd = false then
            Call DisplayMsgBox("800054", "X", "X", "X")
            frm1.txtbizareaCd.value    = ""
            frm1.txtbizareaNM.value    = ""
            frm1.txtbizareaCd.focus
        ELSE    
            frm1.txtbizareaCd.Value	= Trim(Replace(lgF0,Chr(11),""))
            frm1.txtbizareaNM.Value	= Trim(Replace(lgF1,Chr(11),""))
	    End If
    End If
    
End Function

Function txtdeduction_Onchange()
    Dim IntRetCd, strWhere
    
    IF Trim(frm1.txtdeduction.value) = "" Then
        frm1.txtdeduction.value        = ""
        frm1.txtdeductionnm.value    = ""
    Else
        
        strWhere =  "  major_cd = 'A3001' and A.minor_Cd = " & FilterVar(Ucase(Trim(frm1.txtdeduction.value)),"''","S")
        
        IntRetCd = CommonQueryRs( " DISTINCT A.minor_cd, A.minor_NM  " ," B_minor A ", strWhere  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        
        If IntRetCd = false then
            Call DisplayMsgBox("800054", "X", "X", "X")
            frm1.txtdeduction.value    = ""
            frm1.txtdeductionnm.value    = ""
            frm1.txtdeduction.focus
        ELSE    
            frm1.txtdeduction.Value	= Trim(Replace(lgF0,Chr(11),""))
            frm1.txtdeductionnm.Value	= Trim(Replace(lgF1,Chr(11),""))
	    End If
    End If
    
End Function


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
   
		If Row > 0 Then   
			Select Case Col
   
			Case C_deduction_type_pop					
			    .Col = C_deduction_type
			    .Row = Row
			    Call OpenPopup(3)		
			Case C_biz_area_pop							
			    .Col = C_biz_area_cd
			    .Row = Row
			    Call OpenPopup(2)		
			End Select

   		End If		    

	End With
	
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���Լ��׺Ұ����а��ٰŵ��</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
						<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>    	            
    	            <TD HEIGHT=20 WIDTH=100%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
								<TD CLASS="TD5" NOWRAP>�Ű���</TD>
								<TD CLASS="TD6" NOWRAP><OBJECT classid=<%=gCLSIDFPDT%> NAME="txtFromDt" CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT="���ۿ�����" id=fpDateTime1></OBJECT>&nbsp;~&nbsp;
											           <OBJECT classid=<%=gCLSIDFPDT%> NAME="txtToDt"   CLASS=FPDTYYYYMM tag="12" Title="FPDATETIME" ALT="���Ό����" id=fpDateTime2></OBJECT></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>�Ұ�������</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtdeduction" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="�Ұ��������ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(1)">
									                       <INPUT TYPE=TEXT NAME="txtdeductionNm" ALT="�Ұ���������" SIZE=25 tag="14"></TD>
			    	        	<TD CLASS="TD5" NOWRAP>���ݽŰ�����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtbizareacd" SIZE=13 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag="11XXXU" ALT="���ݽŰ������ڵ�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenPopUp(0)">
									                       <INPUT TYPE=TEXT NAME="txtbizareanm" ALT="���ݽŰ������" SIZE=25 tag="14"></TD>
			            	</TR>
			            	
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vaSpread1>
										<PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0">
									</OBJECT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>	
					<TD>
						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

