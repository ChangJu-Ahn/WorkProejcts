<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: 
*  3. Program ID           	: c1906ma1_ko441
*  4. Program Name         	: c1906ma1_ko441
*  5. Program Desc         	: ������ڽ�Ʈ���͵�� 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/05/
*  8. Modified date(Last)  	: 2003/06/11
*  9. Modifier (First)     	: mok young bin
* 10. Modifier (Last)     	: Lee SiNa
* 11. Comment              	:
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

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit 

Dim iDBSYSDate
iDBSYSDate = "<%=GetSvrDate%>"

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "c4021mb1_ko441.asp"                                    'Biz Logic ASP 
Const BIZ_PGM_ID1 = "c4021mb2_ko441.asp"                                       'Biz Logic ASP  

Const C_SHEETMAXROWS    =   21	                                      '�� ȭ�鿡 �������� �ִ밹��*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          
Dim lsInternal_cd

Dim   C_EMP_NO		
Dim   C_EMP_NO_POP	
Dim   C_EMP_NAME
Dim   C_ORG_CHANGE_ID
Dim   C_DEPT_CD
Dim   C_DEPT_NM
Dim   C_DIR_INDIR	
Dim   C_DIR_INDIR_POP
Dim   C_DIR_INDIR_NM
Dim   C_COST_CD	
Dim   C_COST_CD_POP
Dim   C_COST_NM
Dim   C_BIZ_AREA_CD
Dim   C_BIZ_AREA_POP
Dim   C_BIZ_AREA_NM

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
Sub initSpreadPosVariables()  
    C_EMP_NO			= 1
    C_EMP_NO_POP			= 2
    C_EMP_NAME			= 3	
    C_ORG_CHANGE_ID		= 4	
    C_DEPT_CD			= 5	
    C_DEPT_NM			= 6
    C_DIR_INDIR			= 7	
    C_DIR_INDIR_POP		= 8
    C_DIR_INDIR_NM		= 9
    C_COST_CD			= 10
    C_COST_CD_POP		= 11
    C_COST_NM			= 12
    C_BIZ_AREA_CD		= 13
    C_BIZ_AREA_POP		= 14
    C_BIZ_AREA_NM		= 15
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
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom(iDBSYSDate,Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtYyyymm.focus			'��� default value setting
	
	frm1.txtYyyymm.Year = strYear 		 '����� default value setting
	frm1.txtYyyymm.Month = strMonth	

End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream       = left(Frm1.txtYyyymm.Text,4) & right(Frm1.txtYyyymm.Text,2) & parent.gColSep                                           'You Must append one character(parent.gColSep)
	lgKeyStream       = lgKeyStream & Frm1.txtDept_cd.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtCOST_CENTER_CD.Value & parent.gColSep
	lgKeyStream       = lgKeyStream & Frm1.txtEmp_no.Value & parent.gColSep

End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0047", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ' ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_WK_TYPE_CD
    ' ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_WK_TYPE         ''''''''DB���� �ҷ� gread���� 

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0047", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ' Call SetCombo2(frm1.cboWk_type,iCodeArr, iNameArr,Chr(11))                  ''''''''DB���� �ҷ� condition���� 
End Sub
'========================================================================================================
' Name : InitSpreadComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitSpreadComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0047", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_WK_TYPE_CD
    ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_WK_TYPE         ''''''''DB���� �ҷ� gread���� 
End Sub
'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

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
  		.MaxCols = C_BIZ_AREA_NM + 1                                                         ' ��:��: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ��:��: Hide maxcols
       .ColHidden = True                                                            ' ��:��:
       .MaxRows = 0	
		Call GetSpreadColumnPos("A")  
       
		ggoSpread.SSSetEdit       C_EMP_NO,   "���",          10,,, 13,2    	
		ggoSpread.SSSetButton     C_EMP_NO_POP 
		ggoSpread.SSSetEdit       C_EMP_NAME,   "�����",        10,,, 30,2 	
		ggoSpread.SSSetEdit       C_ORG_CHANGE_ID,   "��������ID",    10,,, 6,2      	
		ggoSpread.SSSetEdit       C_DEPT_CD,   "�μ��ڵ�",      8,,, 10,2 	
		ggoSpread.SSSetEdit       C_DEPT_NM,   "�μ���",        16,,, 40,2 	
		ggoSpread.SSSetEdit       C_DIR_INDIR,   "��������",      8,,, 4,2	
		ggoSpread.SSSetButton     C_DIR_INDIR_POP 
		ggoSpread.SSSetEdit       C_DIR_INDIR_NM,   "���и�",    8,,, 8,2 	
		ggoSpread.SSSetEdit       C_COST_CD,   "�ڽ�Ʈ����",    8,,, 10,2 	
		ggoSpread.SSSetButton     C_COST_CD_POP 
		ggoSpread.SSSetEdit       C_COST_NM,   "�ڽ�Ʈ���͸�",  12,,, 40,2 
		ggoSpread.SSSetEdit       C_BIZ_AREA_CD,   "�����",        6,,, 10,2
		ggoSpread.SSSetButton     C_BIZ_AREA_POP
		ggoSpread.SSSetEdit       C_BIZ_AREA_NM,   "������",      12,,, 40,2 

        Call ggoSpread.MakePairsColumn(C_EMP_NO,C_EMP_NO_POP)
        Call ggoSpread.MakePairsColumn(C_DIR_INDIR,C_DIR_INDIR_POP)
        Call ggoSpread.MakePairsColumn(C_COST_CD,C_COST_CD_POP)        
        Call ggoSpread.MakePairsColumn(C_BIZ_AREA_CD,C_BIZ_AREA_POP)              
        
        'call ggoSpread.SSSetColHidden(C_ORG_CHANGE_ID,C_ORG_CHANGE_ID,True)
        
	   .ReDraw = true
	
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
            
		    C_EMP_NO			= iCurColumnPos(1)	
		    C_EMP_NO_POP		= iCurColumnPos(2)	
		    C_EMP_NAME			= iCurColumnPos(3)
		    C_ORG_CHANGE_ID		= iCurColumnPos(4)
		    C_DEPT_CD			= iCurColumnPos(5)	
		    C_DEPT_NM			= iCurColumnPos(6)
		    C_DIR_INDIR			= iCurColumnPos(7)	
		    C_DIR_INDIR_POP		= iCurColumnPos(8)	
		    C_DIR_INDIR_NM		= iCurColumnPos(9)
		    C_COST_CD			= iCurColumnPos(10)
		    C_COST_CD_POP		= iCurColumnPos(11)
		    C_COST_NM			= iCurColumnPos(12)
		    C_BIZ_AREA_CD		= iCurColumnPos(13)
		    C_BIZ_AREA_POP		= iCurColumnPos(14)
		    C_BIZ_AREA_NM		= iCurColumnPos(15)   
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False

      ggoSpread.SpreadLock        C_EMP_NO, -1, C_EMP_NO
      ggoSpread.SpreadLock        C_EMP_NO_POP, -1, C_EMP_NO_POP
      ggoSpread.SpreadLock        C_EMP_NAME, -1, C_EMP_NAME	
      ggoSpread.SpreadLock        C_ORG_CHANGE_ID, -1, C_ORG_CHANGE_ID
      ggoSpread.SpreadLock        C_DEPT_CD, -1, C_DEPT_CD
      ggoSpread.SpreadLock        C_DEPT_NM, -1, C_DEPT_NM
      	                                                 
      ggoSpread.SSSetRequired      C_DIR_INDIR	, -1, -1	
      'ggoSpread.SSSetRequired      C_DIR_INDIR_POP, -1, -1
      ggoSpread.SpreadLock         C_DIR_INDIR_NM, -1, C_DIR_INDIR_NM
      ggoSpread.SSSetRequired      C_COST_CD, -1, -1
      'ggoSpread.SSSetRequired      C_COST_CD_POP, -1, -1 
      ggoSpread.SpreadLock     	   C_COST_NM, -1, C_COST_NM
      ggoSpread.SSSetRequired      C_BIZ_AREA_CD, -1, -1 
      'ggoSpread.SSSetRequired      C_BIZ_AREA_POP, -1, -1 
      ggoSpread.SpreadLock     	   C_BIZ_AREA_NM, -1, C_BIZ_AREA_NM           
                                                                                                                                                                 
	  ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1       
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

		ggoSpread.SSSetRequired        	   C_EMP_NO, pvStartRow, pvEndRow		                        
		ggoSpread.SSSetProtected            C_EMP_NAME, pvStartRow, pvEndRow		           
		ggoSpread.SSSetProtected            C_ORG_CHANGE_ID, pvStartRow, pvEndRow	      
		ggoSpread.SSSetProtected            C_DEPT_CD, pvStartRow, pvEndRow		       
		ggoSpread.SSSetProtected            C_DEPT_NM, pvStartRow, pvEndRow		    
		ggoSpread.SSSetRequired            C_DIR_INDIR, pvStartRow, pvEndRow		
		ggoSpread.SSSetProtected            C_DIR_INDIR_NM, pvStartRow, pvEndRow	  
		ggoSpread.SSSetRequired            C_COST_CD, pvStartRow, pvEndRow		   
		ggoSpread.SSSetProtected            C_COST_NM, pvStartRow, pvEndRow		          
		ggoSpread.SSSetRequired            C_BIZ_AREA_CD, pvStartRow, pvEndRow	    
		ggoSpread.SSSetProtected            C_BIZ_AREA_NM, pvStartRow, pvEndRow	                  

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
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock Field
	Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat, 2)             
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' �ڷ����:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
   ' Call InitComboBox
    Call SetToolbar("1100110100101111")										        '��ư ���� ���� 
    
	Call CookiePage (0)
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
    
	ggoSpread.ClearSpreadData  										 '��: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    If txtEmp_no_Onchange() Then         'ENTER KEY �� ��ȸ�� ����� ����� CHECK �Ѵ� 
        Exit Function
    End if
    
    If txtDept_cd_Onchange() Then        'enter key �� ��ȸ�� �μ��ڵ带 check�� �ش���� ������ query����...
        Exit Function
    End if
        
    Call InitVariables                                                           '��: Initializes local global variables
    Call MakeKeyStream("X")
    
    Call SetSpreadLock                                   '�ڵ��Է¶� Ǯ���� �κ��� �ٽ� ��ȸ�Ҷ� Lock��Ŵ 
	
	'Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
       Exit Function
    End If                                                                 '��: Query db data
       
    FncQuery = True                                                              '��: Processing is OK
    
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
    dim lRow
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
    
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag
					.vspdData.Col = C_EMP_NO
					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					    Call DisplayMsgBox("800048","X","X","X")
						Exit Function
					end if
            End Select
        Next
	End With

    Call MakeKeyStream("X")
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
		Call RestoreToolBar()
       Exit Function
    End If				                                                    '��: Save db data
    
    FncSave = True                                                              '��: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    FncCopy = False           
    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With

	With Frm1.VspdData	    
           .Col  = C_EMP_NO
           .Row  = .ActiveRow
           .Text = ""
           .Col  = C_EMP_NAME
           .Row  = .ActiveRow
           .Text = ""         
    End With
	
    Set gActiveElement = document.ActiveElement   

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

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD,imRow,iRow
    
    On Error Resume Next         
    FncInsertRow = False
    
    if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End if
	With frm1
	    .vspdData.ReDraw = False
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    ggoSpread.InsertRow,imRow
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

	'For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1        
    '	.vspdData.Row = iRow	    
    '    .vspdData.col = C_DEPT_CD
    '    .vspdData.value = frm1.txtDept_Cd.value
    '    .vspdData.col = C_DEPT_NM
    '    .vspdData.value = frm1.txtDept_nm.value        
    'Next             
        Call initData()

	    
	   .vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
	If Err.number =0 Then
		FncInsertRow = True
	End if
	
End Function
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
    Call InitSpreadComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '��: Clear err status

	If LayerShowHide(1) = False then
    		Exit Function 
    End if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    End With
		
	
    If lgIntFlgMode = parent.OPMD_UMODE Then
    Else
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
    
    DbQuery = True
	Call SetToolbar("110011110011111")	    
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

    Dim iColSep, iRowSep
    Dim strCUTotalvalLen					'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����,�ű�] 
	Dim strDTotalvalLen						'���ۿ� ä������ ���� 102399byte�� �Ѿ� ���°��� üũ�ϱ����� ���� ����Ÿ ũ�� ����[����]
 	Dim iFormLimitByte						'102399byte
 	Dim objTEXTAREA							'������ HTML��ü(TEXTAREA)�� ��������� �ӽ� ���� 
 	Dim iTmpCUBuffer						'������ ���� [����,�ű�] 
	Dim iTmpCUBufferCount					'������ ���� Position
	Dim iTmpCUBufferMaxCount				'������ ���� Chunk Size
 	Dim iTmpDBuffer							'������ ���� [����] 
	Dim iTmpDBufferCount					'������ ���� Position
	Dim iTmpDBufferMaxCount					'������ ���� Chunk Size
    
    iColSep = parent.gColSep : iRowSep = parent.gRowSep 
 	
 	'�ѹ��� ������ ������ ũ�� ���� 
    iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT	
 	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT	
     
     '102399byte
     iFormLimitByte = parent.C_FORM_LIMIT_BYTE
     
     '������ �ʱ�ȭ 
    ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)			
 	ReDim iTmpDBuffer (iTmpDBufferMaxCount)				
 
 	iTmpCUBufferCount = -1 : iTmpDBufferCount = -1
 	
 	strCUTotalvalLen = 0 : strDTotalvalLen  = 0
	
    DbSave = False                                                          
    
    If LayerShowHide(1) = False then
    	Exit Function 
    End if

    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
          
               Case ggoSpread.InsertFlag                                      '��: Insert�߰� 
					strVal = ""
                                                    strVal = strVal & "C" & parent.gColSep 'array(0)
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & left(Frm1.txtYyyymm.Text,4) & right(Frm1.txtYyyymm.Text,2) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ORG_CHANGE_ID : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DIR_INDIR	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COST_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                                  
                    .vspdData.Col = C_BIZ_AREA_CD   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1 
                    
               Case ggoSpread.UpdateFlag                                      '��: Update
					strVal = ""
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & left(Frm1.txtYyyymm.Text,4) & right(Frm1.txtYyyymm.Text,2) & parent.gColSep                                                    
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ORG_CHANGE_ID : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DEPT_CD		: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DIR_INDIR	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_COST_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep                                  
                    .vspdData.Col = C_BIZ_AREA_CD   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '��: Delete
					strDel = ""               
                                                    strDel = strDel & "D" & parent.gColSep
                                                    strDel = strDel & lRow & parent.gColSep
                                                    strDel = strDel & left(Frm1.txtYyyymm.Text,4) & right(Frm1.txtYyyymm.Text,2) & parent.gColSep 
                    .vspdData.Col = C_EMP_NO      : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep	'������ key��								
                    lGrpCnt = lGrpCnt + 1
           End Select
           
           

			.vspdData.Col = 0
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag

			         If strCUTotalvalLen + Len(strVal) >  iFormLimitByte Then  '�Ѱ��� form element�� ���� Data �Ѱ�ġ�� ������ 
			                            
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

			   Case ggoSpread.DeleteFlag
			         If strDTotalvalLen + Len(strDel) >  iFormLimitByte Then   '�Ѱ��� form element�� ���� �Ѱ�ġ�� ������ 
			            Set objTEXTAREA   = document.createElement("TEXTAREA")
			            objTEXTAREA.name  = "txtDSpread"
			            objTEXTAREA.value = Join(iTmpDBuffer,"")
			            divTextArea.appendChild(objTEXTAREA)     
			          
			            iTmpDBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT              
			            ReDim iTmpDBuffer(iTmpDBufferMaxCount)
			            iTmpDBufferCount = -1
			            strDTotalvalLen = 0 
			         End If
			       
			         iTmpDBufferCount = iTmpDBufferCount + 1

			         If iTmpDBufferCount > iTmpDBufferMaxCount Then                         '������ ���� ����ġ�� ������ 
			            iTmpDBufferMaxCount = iTmpDBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
			            ReDim Preserve iTmpDBuffer(iTmpDBufferMaxCount)
			         End If   
			         
			         iTmpDBuffer(iTmpDBufferCount) =  strDel         
			         strDTotalvalLen = strDTotalvalLen + Len(strDel)
			         
			End Select
       Next
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	End With

    If iTmpCUBufferCount > -1 Then   ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   
	   divTextArea.appendChild(objTEXTAREA)
	End If   
	
	If iTmpDBufferCount > -1 Then    ' ������ ������ ó�� 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name = "txtDSpread"
	   objTEXTAREA.value = Join(iTmpDBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '��: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '��:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
	Call DisableToolBar(parent.TBC_DELETE)
    If DbDelete= False Then
		Call RestoreToolBar()
        Exit Function
    End If
    
    FncDelete = True                                                        '��: Processing is OK


End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
    Call InitData()
	Call SetToolbar("110011110011111")									
	frm1.vspdData.focus	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData
	Call RemovedivTextArea    	      
    Call InitVariables															'��: Initializes local global variables
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If
	
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If
End Function

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		
		.Row = Row
   
        Select Case Col
            Case C_WK_TYPE
                .Col = Col
                intIndex = .Value
				.Col = C_WK_TYPE_CD
				.Value = intIndex
            Case C_WK_TYPE_CD
                .Col = Col
                intIndex = .Value
				.Col = C_WK_TYPE
				.Value = intIndex
				
		End Select
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
' Name : OpenEmptName()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)
	Dim  yyyymmdd
	yyyymmdd = frm1.txtYyyymm.year & "-" & Right("0" & frm1.txtYyyymm.month , 2) & "-" & Right("0" & frm1.txtYyyymm.day , 2)		
		
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = "" 'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  
	Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_EMP_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  
	End If
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else 
			frm1.vspdData.Col = C_EMP_NO			
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SubSetCondEmp()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
    Dim strDept_cd, strDept_nm
    Dim strDept_cd2, strDept_nm2
    Dim strFg, strChange_dt, strEmp_no
    Dim  yyyymmdd , yyyymmdd2, IntRetCd2
    
	yyyymmdd = frm1.txtYyyymm.year & "-" & Right("0" & frm1.txtYyyymm.month , 2) & "-" & Right("0" & frm1.txtYyyymm.day , 2)
	yyyymmdd2 = frm1.txtYyyymm.year & Right("0" & frm1.txtYyyymm.month , 2) & Right("0" & frm1.txtYyyymm.day , 2)
  
	With frm1
	 		IntRetCd2 = CommonQueryRs(" max(orgid) "," horg_abs ", " orgdt <= " & FilterVar(yyyymmdd2 , "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)		    
	     	.vspdData.Col = C_ORG_CHANGE_ID
		 	.vspdData.text = Trim(Replace(lgF0,Chr(11),""))  
		 	
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_EMP_NAME
			.vspdData.Text = arrRet(1)
			strEmp_no = arrRet(0)

            Call CommonQueryRs(" DEPT_CD, DEPT_NM "," HAA010T "," EMP_NO =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            strDept_cd = Trim(Replace(lgF0,Chr(11),""))
            strDept_nm = Trim(Replace(lgF1,Chr(11),""))       	    

	       	'Frm1.vspdData.Col = C_CHANG_DT
            'strChange_dt = UNIConvDate(frm1.vspdData.Text)
            
	        Call CommonQueryRs(" count(*) "," hba010t a", " a.gazet_dt = (select MAX(gazet_dt) from hba010t " & _
                                                          " where gazet_dt <=  " & FilterVar(yyyymmdd , "''", "S") & "" & _
                                                            " and dept_cd is not null " &_
                                                            " and emp_no = a.emp_no) " &_
                               " and a.emp_no =  " & FilterVar(strEmp_no , "''", "S") & "" &_
                               " and a.dept_cd is not null ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            strFg = Trim(Replace(lgF0, Chr(11), ""))

            If IsNull(strFg) OR strFg = "" OR strFg = 0 Then
	       	    Frm1.vspdData.Col = C_DEPT_CD
	       	    Frm1.vspdData.value = strDept_cd
	       	    Frm1.vspdData.Col = C_DEPT_NM
	       	    Frm1.vspdData.value = strDept_nm
            Else
                Call CommonQueryRs(" dept_cd "," hba010t a", " a.gazet_dt = (select MAX(gazet_dt) from hba010t " & _
                                                             " where gazet_dt <=  " & FilterVar(yyyymmdd , "''", "S") & "" & _
                                                               " and dept_cd is not null " &_
                                                               " and emp_no = a.emp_no" & ")" &_
                                   " and a.emp_no =  " & FilterVar(strEmp_no , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                strDept_cd2 = Trim(Replace(lgF0, Chr(11), ""))

                Call CommonQueryRs(" dept_nm "," b_acct_dept ", " org_change_dt = (select MAX(org_change_dt) from b_acct_dept " & _
                                                                    " where org_change_dt <=  " & FilterVar(yyyymmdd , "''", "S") & ")" & _
                                               " and dept_cd =  " & FilterVar(strDept_cd2 , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                strDept_nm2 = Trim(Replace(lgF0, Chr(11), ""))

                If IsNull(strDept_nm2) OR strDept_nm2 = "" Then
	       	            Frm1.vspdData.Col = C_DEPT_CD
	       	            Frm1.vspdData.value = strDept_cd
	       	            Frm1.vspdData.Col = C_DEPT_NM
	       	            Frm1.vspdData.value = strDept_nm
                Else
	       	            Frm1.vspdData.Col = C_DEPT_CD
	       	            Frm1.vspdData.value = strDept_cd2
	       	            Frm1.vspdData.Col = C_DEPT_NM
	       	            Frm1.vspdData.value = strDept_nm2
                End If                                              
            End If


			 				 	            
			.vspdData.Col = C_EMP_NO			
			.vspdData.action =0
		End If
	End With
End Sub

'========================================================================================================
' Name : OpenDept
' Desc : �μ� POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	Dim  yyyymmdd
	yyyymmdd = frm1.txtYyyymm.year & "-" & Right("0" & frm1.txtYyyymm.month , 2) & "-" & Right("0" & frm1.txtYyyymm.day , 2)	
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then
		arrParam(0) = frm1.txtDept_cd.value			            '  Code Condition
	End If
    arrParam(1) = yyyymmdd  'frm1.txtChang_dt.Text
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		Select Case iWhere
		     Case "0"
               .txtDept_cd.value = arrRet(0)
               .txtDept_nm.value = arrRet(1)
               .txtInternal_cd.value = arrRet(2)
        End Select
	End With
End Function    



Function OpenCCCd(Byval iWhere)
	Dim arrRet
	Dim strWhere, strFrom
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	

	IsOpenPop = True
			
	arrParam(0) = "C/C"						' �˾� ��Ī
	arrParam(1) =" b_cost_center "					' TABLE ��Ī
	arrParam(2) = Trim(frm1.txtCOST_CENTER_CD.Value)	' Code Condition
	arrParam(3) =""										' Name Cindition
	arrParam(4) =""							' Where Condition
	arrParam(5) = "C/C"							' TextBox ��Ī
	
    arrField(0) ="ED10" & Parent.gColSep &  "cost_cd"					' Field��(0)
    arrField(1) = "ED31" & Parent.gColSep & "cost_nm"					' Field��(1)
    
    
    arrHeader(0) = "C/C"						' Header��(0)
    arrHeader(1) = "C/C��"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetCCCd(iWhere, arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtCOST_CENTER_CD.focus
	
End Function
   		
Function SetCCCd(Byval iWhere,  byval arrRet)

	With frm1
		If iWhere = 0 Then
			frm1.txtCOST_CENTER_CD.Value    = arrRet(0)		
			frm1.txtCOST_CENTER_NM.Value   = arrRet(1)
	Else
			.vspdData.Col = C_COST_CD
			.vspdData.text = arrRet(0) 
			.vspdData.Col = C_COST_NM 
			.vspdData.text = arrRet(1) 					
		End If	
	End With		
	
		
End Function


Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	Dim  yyyymmdd
	
	yyyymmdd = frm1.txtYyyymm.year & "-" & Right("0" & frm1.txtYyyymm.month , 2) & "-" & Right("0" & frm1.txtYyyymm.day , 2)		
	
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
                      
	    Case C_EMP_NO_POP
			arrParam(0) = "�λ���ȸ �˾�"			' �˾� ��Ī 
			arrParam(1) = "HAA010T"				 		' TABLE ��Ī 
			arrParam(2) = ""	    ' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = "(retire_dt is null or retire_dt >= " & FilterVar(yyyymmdd, "''", "S") & ")" & " and entr_dt <= " & FilterVar(yyyymmdd, "''", "S") ' Where Condition 
			arrParam(5) = "���"			    ' TextBox ��Ī 
			
			arrField(0) = "emp_no "					' Field��(0)
			arrField(1) = "name"				    ' Field��(1)
			arrField(2) = "dept_cd"					' Field��(2)
			arrField(3) = "dept_nm"					' Field��(3)
			arrField(4) = ""						' Field��(4)
			arrField(5) = ""						' Field��(5)
			arrField(6) = ""				    ' Field��(6)
			
			arrHeader(0) = "���"				' Header��(0)
			arrHeader(1) = "����"			    ' Header��(1)
			arrHeader(2) = "�μ��ڵ�"			' Header��(2)
			arrHeader(3) = "�μ���"				' Header��(3)
			arrHeader(4) = ""			    ' Header��(4)
			arrHeader(5) = ""			    ' Header��(5)
			arrHeader(6) = ""			' Header��(6)
			
	    Case C_DIR_INDIR_POP			
			arrParam(0) = "���������� �˾�"				' �˾� ��Ī 
	        arrParam(1) = "B_MINOR"				 		' TABLE ��Ī 
	        arrParam(2) = ""		   				 	' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = "MAJOR_CD = "  & FilterVar("H0071", "''", "S") 						' Where Condition
	        arrParam(5) = "����������"			    	' TextBox ��Ī 
	
            arrField(0) = "MINOR_CD"					' Field��(0)
            arrField(1) = "MINOR_NM"				    ' Field��(1)
            arrField(2) = ""				    		' Field��(2)
    
            arrHeader(0) = "����������"				' Header��(0)
            arrHeader(1) = "����������"			    ' Header��(1)
            arrHeader(2) = ""						' Header��(2)
			
	    Case C_BIZ_AREA_POP			
			arrParam(0) = "����� �˾�"				' �˾� ��Ī 
	        arrParam(1) = "B_BIZ_AREA"				 		' TABLE ��Ī 
	        arrParam(2) = ""		   				 	' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = "" 						' Where Condition
	        arrParam(5) = "�����"			    	' TextBox ��Ī 
	
            arrField(0) = "BIZ_AREA_CD"					' Field��(0)
            arrField(1) = "BIZ_AREA_NM"				    ' Field��(1)
            arrField(2) = ""				    		' Field��(2)
    
            arrHeader(0) = "������ڵ�"				' Header��(0)
            arrHeader(1) = "������"			    ' Header��(1)
            arrHeader(2) = ""						' Header��(2)						
           
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=500px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

End Function



Function SetCode(Byval arrRet, Byval iWhere)
	Dim IntRetCd
	Dim  yyyymmdd
	yyyymmdd = frm1.txtYyyymm.year & Right("0" & frm1.txtYyyymm.month , 2) & Right("0" & frm1.txtYyyymm.day , 2)
	
	With frm1
		Select Case iWhere
		    Case C_EMP_NO_POP
		    	.vspdData.Col = C_EMP_NO 
				.vspdData.text = arrRet(0) 
				.vspdData.Col = C_EMP_NAME
				.vspdData.text = arrRet(1) 
		    	.vspdData.Col = C_DEPT_CD
				.vspdData.text = arrRet(2) 
				.vspdData.Col = C_DEPT_NM
				.vspdData.text = arrRet(3) 
					    	
	 			IntRetCd = CommonQueryRs(" max(orgid) "," horg_abs ", " orgdt <= " & FilterVar(yyyymmdd , "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)		    
		     	.vspdData.Col = C_ORG_CHANGE_ID
			 	.vspdData.text = Trim(Replace(lgF0,Chr(11),""))  			
			 	
		    Case C_DIR_INDIR_POP
		    	.vspdData.Col = C_DIR_INDIR
				.vspdData.text = arrRet(0) 
				.vspdData.Col = C_DIR_INDIR_NM		
				.vspdData.text = arrRet(1) 		
		    Case C_BIZ_AREA_POP
		    	.vspdData.Col = C_BIZ_AREA_CD
				.vspdData.text = arrRet(0) 
				.vspdData.Col = C_BIZ_AREA_NM		
				.vspdData.text = arrRet(1) 						
															 				 	 								
        End Select
	End With
End Function
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col

    Case C_EMP_NO_POP
		Call OpenEmptName("1")           
    Case C_COST_CD_POP
		Call OpenCCCd(1)     
    Case C_DIR_INDIR_POP
        Call OpenCode("", C_DIR_INDIR_POP, Row)		        
    Case C_BIZ_AREA_POP
        Call OpenCode("", C_BIZ_AREA_POP, Row)		        
        						     
    End Select 
    
End Sub

Sub autoInsert_ButtonClicked(Byval ButtonDown)

	Call BtnDisabled(1)
	
    Dim strKeyStream
    Dim strVal
   ' Dim IntRetCD
   ' 
   '  
   ' Dim strInternalCd
   ' Dim strInternalCd2
   ' Dim strChangDt   Dim strEmpNo
 
    
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Call BtnDisabled(0)
			Exit sub
		End If
    End If
    
    If Not chkField(Document, "1") Then									         '��: This function check required field
		Call BtnDisabled(0)
		Exit sub
    End If
         
    
    ggoSpread.Source = Frm1.vspdData    
	'ggoSpread.ClearSpreadData  	
	'frm1.vspdData.MaxRows = 0
	
    strKeyStream = left(Frm1.txtYyyymm.Text,4) & right(Frm1.txtYyyymm.Text,2) & parent.gColSep                                           'You Must append one character(parent.gColSep)
	strKeyStream = strKeyStream & Frm1.txtDept_cd.Value & parent.gColSep
	strKeyStream = strKeyStream & Frm1.txtCOST_CENTER_CD.Value & parent.gColSep
	strKeyStream = strKeyStream & Frm1.txtEmp_no.Value & parent.gColSep
	

    
    With Frm1
    	strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001                          'mb2 �ڵ��Է�......						         
        strVal = strVal     & "&txtKeyStream="       & strKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    End With
    Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
	Call BtnDisabled(0)
End Sub

Sub DBAutoQueryOk()
    Dim lRow
    With Frm1
        .vspdData.ReDraw = false
        ggoSpread.Source = .vspdData
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0

            .vspdData.Text = ggoSpread.InsertFlag
        Next

     ' ggoSpread.SpreadLock C_CHANG_DT, -1,C_CHANG_DT
      ggoSpread.SpreadLock C_EMP_NO, -1,C_EMP_NO
      ggoSpread.SpreadLock C_EMP_NO_POP, -1,C_EMP_NO_POP

    .vspdData.ReDraw = TRUE
    'ggoSpread.ClearSpreadData "T"     
    
          
     
    End With    
    lgStrPrevKey = ""
    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCd
    Dim strName
    Dim strDept_cd, strDept_nm
    Dim strDept_cd2, strDept_nm2
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strFg, strChange_dt
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_EMP_NO
            iDx = Frm1.vspdData.value
   	        Frm1.vspdData.Col = C_EMP_NO
    
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_NAME
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_DEPT_CD
                Frm1.vspdData.value = ""
  	            Frm1.vspdData.Col = C_DEPT_NM
                Frm1.vspdData.value = ""
            Else
	            IntRetCd = FuncGetEmpInf2(iDx,lgUsrIntCd,strName,strDept_nm, strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	            if  IntRetCd < 0 then
	                if  IntRetCd = -1 then
                		Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
                    else
                        Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
                    end if
  	                Frm1.vspdData.Col = C_NAME
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_DEPT_CD
                    Frm1.vspdData.value = ""
  	                Frm1.vspdData.Col = C_DEPT_NM
                    Frm1.vspdData.value = ""
                    vspdData_Change = true
                Else
		       	    Frm1.vspdData.Col = C_NAME
		       	    Frm1.vspdData.value = strName

                    Call CommonQueryRs(" DEPT_CD, DEPT_NM "," HAA010T "," EMP_NO =  " & FilterVar(iDx , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                    strDept_cd = Trim(Replace(lgF0,Chr(11),""))
                    strDept_nm = Trim(Replace(lgF1,Chr(11),""))       	    

		       	    Frm1.vspdData.Col = C_CHANG_DT
                    strChange_dt = UNIConvDate(frm1.vspdData.Text)
 		            Call CommonQueryRs(" count(*) "," hba010t a", " a.gazet_dt = (select MAX(gazet_dt) from hba010t " & _
	                                                              " where gazet_dt <=  " & FilterVar(strChange_dt , "''", "S") & "" & _
	                                                                " and dept_cd is not null " &_
	                                                                " and emp_no = a.emp_no) " &_
                                       " and a.emp_no =  " & FilterVar(iDx , "''", "S") & "" &_
                                       " and a.dept_cd is not null ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                    strFg = Trim(Replace(lgF0, Chr(11), ""))

                    If IsNull(strFg) OR strFg = "" OR strFg = 0 Then
		       	       Frm1.vspdData.Col = C_DEPT_CD
		       	       Frm1.vspdData.value = strDept_cd
		       	       Frm1.vspdData.Col = C_DEPT_NM
		       	       Frm1.vspdData.value = strDept_nm
	                Else
	                    Call CommonQueryRs(" dept_cd "," hba010t a", " a.gazet_dt = (select MAX(gazet_dt) from hba010t " & _
	                                                                  " where gazet_dt <=  " & FilterVar(strChange_dt , "''", "S") & "" & _
   	                                                                    " and dept_cd is not null " &_
  	                                                                    " and emp_no = a.emp_no" & ")" &_
                                            " and a.emp_no =  " & FilterVar(iDx , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                        strDept_cd2 = Trim(Replace(lgF0, Chr(11), ""))

	                    Call CommonQueryRs(" dept_nm "," b_acct_dept ", " org_change_dt = (select MAX(org_change_dt) from b_acct_dept " & _
	                                                                    " where org_change_dt <=  " & FilterVar(strChange_dt , "''", "S") & ")" & _
                                                       " and dept_cd =  " & FilterVar(strDept_cd2 , "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                        strDept_nm2 = Trim(Replace(lgF0, Chr(11), ""))

                        If IsNull(strDept_nm2) OR strDept_nm2 = "" Then
		       	            Frm1.vspdData.Col = C_DEPT_CD
		       	            Frm1.vspdData.value = strDept_cd
		       	            Frm1.vspdData.Col = C_DEPT_NM
		       	            Frm1.vspdData.value = strDept_nm
                        Else
		       	            Frm1.vspdData.Col = C_DEPT_CD
		       	            Frm1.vspdData.value = strDept_cd2
		       	            Frm1.vspdData.Col = C_DEPT_NM
		       	            Frm1.vspdData.value = strDept_nm2
                        End If
                    End If		       	    
                End if 
            End if 
            
    End Select    
             
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function

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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub    

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
            Exit Function      
        Else
            frm1.txtName.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
'   Event Name : txtFr_dept_cd_change
'   Event Desc :
'========================================================================================================
Function txtDept_cd_Onchange()
    Dim IntRetCd
    Dim strDept_nm

    If frm1.txtDept_cd.value = "" Then
		frm1.txtDept_nm.value = ""
		frm1.txtInternal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtDept_cd.value,UNIConvDate(frm1.txtChang_dt.Text),lgUsrIntCd,strDept_nm,lsInternal_cd)
        
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   '�μ��ڵ������� ��ϵ��� ���� �ڵ��Դϴ�.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' �ڷ������ �����ϴ�.
            end if
		    frm1.txtDept_nm.value = ""
		    frm1.txtInternal_cd.value = ""
            lsInternal_cd = ""
            frm1.txtDept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtDept_cd_Onchange = true
            Exit Function      
        else
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtInternal_cd.value = lsInternal_cd
        end if
    End if  
End Function

'=======================================
'   Event Name : txtIntchng_yymm_dt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================

Sub txtYyyymm_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYyyymm.Action = 7
        frm1.txtYyyymm.focus
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

'=======================================================================================================
'   Event Name : txtChang_dt_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtChang_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub
'========================================================================================
' Function Name : RemovedivTextArea
'========================================================================================
Function RemovedivTextArea()
	Dim i
	For i = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>���κ��ڽ�Ʈ��Ÿ����</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
						<TD CLASS="TD5" NOWRAP>���޿���</TD>
						<TD CLASS="TD6" NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtYyyymm NAME="txtYyyymm" CLASS=FPDTYYYYMMDD tag="12X1" Alt="���ۿ�" Title="FPDATETIME"></OBJECT>');</SCRIPT>
						</TD>
				    	 <TD CLASS=TD5 NOWRAP>�μ��ڵ�</TD>              
			             <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="�μ��ڵ�" TYPE="Text" SiZE=13 MAXLENGTH=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenDept(0)">
			                                  <INPUT NAME="txtDept_nm" ALT="�μ��ڵ��" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14">
						                      <INPUT NAME="txtInternal_cd" ALT="�����ڵ�" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="14"></TD>
			           </TR>
		               <TR>		
						<TD CLASS="TD5" NOWRAP>�ڽ�Ʈ��Ÿ</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtCOST_CENTER_CD" MAXLENGTH="10" SIZE=13 ALT ="�ڽ�Ʈ��Ÿ �ڵ�" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCCCd(0)" > <INPUT NAME="txtCOST_CENTER_NM" MAXLENGTH="20" SIZE=30 STYLE="TEXT-ALIGN:left" ALT ="�ڽ�Ʈ��Ÿ��" tag="14"></TD>		               			           
			             <!--TD CLASS=TD5 NOWRAP>����ٹ���</TD>
						 <TD CLASS="TD6"><SELECT NAME="cboWk_type" ALT="����ٹ���" CLASS ="cbonormal" TAG="11N"><OPTION VALUE=""></OPTION></SELECT></TD !-->
				         <TD CLASS=TD5 NOWRAP>���</TD>
				     	 <TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="���" TYPE="Text" SiZE=13 MAXLENGTH=13  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmptName('0')">
				     	                      <INPUT NAME="txtName" ALT="����" TYPE="Text" SiZE=20 MAXLENGTH=30  tag="14"></TD>
					   </TR>
					  </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
	
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/c4021ma1_ko441_vaSpread_vspdData.js'></script>
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
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD><BUTTON NAME="btnCb_autoisrt" CLASS="CLSMBTN" ONCLICK="VBScript: autoInsert_ButtonClicked('1')" flag=1>�ڵ��Է�</BUTTON></TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
