<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: Multi Sample
*  3. Program ID           	: H8006ma1
*  4. Program Name         	: H8006ma1
*  5. Program Desc         	: �ұޱ޿����޺���ȸ 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2001/04/18
*  8. Modified date(Last)  	: 2003/06/13
*  9. Modifier (First)     	: TGS �ֿ�ö 
* 10. Modifier (Last)      	: Lee SiNa
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h8006mb1.asp"						           '��: Biz Logic ASP Name
Const C_SHEETMAXROWS    = 21	                                      '��: Visble row
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgOldRow

Dim C_PAY_GRD1
Dim C_EMP_COUNT
Dim C_ORIGINAL_BASE_AMT
Dim C_ORIGINAL_OTHER_AMT
Dim C_RAISE_BASE_AMT
Dim C_RAISE_OTHER_AMT
Dim C_RETRO_BASE_AMT
Dim C_RETRO_OTHER_AMT

Dim C_PAY_GRD12
Dim C_EMP_COUNT2
Dim C_ORIGINAL_BASE_AMT2
Dim C_ORIGINAL_OTHER_AMT2
Dim C_RAISE_BASE_AMT2
Dim C_RAISE_OTHER_AMT2
Dim C_RETRO_BASE_AMT2
Dim C_RETRO_OTHER_AMT2
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
 
    C_PAY_GRD1 = 1
    C_EMP_COUNT = 2
    C_ORIGINAL_BASE_AMT = 3
    C_ORIGINAL_OTHER_AMT = 4
    C_RAISE_BASE_AMT = 5
    C_RAISE_OTHER_AMT = 6
    C_RETRO_BASE_AMT = 7
    C_RETRO_OTHER_AMT = 8

    C_PAY_GRD12 = 1     	
    C_EMP_COUNT2 = 2    
    C_ORIGINAL_BASE_AMT2 = 3     
    C_ORIGINAL_OTHER_AMT2  = 4   
    C_RAISE_BASE_AMT2  = 5     
    C_RAISE_OTHER_AMT2  = 6     
    C_RETRO_BASE_AMT2 = 7 
    C_RETRO_OTHER_AMT2 = 8
    
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
	lgOldRow = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtpay_yymm_dt.Focus
	frm1.txtpay_yymm_dt.Year = strYear 		'��� default value setting
	frm1.txtpay_yymm_dt.Month = strMonth 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
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
Sub MakeKeyStream(pOpt)
    lgKeyStream  = frm1.txtpay_yymm_dt.year & Right("0" & frm1.txtpay_yymm_dt.month, 2)  & Parent.gColSep
    lgKeyStream  = lgKeyStream & Frm1.txtpay_grd1.value & Parent.gColSep
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
    Dim dblSum
    	
	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData2
        ggoSpread.UpdateRow 1
        frm1.vspdData2.Col = 0
        frm1.vspdData2.Text = "�հ�"

        frm1.vspdData2.Col = C_EMP_COUNT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_EMP_COUNT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_ORIGINAL_BASE_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_ORIGINAL_BASE_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_ORIGINAL_OTHER_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_ORIGINAL_OTHER_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_RAISE_BASE_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_RAISE_BASE_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_RAISE_OTHER_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_RAISE_OTHER_AMT, 1, .MaxRows, FALSE, -1, -1, "V")
        frm1.vspdData2.Col = C_RETRO_BASE_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_RETRO_BASE_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        frm1.vspdData2.Col = C_RETRO_OTHER_AMT2
        frm1.vspdData2.Text = FncSumSheet(frm1.vspdData,C_RETRO_OTHER_AMT, 1, .MaxRows,FALSE , -1, -1, "V")
        
    End With
    
    Call SetSpreadLock("B")

End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call initSpreadPosVariables()   'sbk 

    If pvSpdNo = "" OR pvSpdNo = "A" Then

	    With frm1.vspdData

            ggoSpread.Source = frm1.vspdData
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false
	
           .MaxCols   = C_RETRO_OTHER_AMT + 1                                                      ' ��:��: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ��:��: Hide maxcols
           .ColHidden = True                                                            ' ��:��:
 
           .MaxRows = 0
            ggoSpread.ClearSpreadData

            Call GetSpreadColumnPos("A") 'sbk

            Call AppendNumberPlace("6","15","0")

            ggoSpread.SSSetEdit C_PAY_GRD1           , "����", 11,,, 50,2
            ggoSpread.SSSetFloat C_EMP_COUNT         , "�ο�"            ,7,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_ORIGINAL_BASE_AMT , "�����޺б⺻����",17,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_ORIGINAL_OTHER_AMT, "�����޺б�Ÿ����",17,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	        ggoSpread.SSSetFloat C_RAISE_BASE_AMT    , "�λ�б⺻����"  ,16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	        ggoSpread.SSSetFloat C_RAISE_OTHER_AMT   , "�λ�б�Ÿ����"  ,16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	        ggoSpread.SSSetFloat C_RETRO_BASE_AMT    , "�ұ޺б⺻����"  ,17,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	        ggoSpread.SSSetFloat C_RETRO_OTHER_AMT   , "�ұ޺б�Ÿ����"  ,16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

	       .ReDraw = true
	
            Call SetSpreadLock("A")
            
        End With
    End If

    If pvSpdNo = "" OR pvSpdNo = "B" Then
    
   	    With frm1.vspdData2
	
            ggoSpread.Source = frm1.vspdData2
            ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	       .ReDraw = false

           .MaxCols   = C_RETRO_OTHER_AMT2 + 1                                                      ' ��:��: Add 1 to Maxcols
	       .Col = .MaxCols                                                              ' ��:��: Hide maxcols
           .ColHidden = True                                                            ' ��:��:
 
           .MaxRows = 0
            ggoSpread.ClearSpreadData

           .DisplayColHeaders = False

            Call GetSpreadColumnPos("B") 'sbk

            Call AppendNumberPlace("6","15","0")

            ggoSpread.SSSetEdit C_PAY_GRD12           , "", 11 , , ,50		
            ggoSpread.SSSetFloat C_EMP_COUNT2         , ""            ,7,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_ORIGINAL_BASE_AMT2 , "",17,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
            ggoSpread.SSSetFloat C_ORIGINAL_OTHER_AMT2, "",17,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	        ggoSpread.SSSetFloat C_RAISE_BASE_AMT2    , ""  ,16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	        ggoSpread.SSSetFloat C_RAISE_OTHER_AMT2   , ""  ,16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	        ggoSpread.SSSetFloat C_RETRO_BASE_AMT2    , ""  ,17,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	        ggoSpread.SSSetFloat C_RETRO_OTHER_AMT2   , ""  ,16,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

	       .ReDraw = true
	
            Call SetSpreadLock("B")
	
        End With
    End If

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
    With frm1
        If pvSpdNo = "A" Then
             ggoSpread.Source = frm1.vspdData
			 ggoSpread.SpreadLockWithOddEvenRowColor()
        ElseIf pvSpdNo = "B" Then
             ggoSpread.Source = frm1.vspdData2
            .vspdData2.ReDraw = False

            ggoSpread.SpreadLock C_PAY_GRD12, -1 , C_PAY_GRD12, -1            
            ggoSpread.SpreadLock C_EMP_COUNT2, -1 , C_EMP_COUNT2, -1
            ggoSpread.SpreadLock C_ORIGINAL_BASE_AMT2, -1 , C_ORIGINAL_BASE_AMT2, -1
            ggoSpread.SpreadLock C_ORIGINAL_OTHER_AMT2, -1 , C_ORIGINAL_OTHER_AMT2, -1
	        ggoSpread.SpreadLock C_RAISE_BASE_AMT2, -1 , C_RAISE_BASE_AMT2, -1
	        ggoSpread.SpreadLock C_RAISE_OTHER_AMT2, -1 , C_RAISE_OTHER_AMT2, -1
	        ggoSpread.SpreadLock C_RETRO_BASE_AMT2, -1 , C_RETRO_BASE_AMT2, -1
	        ggoSpread.SpreadLock C_RETRO_OTHER_AMT2, -1 , C_RETRO_OTHER_AMT2, -1
	        ggoSpread.SSSetProtected   .vspdData2.MaxCols   , -1, -1
             
            .vspdData2.ReDraw = True
        End If

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
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
            
            C_PAY_GRD1 = iCurColumnPos(1)     	
            C_EMP_COUNT = iCurColumnPos(2)    
            C_ORIGINAL_BASE_AMT = iCurColumnPos(3)     
            C_ORIGINAL_OTHER_AMT  = iCurColumnPos(4)   
            C_RAISE_BASE_AMT  = iCurColumnPos(5)     
            C_RAISE_OTHER_AMT  = iCurColumnPos(6)     
            C_RETRO_BASE_AMT = iCurColumnPos(7) 
            C_RETRO_OTHER_AMT = iCurColumnPos(8)

        Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_PAY_GRD12 = iCurColumnPos(1)     	
            C_EMP_COUNT2 = iCurColumnPos(2)    
            C_ORIGINAL_BASE_AMT2 = iCurColumnPos(3)     
            C_ORIGINAL_OTHER_AMT2  = iCurColumnPos(4)   
            C_RAISE_BASE_AMT2  = iCurColumnPos(5)     
            C_RAISE_OTHER_AMT2  = iCurColumnPos(6)     
            C_RETRO_BASE_AMT2 = iCurColumnPos(7) 
            C_RETRO_OTHER_AMT2 = iCurColumnPos(8)            
            
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
    Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock Field
            
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call ggoOper.FormatDate(frm1.txtpay_yymm_dt, Parent.gDateFormat, 2) '<==== �̱ۿ��� ����� �Է��ϰ� ������� ���� �Լ��� ���Ѵ�.
    
    Call SetDefaultVal
    Call SetToolbar("1100000000001111")												'��: Set ToolBar
    
    Call CookiePage (0)                                                             '��: Check Cookie
    
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
	IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '��: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    if  txtPay_grd1_Onchange() then
       Exit Function
    End If
    
    Call SetSpreadLock("B")
    
    Call InitVariables                                                           '��: Initializes local global variables
    Call MakeKeyStream("X")
       
	Call DisableToolBar(Parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
       
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
    
    Call DisableToolBar(Parent.TBC_SAVE)
		IF DBSAVE =  False Then
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

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	With Frm1.VspdData
           .Col  = C_MAJORCD
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
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow
       .vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
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
	Call Parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                    '��:ȭ�� ����, Tab ���� 
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

	ggoSpread.Source = frm1.vspdData2 
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)  
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
    
    If isEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub
	Elseif	UCase(gActiveSpdSheet.id) = "VASPREAD" Then
		ggoSpread.Source = frm1.vspdData2 
		Call ggoSpread.SaveSpreadColumnInf()
	End if

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	dim temp
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("A")      
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ReOrderingSpreadData()

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("B")      
    ggoSpread.Source = frm1.vspdData2
	Call ggoSpread.ReOrderingSpreadData()
	
	temp = GetSpreadText(frm1.vspdData,1,1,"X","X")

	if temp <>"" then
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.InsertRow
		Call InitData()
	end if
End Sub

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit? 
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
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '��: Next key tag
    End With
		
    If lgIntFlgMode = Parent.OPMD_UMODE Then
    Else
    End If
	
	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic
    
    DbQuery = True
    
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
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
 
        Case ggoSpread.InsertFlag                                      '��: Update
                                            strVal = strVal & "C" & Parent.gColSep
                                            strVal = strVal & lRow & Parent.gColSep
                                         
             .vspdData.Col = C_NAME	      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_EMP_NO	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_DEPT_CD	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_ALLOW_CD   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_ALLOW_AMT  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_PAY_CD     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_PROV_TYPE  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
             
             lGrpCnt = lGrpCnt + 1
      
        Case ggoSpread.UpdateFlag                                      '��: Update
                                           strVal = strVal & "U" & Parent.gColSep
                                           strVal = strVal & lRow & Parent.gColSep
             
             .vspdData.Col = C_NAME	      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_EMP_NO	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_DEPT_CD	  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_ALLOW_CD   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_ALLOW_AMT  : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_PAY_CD     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_PROV_TYPE   : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   
             
             lGrpCnt = lGrpCnt + 1
             
        Case ggoSpread.DeleteFlag                                      '��: Delete

                                           strDel = strDel & "D" & Parent.gColSep
                                           strDel = strDel & lRow & Parent.gColSep
             .vspdData.Col = C_NAME	     : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
             .vspdData.Col = C_EMP_NO	 : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep								
             lGrpCnt = lGrpCnt + 1
        End Select
    Next
	
       .txtMode.value        = Parent.UID_M0002
       .txtUpdtUserId.value  = Parent.gUsrID
       .txtInsrtUserId.value = Parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

    End With
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '��: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '��: Processing is NG
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '��:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If

    Call DisableToolBar(Parent.TBC_DELETE)

	IF DBDELETE =  False Then
		Call RestoreToolBar()
		Exit Function
	End If													'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.InsertRow
    Call InitData()
	Call SetToolbar("1100000000011111")									
	frm1.vspdData.focus	
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call InitVariables															'��: Initializes local global variables
	Call DisableToolBar(Parent.TBC_QUERY)
		IF DBQUERY =  False Then
			Call RestoreToolBar()
			Exit Function
		End If
End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function

'========================================================================================================
' Name : OpenCondAreaPopup()        
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Then  
	   Exit Function
	End If   

	IsOpenPop = True
	Select Case iWhere
	     Case "1"
            arrParam(0) = "�����˾�"
	        arrParam(1) = "B_minor"							    ' TABLE ��Ī 
	    	arrParam(2) = frm1.txtpay_grd1.Value  			    ' Code Condition
	    	arrParam(3) = ""'frm1.txtpay_grd1_nm.Value  			' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "����"    						    ' TextBox ��Ī 
	
	    	arrField(0) = "minor_cd"							' Field��(0)
	    	arrField(1) = "minor_nm"    						' Field��(1)
    
	    	arrHeader(0) = "�����ڵ�"			        		' Header��(0)
	    	arrHeader(1) = "���޸�"	       
         
        
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
		
	
	If arrRet(0) = "" Then
		frm1.txtpay_grd1.focus	
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtpay_grd1.value = arrRet(0)
		        .txtpay_grd1_nm.value = arrRet(1)
		        .txtpay_grd1.focus
        End Select
	End With

End Sub

'========================================================================================================
'   Event Name : txtpay_grd1_Onchange()             '<==�ڵ常 �Է��ص� ����Ű,��Ű�� ġ�� �ڵ���� �ҷ��ش� 
'   Event Desc :
'========================================================================================================
Function txtpay_grd1_Onchange()
    Dim IntRetCd

    If Trim(frm1.txtPay_grd1.Value) = "" Then
        frm1.txtpay_grd1_nm.Value=""
    Else
        IntRetCD = CommonQueryRs(" minor_cd,minor_nm "," b_minor "," major_cd=" & FilterVar("H0001", "''", "S") & " And minor_cd =  " & FilterVar(frm1.txtpay_grd1.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtPay_grd1.Value)<>""  Then
            frm1.txtpay_grd1_nm.Value=""
            Call DisplayMsgBox("970000","X","��ȣ�ڵ�","X")             '�� : ��ϵ��� ���� �ڵ��Դϴ�.
            txtpay_grd1_Onchange = true
        Else
            frm1.txtpay_grd1_nm.Value=Trim(Replace(lgF1,Chr(11),""))
        End If
    End If
    
End Function 

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
    End Select    
             
   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.Text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.Text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("0000101111")

    gMouseClickStatus = "SPC" 

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If
End Sub

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc :
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000000000")

    gMouseClickStatus = "SP1C" 

    Set gActiveSpdSheet = frm1.vspdData2
   
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
    
End Sub

'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
        End If
End Sub 

Sub vspdData2_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SP1C" Then
          gMouseClickStatus = "SP1CR"
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

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")

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

    frm1.vspdData.Col = pvCol1
    frm1.vspdData2.ColWidth(pvCol1) = frm1.vspdData.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

    frm1.vspdData2.Col = pvCol1
    frm1.vspdData.ColWidth(pvCol1) = frm1.vspdData2.ColWidth(pvCol1)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col , ByVal Row, ByVal newCol , ByVal newRow ,Cancel )
    frm1.vspdData2.Col = newCol
    frm1.vspdData2.Action = 0
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData2.LeftCol=NewLeft   	
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
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData.LeftCol=NewLeft   	
		Exit Sub
	End If
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
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
'   Event Name : txtpay_yymm_dt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtpay_yymm_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")
        frm1.txtpay_yymm_dt.Action = 7
        frm1.txtpay_yymm_dt.focus
    End If
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDilig_dt_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtpay_yymm_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ұޱ޿����޺���ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%>></TD>
			   </TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
								
								<TD CLASS=TD5 NOWRAP>��ȸ���</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h8006ma1_txtpay_yymm_dt_txtpay_yymm_dt.js'></script></TD>		
							   	<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtpay_grd1" ALT="����" TYPE="Text" SiZE="10" MAXLENGTH="2" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
								                    <INPUT NAME="txtpay_grd1_nm" ALT="����" TYPE="Text" SiZE="20" MAXLENGTH="50" tag="14XXXU"></td>
			               </TR>
	                   </TABLE>
				     </FIELDSET>
				   </TD>
				</TR>
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
                 <TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h8006ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=44 VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%" width=100%>
									<script language =javascript src='./js/h8006ma1_vaSpread1_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD width=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
	
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtCheck"       tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

