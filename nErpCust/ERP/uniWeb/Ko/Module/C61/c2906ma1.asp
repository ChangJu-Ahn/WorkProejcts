<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--
======================================================================================================
*  1. Module Name          : Costing
*  2. Function Name        : ǥ�ؿ����ݿ� 
*  3. Program ID           : C2906MA1.ASP
*  4. Program Name         : ǥ�ؿ����ݿ� 
*  5. Program Desc         : ǥ�ؿ��� ���ǥ�شܰ� �ݿ� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/10/31
*  8. Modified date(Last)  : 2002/06/22
*  9. Modifier (First)     : Lee Tae Soo 
* 10. Modifier (Last)      : Park, Joon-Won
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "C2906MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================

'@Grid_Column
Dim C_ChkFlag 
Dim C_ItemCd 
Dim C_ItemNm 
Dim C_Basicunit 
Dim C_ItemSpec 
Dim C_StdPrc 
Dim C_StockStdPrc 
Dim C_Reference
'@Grid_Row


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop 
         



'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub InitSpreadPosVariables()
 C_ChkFlag = 1
 C_ItemCd = 2										'Spread Sheet�� Column�� ��� 
 C_ItemNm = 3
 C_Basicunit = 4
 C_ItemSpec = 5
 C_StdPrc = 6
 C_StockStdPrc = 7
 C_Reference = 8

End Sub

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call SetToolbar("11000000000000")
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub MakeKeyStream(pRow)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        


'========================================================================================================
Sub InitComboBox()

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    

	With frm1.vspdData
	
       .MaxCols = C_Reference + 1                                                      ' ��:��: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ��:��: Hide maxcols
       .ColHidden = True                                                            ' ��:��:

        ggoSpread.Source = Frm1.vspdData
		ggoSpread.Spreadinit "V20021103",, parent.gAllowDragDropSpread
				
		Call ggoSpread.ClearSpreadData()

	   .ReDraw = false
	
       'Call AppendNumberPlace("6","2","0")

        Call GetSpreadColumnPos("A")
       

 		ggoSpread.SSSetCheck C_ChkFlag, "���౸��", 10, ,"",true    
		ggoSpread.SSSetEdit C_ItemCd, "ǰ���ڵ�",15, 0
		ggoSpread.SSSetEdit C_ItemNm, "ǰ���", 33, 0
		ggoSpread.SSSetEdit C_BasicUnit, "���ش���",10,0
		ggoSpread.SSSetEdit	C_ItemSpec, "ǰ��԰�",20,0
		ggoSpread.SSSetFloat C_StdPrc,"ǥ�شܰ�",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetFloat C_StockStdPrc,"���ǥ�شܰ�",15,Parent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
		ggoSpread.SSSetEdit C_Reference, "", 10, 0, -1, 40

       Call ggoSpread.SSSetColHidden(C_Reference,C_Reference,True)

	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
		ggoSpread.SpreadLock C_ItemCd, -1, C_ItemCd
		ggoSpread.SpreadLock C_ItemNm, -1, C_ItemNm
		ggoSpread.SpreadLock C_BasicUnit , -1, C_BasicUnit
		ggoSpread.SpreadLock C_ItemSpec , -1, C_ItemSpec      
		ggoSpread.SpreadLock C_StdPrc , -1, C_StdPrc
		ggoSpread.SpreadLock C_StockStdPrc , -1, C_StockStdPrc  
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
		ggoSpread.SSSetProtected C_ItemCd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm, IpvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_BasicUnit , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemSpec , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_StcPrc , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_StockStdPrc , pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to pos of err
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
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ChkFlag          = iCurColumnPos(1)
			C_ItemCd          = iCurColumnPos(2)
			C_ItemNm       = iCurColumnPos(3)    
			C_Basicunit        = iCurColumnPos(4)
			C_ItemSpec      = iCurColumnPos(5)
			C_StdPrc = iCurColumnPos(6)
			C_StockStdPrc    = iCurColumnPos(7)
			C_Reference = iCurColumnPos(8)
			
    End Select    
End Sub
 
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================


'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '��: Clear err status
    
	Call LoadInfTB19029                                                              '��: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '��: Lock Field
            
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
	Call SetToolbar("11000000000000")                                              '��: Developer must customize
	
	frm1.txtPlantCd.focus
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
	Call CookiePage (0)                                                              '��: Check Cookie
   	Set gActiveElement = document.activeElement					
End Sub
	

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub


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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '��: Initializes local global variables
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery = False Then                                                      '��: Query db data
       Exit Function
    End If
    
    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '��: Processing is OK
End Function
	

'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '��: Processing is OK
End Function
	

'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '��: Processing is OK
End Function


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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '��: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                              '��: Processing is OK
End Function


'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

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
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '��: Processing is OK
End Function


'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '��: Processing is OK
End Function


'========================================================================================================
Function FncInsertRow()
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '��: If process fails
    Err.Clear                                                                     '��: Clear error status
    
    FncInsertRow = False                                                         '��: Processing is NG

    imRow = AskSpdSheetAddRowCount()
    If imRow = "" Then
        Exit Function
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '��: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '��: Processing is OK
End Function


'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call parent.FncPrint()                                                       '��: Protect system from crashing
    FncPrint = True                                                              '��: Processing is OK
End Function


'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '��: Processing is OK
End Function


'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '��: Processing is OK
End Function


'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '��: Processing is OK
End Function


'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

	Call parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '��: Processing is OK
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		             '��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '��: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================


'========================================================================================================
Function DbQuery()
	Dim strVal

    Err.Clear                                                                    '��: Clear err status
    DbQuery = False                                                              '��: Processing is NG
	
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF                                                       '��: Show Processing Message
    
    With Frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
     '@Query_Hidden     
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'Hidden�� �˻��������� Query
			strVal = strVal & "&txtPlantCd=" & .hPlantCd.value				
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtItemAccntCd=" & .hItemAccntCd.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtItemCd=" & .hItemCd.value
		Else
      '@Query_Text     
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001						'���� �˻��������� Query
			strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value				
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtItemAccntCd=" & .txtItemAccntCd.value
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&txtItemCd=" & .txtItemCd.value
		END IF	    
    End With
   
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic
	
	Call SetToolbar("11000000000111")
    DbQuery = True                                                               '��: Processing is OK
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbSave()
 
End Function


'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '��: Clear err status
    DbDelete = False                                                             '��: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    DbDelete = True                                                              '��: Processing is OK
End Function


'========================================================================================================
Function DbQueryOk()

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
  
                                        
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
 	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
   
End Function

'========================================================================================================
Function DbSaveOk()
    Call InitVariables															     '��: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
                                   '��: Developer must customize
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
   
    DBQuery()
    Set gActiveElement = document.ActiveElement   
    
End Function
	

'========================================================================================================
Function DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Function

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(Col, Row)

	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   

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
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
 	
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
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
      	   DbQuery
    	End If
    End if
End Sub

'======================================================================================================
'	Name : OpenPlant()
'	Description : Plant PopUp
'=======================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "�����ڵ�"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	If arrRet(0) = "" Then
		frm1.txtPlantCD.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If	
	
End Function

'======================================================================================================
'	Name : SetPlant()
'	Description : Plant Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCD.focus
	frm1.txtPlantCd.Value = arrRet(0)
	frm1.txtPlantNM.value = arrRet(1)
End Function

'===========================================================================
' Function Name : OpenItemAccnt()
' Function Desc : OpenItemAccnt(ǰ�����) Reference Popup
'===========================================================================
Function OpenItemAccnt()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function

	
	lgIsOpenPop = True
	
	arrParam(0) = "ǰ������˾�"				' �˾� ��Ī 
	arrParam(1) = "B_MINOR a,b_item_acct_inf b"							' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemAccntCd.value)		' Code Condition
	arrParam(3) = ""	' Name Cindition
	arrParam(4) = "MAJOR_CD=" & FilterVar("P1001", "''", "S") & "  and A.MINOR_CD = B.ITEM_ACCT AND B.ITEM_ACCT_GROUP <> " & FilterVar("6MRO","''","S")
	arrParam(5) = "ǰ�����"						' TextBox ��Ī 
		
    arrField(0) = "MINOR_CD"						' Field��(0)
    arrField(1) = "MINOR_NM"						' Field��(1)
    
    arrHeader(0) = "ǰ������ڵ�"					' Header��(0)
    arrHeader(1) = "ǰ�������"						' Header��(1)

   
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemAccntCd.focus
		Exit Function
	Else
		Call SetItemAccnt(arrRet)
	End If	
	
End Function

'------------------------------------------  SetMinor()  --------------------------------------------------
'	Name : SetItemAccnt()
'	Description : Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemAccnt(Byval arrRet)

If arrRet(0) <> "" Then 
	frm1.txtItemAccntCd.focus									' ���� 
	frm1.txtItemAccntCd.value = arrRet(0)
	frm1.txtItemAccntNm.value = arrRet(1)

End If

End Function

'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item Code PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If lgIsOpenPop = True Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("189220","x","x","x") '������ ���� �Է��ϼ��� 
		frm1.txtPlantCd.focus
		Exit Function
	End If
	
	If Trim(frm1.txtItemAccntCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("990003","x","x","x") 'ǰ�񱸺��� ���� �Է��ϼ��� 
		frm1.txtPlantCd.focus
		Exit Function
	End If

	lgIsOpenPop = True
	
	' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 

	If Trim(frm1.txtItemAccntCd.value) <> "" Then
		arrParam(2) = Mid(CStr(Trim(frm1.txtItemAccntCd.value)),1,1) & Mid(CStr(Trim(frm1.txtItemAccntCd.value)),1,1) 
	ELSE
		arrParam(2) = "15"
	END IF
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(3) = ""							' Default Value
	

	arrField(0) = 1 								' Field��(0) :"ITEM_CD"
	arrField(1) = 2									' Field��(1) :"ITEM_NM"

	arrRet = window.showModalDialog("../../comasp/b1b11pa3.asp", Array(window.parent,arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If	

End Function


'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetItemCd(Byval arrRet)
	With frm1
		 frm1.txtItemCd.focus
		.TxtItemCd.Value = arrRet(0)
		.TxtItemNm.Value = arrRet(1)

		lgBlnFlgChgValue = True
		
	End With
	
End Function

'======================================================================================================
' Function Name : FncBtnTotalExe
' Function Desc : This function is related to BtnTotalExe(�ϰ��ݿ�)
'=======================================================================================================
Function FncBtnTotalExe() 
	
	Dim IntRetCD 
	Dim lRow
	Dim lGrpCnt
	Dim strVal
	
	FncBtnTotalExe = False                                                  		       '��: Processing is NG

	Err.Clear                                                            	 		  '��: Protect system from crashing
	
	On Error Resume Next                                           	       '��: Protect system from crashing

	
	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	If Not chkField(Document, "1")  Then  '��: Check contents area
		Exit Function
	End If
    	
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	

  With frm1
		.txtMode.value = Parent.UID_M0003
		.txtUpdtUserId.value = Parent.gUsrID
		.hChecked.value = "A"
			

		.txtMaxRows.value = lGrpCnt 
		
	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					'��: �����Ͻ� ASP �� ���� 
	End With
	

	'-----------------------
	'Check content area
	'-----------------------
	'ggoSpread.Source = frm1.vspdData
	'ggoSpread.SpreadLock C_ChkFlag, -1, C_ChkFlag
	

		
'	frm1.txtMode.value = Parent.UID_M0002
'	frm1.txtUpdtUserId.value = Parent.gUsrID
'	frm1.hChecked.value = "A"

	
'	Call ExecMyBizASP(frm1, BIZ_PGM_ID)					'��: �����Ͻ� ASP �� ���� 
	
	FncBtnTotalExe = True                                      	                    '��: Processing is OK
End Function


'======================================================================================================
' Function Name : FncBtnSelectedExe
' Function Desc : This function is related to BtnSelectedExe(���ùݿ�)
'=======================================================================================================
Function FncBtnSelectedExe() 
	Dim IntRetCD 
	Dim lRow
	Dim lGrpCnt
	Dim strVal
		
	FncBtnSelectedExe = False                                                  		       '��: Processing is NG

	Err.Clear                                                            	 		  '��: Protect system from crashing
	
	On Error Resume Next                                           	       '��: Protect system from crashing

	if SpreadWorkingChk = false then  Exit Function							'spread check box üũ ���� 

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")
	
	If IntRetCD = vbNo Then
		Exit Function
	End If

	
	'-----------------------
	'Check content area
	'-----------------------
	'ggoSpread.Source = frm1.vspdData
	'ggoSpread.SpreadLock C_ChkFlag, -1, C_ChkFlag
	
	If Not chkField(Document, "1")  Then  '��: Check contents area
		Exit Function
	End If
    	

	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	

 	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.hChecked.value = "S"
			
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 0
    	strVal = ""
    		
    	'-----------------------
		'Data manipulate area
		'-----------------------
		
		
		For lRow = 1 To .vspdData.MaxRows
	    	.vspdData.Row = lRow
			.vspdData.Col = C_ChkFlag
			
			if .vspdData.value = 1  then
					
					.vspdData.Col = C_ItemCd
					strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
					
					strVal = strVal & Trim(CStr(lRow)) & Parent.gRowSep	
					
					lGrpCnt = lGrpCnt + 1
			End if
		Next

		.txtMaxRows.value = lGrpCnt
		.txtSpread.value =  strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)					'��: �����Ͻ� ASP �� ���� 
	End With
	
	
	FncBtnSelectedExe = True                                      	                    '��: Processing is OK

End Function

'========================================================================================================
' Name : FncBtnOk
' Desc : Called by MB Area when update operation is successful
'========================================================================================================
Sub FncBtnOk()
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	Call InitVariables	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadUnLock C_ChkFlag, -1, C_ChkFlag	
	
	frm1.vspdData.MaxRows = 0
    
	Call MainQuery()
                                '��: Developer must customize


	'------ Developer Coding part (End )   -------------------------------------------------------------- 
 	Call ggoOper.LockField(Document, "N")
    Set gActiveElement = document.ActiveElement   
End Sub

'======================================================================================================
' Function Name   : User-defined Method Part
' Description : This part declares user-defined method
'=======================================================================================================

Function SpreadWorkingChk()
   Dim iRows
   Dim ichkCnt
   Dim IntRetCD

   SpreadWorkingChk = False
   ichkCnt = 0

   with frm1.vspdData
	For iRows = 1 to .MaxRows
	    .Col =  C_ChkFlag
	    .Row =  iRows
	    
	    if .Value = 1 then 
		ichkCnt = ichkCnt + 1
	    end if

	Next
	if ichkCnt = 0 then 
	   IntRetCD = DisplayMsgBox("236021","X","X","X")  '���õ� �۾��� �����ϴ�.
	   Exit Function
	end if
   End With
   
   SpreadWorkingChk = True

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<!--'======================================================================================================
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	
'======================================================================================================= -->

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>ǥ�ؿ����ݿ�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD>&nbsp;</TD>					
					<TD>&nbsp;</TD>					
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
									<TD CLASS="TD5">����</TD>
									<TD CLASS="TD6"><INPUT NAME="txtPlantCD" MAXLENGTH="4" SIZE=10  ALT ="����" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
														<INPUT NAME="txtPlantNM" MAXLENGTH="30" SIZE=25  ALT ="�����" tag="14X"></TD>
										
									<TD CLASS="TD5">ǰ�����</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemAccntCd" MAXLENGTH="2" SIZE=10  ALT ="ǰ�����" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenItemAccnt()">
														<INPUT NAME="txtItemAccntNM" MAXLENGTH="30" SIZE=20  ALT ="ǰ���" tag="14X"></TD>
								</TR>
								<TR>	
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT  TYPE=TEXT NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="11XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenItemCd()">
														<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14"></TD>
									<TD CLASS="TDT"></TD>
									<TD CLASS="TD6"></TD>	
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% VALIGN=top COLSPAN=4>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD><BUTTON NAME="btnTotalExe" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnTotalExe()" Flag=1>�ϰ��ݿ�</BUTTON>&nbsp;
					<BUTTON NAME="btnSelectedExe" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnSelectedExe()" Flag=1>���ùݿ�</BUTTON>&nbsp;
				<TD>&nbsp</TD>
				<TD>&nbsp</TD>				
			</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="2x" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAccntCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="2x" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hChecked" tag="2x" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
