
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Account
'*  2. Function Name        : 감가상각률 등록 
'*  3. Program ID           : a7133ma1
'*  4. Program Name         : 감가상각률 등록 
'*  5. Program Desc         : 상각방법별로 감가상각을 등록한다.
'*  6. Modified date(First) : 2003/09/19
'*  7. Modified date(Last)  : 2003/09/19
'*  8. Modifier (First)     : Park, Joon Won
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit				

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "a7133mb1.asp"							'Biz Logic ASP 
 
Dim C_DeprMthCd 
Dim C_DeprMthNm
Dim C_DurYrs
Dim C_DeprRate
Dim C_HDurYrs


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          



'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  
	C_DeprMthCd		= 1 
	C_DeprMthNm		= 2
	C_DurYrs		= 3
	C_DeprRate		= 4
	C_HDurYrs       = 5

End Sub


'========================================================================================================
sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE									'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False									'⊙: Indicates that current mode is Create mode	
    lgIntGrpCount = 0 
    
    lgStrPrevKey = ""											'⊙: initializes Previous Key	
    lgLngCurRows = 0   
	lgSortKey = 1												'⊙: initializes sort direction
	    
End Sub


'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	With frm1.vspdData
	
    .MaxCols = C_HDurYrs+1	
	.Col = .MaxCols						
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20030918",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false

  
    Call GetSpreadColumnPos("A")
	Call AppendNumberPlace("6","3","0")

    ggoSpread.SSSetEdit C_DeprMthCd, "상각방법", 15,,,15,2
    ggoSpread.SSSetEdit C_DeprMthNm, "상각방법명", 20
	ggoSpread.SSSetFloat C_DurYrs,     "내용연(월)수", 14, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
    'ggoSpread.SSSetFloat    C_DurYrs,    "내용년수",      15,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z","0","60"
    ggoSpread.SSSetFloat C_DeprRate,	"상각률*1000", 33,Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z",,999999999
    ggoSpread.SSSetFloat C_HDurYrs,     "내용연(월)수", 14, "6", ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

'	call ggoSpread.MakePairsColumn(C_DeprMtdCd,C_DeprMthNm)
	Call ggoSpread.SSSetColHidden(C_HDurYrs,C_HDurYrs,True)


	.ReDraw = true

    Call SetSpreadLock 
 
    End With
    
End Sub


'======================================================================================================
sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_DeprMthCd	, -1, C_DeprMthCd
	ggoSpread.SpreadLock C_DeprMthNm	, -1, C_DeprMthNm
	ggoSpread.SSSetRequired		C_DurYrs, -1, -1
	ggoSpread.SSSetRequired		C_DeprRate, -1, -1
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub



'======================================================================================================
sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                         'Col               Row           Row2
    ggoSpread.SSSetRequired		C_DurYrs		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_DeprMthCd		,pvStartRow		,pvEndRow    
    ggoSpread.SSSetRequired		C_DeprRate		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_DeprMthNm		,pvStartRow		,pvEndRow    
    .vspdData.ReDraw = True
    
    End With
End Sub


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
			C_DeprMthCd				= iCurColumnPos(1)
			C_DeprMthNm				= iCurColumnPos(2)
			C_DurYrs				= iCurColumnPos(3)    
			C_DeprRate				= iCurColumnPos(4)
			C_HDurYrs				= iCurColumnPos(5)
    End Select    
End Sub



Function OpenDepr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "상각방법팝업"	
	arrParam(1) = "a_asset_depr_method "
	arrParam(2) = Trim(frm1.txtDeprCd.Value)
	arrParam(3) = ""			
	arrParam(4) = ""			
	arrParam(5) = "상각방법"		
	
    arrField(0) = "depr_mthd"	
    arrField(1) = "depr_mthd_nm"		
    
    arrHeader(0) = "상각방법코드"		
    arrHeader(1) = "상각방법명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtDeprCd.focus
		Exit Function
	Else
		Call SetDepr(arrRet)
	End If
		
End Function


Function SetDepr(byval arrRet)
	frm1.txtDeprCd.focus
	frm1.txtDeprCd.Value    = arrRet(0)		
	frm1.txtDeprNm.Value    = arrRet(1)		
End Function


'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================

sub Form_Load()

    Call LoadInfTB19029 

    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
  
    Call InitSpreadSheet       
    Call InitVariables
   
    Call SetDefaultVal
    Call SetToolbar("110011010010111")			

    frm1.txtDeprCd.focus

   	Set gActiveElement = document.activeElement			
     
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

 If lgIntFlgMode <> Parent.OPMD_UMODE Then
	Call SetPopupMenuItemInf("1001111111")
 Else 
	Call SetPopupMenuItemInf("1101111111")
 End If	
 		
    gMouseClickStatus = "SPC"	'Split 상태코드 

    Set gActiveSpdSheet = frm1.vspdData

	 If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row

	'------ Developer Coding part (End   ) --------------------------------------------------------------         
	
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
   

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub


sub vspdData_Change(ByVal Col , ByVal Row )
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub


sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

    
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_IndElmtNm Or NewCol <= C_IndElmtNm Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	IF CheckRunningBizProcess = True Then
		Exit Sub
	End If	

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKey <> "" Then        
      	DbQuery
    	End If

    End if
    
End Sub


'========================================================================================================
function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False 
    
    Err.Clear 

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")		
    	If IntRetCD = vbNo Then
	      	Exit Function
    	End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    	
    Call InitVariables 			
    															
    if frm1.txtDeprCd.value = "" then
		frm1.txtDeprNm.value = ""
    end if
    
    If Not chkField(Document, "1") Then		
       Exit Function
    End If

    IF DbQuery = False Then
		Exit function	
    END If
       
    If Err.number = 0 Then	
       FncQuery = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function


'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False																  '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
function FncSave() 
    Dim IntRetCD 
    
    FncSave = False             
    
    Err.Clear 
    
    ggoSpread.Source = frm1.vspddata
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False  Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If

    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then   
       Exit Function
    End If
    
	If DbSave = False Then
		Exit Function
	End If
    
    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                                         
    
End Function

'========================================================================================================
function FncCopy() 
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow


'    frm1.vspdData.Col = C_DirElmt
'    frm1.vspdData.Text = ""

'    frm1.vspdData.Col = C_DirElmtNm
'    frm1.vspdData.Text = ""
    
	frm1.vspdData.ReDraw = True
	
End Function


'========================================================================================================
function FncCancel() 

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  
End Function


'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim iRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
			Exit Function
		End If	
	End If
	
	If Not chkField(Document, "1") Then		
       Exit Function
    End If

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        
        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
			
			.vspdData.Row = iRow
			.vspdData.Col = C_DeprMthCd
			.vspdData.value = frm1.txtDeprCd.value
			
			.vspdData.Col = C_DeprMthNm
			.vspdData.value = frm1.txtDeprNm.value
    
			.vspdData.ReDraw = True
        Next
        
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function


'========================================================================================================
function FncDeleteRow() 
    Dim lDelRows
    
    if frm1.vspdData.maxrows < 1 then exit function 
	   
    
    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
End Function


'========================================================================================================
function FncPrint()
    Call parent.FncPrint()
End Function


'========================================================================================================
function FncPrev() 
	On Error Resume Next
End Function


'========================================================================================================
function FncNext() 
	On Error Resume Next  
End Function

function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)		
End Function

function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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
'    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub


'========================================================================================================
function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
function DbQuery() 
	Dim strVal

    DbQuery = False
    
	IF LayerShowHide(1) = False Then
		Exit Function                                        
	End If
	
    Err.Clear 

    
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
 
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey				
			strVal = strVal & "&txtDeprCd=" & Trim(.hDeprCd.value)	 	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    	Else
	    	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001	
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey				
			strVal = strVal & "&txtDeprCd=" & Trim(.txtDeprCd.value)	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
	
		Call RunMyBizASP(MyBizASP, strVal)	
        
    End With
    
    DbQuery = True

End Function


'========================================================================================================
function DbQueryOk()					
	
    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")	
	Call SetToolbar("110011110011111")	
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
function DbSave() 
    Dim lRow        
    Dim lGrpCnt
    Dim iColSep
    Dim iRowSep     
	Dim strVal, strDel
	
    DbSave = False        
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If	
	
	With frm1
		.txtMode.value = Parent.UID_M0002
		
		lGrpCnt = 1

		strVal = ""
	    strDel = ""
    
		iColSep = Parent.gColSep
		iRowSep = Parent.gRowSep	

		For lRow = 1 To .vspdData.MaxRows

			.vspdData.Row = lRow
			.vspdData.Col = 0
        
			Select Case .vspdData.Text

	            Case ggoSpread.InsertFlag		

					strVal = strVal & "C" & iColSep & lRow & iColSep

					.vspdData.Col = C_DeprMthCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_DurYrs
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_DeprRate
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
					
					.vspdData.Col = C_HDurYrs
					strVal = strVal & Trim(.vspdData.Text) & iRowSep
					

					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag		

					strVal = strVal & "U" & iColSep & lRow & iColSep	

					.vspdData.Col = C_DeprMthCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_DurYrs	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_DeprRate	
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iColSep
					
					.vspdData.Col = C_HDurYrs
					strVal = strVal & Trim(.vspdData.Text) & iRowSep

					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag		

					strDel = strDel & "D" & iColSep & lRow & iColSep	

					.vspdData.Col = C_DeprMthCd		
					strDel = strDel & Trim(.vspdData.Text) & iColSep
					
					.vspdData.Col = C_DurYrs	
					strDel = strDel & Trim(.vspdData.Text) & iRowSep


					lGrpCnt = lGrpCnt + 1
                
	        End Select
                
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)		
	
	End With
	
    DbSave = True  
    
End Function


'========================================================================================================
Function DbSaveOk()			
   
	Call InitVariables
	frm1.vspdData.MaxRows = 0

	Call MainQuery()
		
End Function


Sub txtDeprCd_onChange()
	Dim IntRetCD
	Dim arrVal

	If frm1.txtDeprCd.value = "" Then Exit Sub

	If CommonQueryRs("DEPR_MTHD_NM", "A_ASSET_DEPR_METHOD ", " DEPR_MTHD=  " & FilterVar(frm1.txtDeprCd.value , "''", "S") & "", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrVal = Split(lgF0, Chr(11)) 
		frm1.txtDeprNm.value= Trim(arrVal(0)) 
	Else
		IntRetCD = DisplayMsgBox("117420","X","X","X")  	
		frm1.txtDeprCd.focus
	End If
End Sub




</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>감가상각률등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>상각방법</TD>
									<TD CLASS="TD656" COLSPAN=3><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtDeprCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="상각방법"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDepr()">
										 <INPUT TYPE=TEXT ID="txtDeprNm" NAME="txtDeprNm" SIZE=25 tag="14X">
									</TD>
								</TR>               
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/a7133ma1_vspdData_vspdData.js'></script>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hDeprCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


