
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 간접비배부율등록 
'*  3. Program ID           : c1210ma1
'*  4. Program Name         : 간접비배부율등록 
'*  5. Program Desc         : 공장별 표준계산시 간접비에 대한 배부율을 등록한다 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2002/06/05
'*  8. Modifier (First)     : 강창구 
'*  9. Modifier (Last)      : Cho Ig Sung / Park, Joon-Won
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit				

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "c1902mb1.asp"							'Biz Logic ASP 
 
Dim C_PlantCd
Dim C_PlantCdPop
Dim C_PlantNm 
Dim C_ItemCd
Dim C_ItemCdPop
Dim C_ItemNm 
Dim	C_ItemAcctNm
Dim C_ProcurTypeNm
Dim C_AdjRate


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

Dim lgPlantPrevKey
Dim lgItemPrevKey


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  
	C_PlantCd		= 1
	C_PlantCdPop	= 2
	C_PlantNm		= 3
	C_ItemCd		= 4
	C_ItemCdPop		= 5
	C_ItemNm		= 6
	C_ItemAcctNm	= 7
	C_ProcurTypeNm	= 8	
	C_AdjRate		= 9
End Sub


'========================================================================================================
sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE									'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False									'⊙: Indicates that current mode is Create mode	
    lgIntGrpCount = 0 
    
    lgPlantPrevKey = ""											'⊙: initializes Previous Key	
    lgItemPrevKey = ""											'⊙: initializes Previous Key	
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
	
    .MaxCols = C_AdjRate+1	
	.Col = .MaxCols						
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false
	
   
    Call GetSpreadColumnPos("A")

	Call AppendNumberPlace("6","6","0")
	
    ggoSpread.SSSetEdit		C_PlantCd,	"공장", 10,,,4,2
	ggoSpread.SSSetButton	C_PlantCdPop
    ggoSpread.SSSetEdit		C_PlantNm,	 "공장명", 25,,,40
    ggoSpread.SSSetEdit		C_ItemCd, "품목", 20,,,18,2
	ggoSpread.SSSetButton	C_ItemCdPop
    ggoSpread.SSSetEdit		C_ItemNm, "품목명", 20,,,40
    ggoSpread.SSSetEdit		C_ItemAcctNm, "품목계정",15,,,50
    ggoSpread.SSSetEdit		C_ProcurTypeNm, "조달구분", 15,,,50
    'ggoSpread.SSSetFloat	C_AdjRate,"가중치(%)", 20,Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z",0,999999999
	ggoSpread.SSSetFloat	C_AdjRate,"가중치(%)", 20,6,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
	
	call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantCdPop)
	call ggoSpread.MakePairsColumn(C_ItemCd,C_ItemCdPop)	


	.ReDraw = true

'    ggoSpread.SSSetSplit(C_IndElmtNm)	
    Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_PlantCd		, -1, C_PlantCd
	ggoSpread.SpreadLock C_PlantCdPop	, -1, C_PlantCdPop
	ggoSpread.SpreadLock C_PlantNm		, -1, C_PlantNm
	ggoSpread.SpreadLock C_ItemCd		, -1, C_ItemCd
	ggoSpread.SpreadLock C_ItemCdPop	, -1, C_ItemCdPop
	ggoSpread.SpreadLock C_ItemNm		, -1, C_ItemNm
	ggoSpread.SpreadLock C_ItemAcctNm	, -1, C_ItemAcctNm
	ggoSpread.SpreadLock C_ProcurTypeNm	, -1, C_ProcurTypeNm
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                         'Col               Row           Row2
    ggoSpread.SSSetRequired		C_PlantCd		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_PlantNm		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_ItemAcctNm	,pvStartRow		,pvEndRow    
    ggoSpread.SSSetProtected	C_ProcurTypeNm	,pvStartRow		,pvEndRow        
    ggoSpread.SSSetRequired		C_ItemCd		,pvStartRow		,pvEndRow
    ggoSpread.SSSetProtected	C_ItemNm		,pvStartRow		,pvEndRow    
    ggoSpread.SSSetRequired  	C_AdjRate		,pvStartRow		,pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub


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
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PlantCd				= iCurColumnPos(1)
			C_PlantCdPop			= iCurColumnPos(2)
			C_PlantNm				= iCurColumnPos(3)
			C_ItemCd				= iCurColumnPos(4)
			C_ItemCdPop				= iCurColumnPos(5)
			C_ItemNm				= iCurColumnPos(6)
			C_ItemAcctNm			= iCurColumnPos(7)
			C_ProcurTypeNm			= iCurColumnPos(8)	    
			C_AdjRate				= iCurColumnPos(9)
    End Select    
End Sub


Function OpenCodeCond(ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim tempStr
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	IF iWhere = 1 Then
		arrParam(0) = "공장 팝업"	
		arrParam(1) = "B_PLANT"
		arrParam(2) = Trim(frm1.txtPlantCd.Value)
		arrParam(3) = ""			
		arrParam(4) = ""			
		arrParam(5) = "공장"		
	
		arrField(0) = "PLANT_CD"	
		arrField(1) = "PLANT_NM"		
    
		arrHeader(0) = "공장"		
		arrHeader(1) = "공장명"		
    
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	ElseIf iWhere = 2 Then
		arrParam(0) = "품목계정 팝업"	
		arrParam(1) = "B_MINOR a,b_item_acct_inf b"
		arrParam(2) = Trim(frm1.txtItemAcct.Value)
		arrParam(3) = ""			
		arrParam(4) = "a.major_cd = " & FilterVar("P1001", "''", "S") & "  and A.MINOR_CD = B.ITEM_ACCT AND B.ITEM_ACCT_GROUP IN (" & FilterVar("1FINAL","''","S") & "," & FilterVar("2SEMI","''","S") & ")"
		arrParam(5) = "품목계정"		
	
		arrField(0) = "MINOR_CD"	
		arrField(1) = "MINOR_NM"		
    
		arrHeader(0) = "품목계정"		
		arrHeader(1) = "품목계정명"		
    
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	ElseIF iWhere = 3 Then
		arrParam(0) = "조달구분 팝업"	
		arrParam(1) = "B_MINOR"
		arrParam(2) = Trim(frm1.txtProcurType.Value)
		arrParam(3) = ""			
		arrParam(4) = " major_cd = " & FilterVar("P1003", "''", "S") & "  and minor_cd in (" & FilterVar("M", "''", "S") & " ," & FilterVar("O", "''", "S") & " ) "		'사내가공,외주가공		
		arrParam(5) = "품목계정"		
	
		arrField(0) = "MINOR_CD"	
		arrField(1) = "MINOR_NM"		
    
		arrHeader(0) = "품목계정"		
		arrHeader(1) = "품목계정명"		
    
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		IF Trim(frm1.txtPlantCd.value) <> "" Then
			arrParam(0) = frm1.txtPlantCd.value
			arrParam(1) = frm1.txtItemCD.Value
		
			IF Trim(frm1.txtItemAcct.value) <> "" Then
				tempStr = Trim(frm1.txtItemAcct.Value) & Trim(frm1.txtItemAcct.value) & "!"
			ELSE
				tempStr = "12!"
			End IF

			IF Trim(frm1.txtProcurType.value) <> "" Then
				tempStr = tempstr & Trim(frm1.txtProcurType.Value) & Trim(frm1.txtProcurType.value) 

			End IF

		
			arrParam(2) = tempStr						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
			arrParam(3) = ""							' Default Value

			arrField(0) = 1 							' Field명(0) : "ITEM_CD"
			arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
			iCalledAspName = AskPRAspName("B1b11pa3")
	
			If Trim(iCalledAspName) = "" Then
				IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1b11pa3", "X")
				IsOpenPop = False
				Exit Function
			End If
	
			arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
				"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")			
		ELSE
			arrParam(0) = "품목 팝업"	
			arrParam(1) = "B_ITEM"
			arrParam(2) = Trim(frm1.txtItemCd.Value)
			arrParam(3) = ""			
			
			IF Trim(frm1.txtItemAcct.value) <> "" Then
				arrParam(4) = " item_acct = " & FilterVar(frm1.txtItemAcct.value, "''", "S")
			ELSE
				arrParam(4) = ""
			END IF
	
			
			arrParam(5) = "품목"		
	
			arrField(0) = "ITEM_CD"	
			arrField(1) = "ITEM_NM"		
    
			arrHeader(0) = "품목"		
			arrHeader(1) = "품목명"		
    
			arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			
		END IF
	End IF
	
	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
		   Case 1
			frm1.txtPlantCD.focus
		   Case 2
		    frm1.txtItemAcct.focus
		   Case 3 
		    frm1.txtProcurType.focus
		   Case 4
		    frm1.txtItemCD.focus
	    End Select	     
		Exit Function
	Else
		Call SetCodeCond(arrRet,iWhere)
	End If
End Function

Function SetCodeCond(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				 frm1.txtPlantCD.focus 
				.txtPlantCD.value = arrRet(0)
				.txtPlantNM.value = arrRet(1)
			Case 2
				 frm1.txtItemAcct.focus	
				.txtItemAcct.value = arrRet(0)
				.txtItemAcctNm.value = arrRet(1)
			Case 3
			     frm1.txtProcurType.focus
				.txtProcurType.value = arrRet(0)
				.txtProcurTypeNm.value = arrRet(1)
			Case Else
			     frm1.txtItemCD.focus
				.txtItemCD.value = arrRet(0)
				.txtItemNM.value = arrRet(1)
		End Select

	End With
	
End Function


Function OpenCode(Byval strCode1,Byval StrCode2, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 1 Then
		arrParam(0) = "공장 팝업"
		arrParam(1) = "B_PLANT"	
		arrParam(2) = strCode1
		arrParam(3) = ""		
		arrParam(4) = ""			
		arrParam(5) = "공장"   

		arrField(0) = "PLANT_CD"			
		arrField(1) = "PLANT_NM"		
    
		arrHeader(0) = "공장"	   		
		arrHeader(1) = "공장명"					

		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	Else
		arrParam(0) = strCode1
		arrParam(1) = strCode2
		arrParam(2) = "1025!MO"							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
		arrParam(3) = ""							' Default Value
	
	
		arrField(0) = 1 							' Field명(0) : "ITEM_CD"
		arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
		iCalledAspName = AskPRAspName("B1B11PA2")
	
		If Trim(iCalledAspName) = "" Then
			IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA2", "X")
			IsOpenPop = False
			Exit Function
		End If
	
		arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")			
	End If
    
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
	End If	

End Function


Function SetCode(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.vspdData.Col = C_PlantCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_PlantNm
				.vspdData.Text = arrRet(1)
				
				Call vspddata_Change(.vspddata.col, .vspddata.row)
				.vspdData.Col = C_PlantCd
				.vspdData.Action = 0
				
			Case 2
				.vspdData.Col = C_ItemCd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_ItemNm
				.vspdData.Text = arrRet(1)
			
				Call vspddata_Change(.vspddata.col, .vspddata.row)
				.vspdData.Col = C_ItemCd
				.vspdData.Action = 0
		End Select


	End With
	
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
    frm1.txtPlantCd.focus
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
	Dim strPlantCd
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_PlantCdPop
				.vspdData.Col = C_PlantCd
				.vspdData.Row = Row
				
				Call OpenCode(.vspdData.Text,"", 1)
			Case C_ItemCdPop        
				
				.vspdData.Col = C_PlantCd
				.vspdData.Row = Row
				strPlantCd = .vspdData.text
				
				.vspdData.Col = C_ItemCd
				Call OpenCode(strPlantCd,.vspdData.Text, 2)
		End Select
		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
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
    	If lgPlantPrevKey <> "" Then        
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
    															
    if frm1.txtPlantCd.value = "" then
		frm1.txtPlantNm.value = ""
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

    frm1.vspdData.Col = C_PlantCd
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_PlantNm
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_ItemCd
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_ItemNm
    frm1.vspdData.Text = ""
    
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

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
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
			strVal = strVal & "&lgPlantPrevKey=" & lgPlantPrevKey
			strVal = strVal & "&lgItemPrevKey=" & lgItemPrevKey
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
			strVal = strVal & "&txtItemAcct=" & Trim(.hItemAcct.value)	 	
			strVal = strVal & "&txtProcurType=" & Trim(.hProcurType.value)	 	
			strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)	 		 	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    	Else
    		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgPlantPrevKey=" & lgPlantPrevKey
			strVal = strVal & "&lgItemPrevKey=" & lgItemPrevKey
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtItemAcct=" & Trim(.txtItemAcct.value)	 	
			strVal = strVal & "&txtProcurType=" & Trim(.txtProcurType.value)	 	
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)	 		 	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If
	
		Call RunMyBizASP(MyBizASP, strVal)	
        
    End With
    
    DbQuery = True

End Function


'========================================================================================================
function DbQueryOk()					
	
    lgIntFlgMode = Parent.OPMD_UMODE
    
    With frm1
		.hPlantCd.value = .txtPlantCd.value
		.hItemAcct.Value = .txtItemAcct.value
		.hProcurType.value = .txtProcurType.value
		.hItemCd.value = .txtItemCD.value
    End With    

    Call ggoOper.LockField(Document, "Q")	
	Call SetToolbar("110011110011111")	
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
    Dim iColSep 
    Dim iRowSep   
	
    DbSave = False        
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep		

	With frm1
		.txtMode.value = Parent.UID_M0002
		
		lGrpCnt = 1

		strVal = ""
	    strDel = ""
    
		For lRow = 1 To .vspdData.MaxRows

			.vspdData.Row = lRow
			.vspdData.Col = 0
        
			Select Case .vspdData.Text

	            Case ggoSpread.InsertFlag		

					strVal = strVal & "C" & iColSep & lRow & iColSep

					.vspdData.Col = C_PlantCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ItemCd
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_AdjRate
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep

					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag		

					strVal = strVal & "U" & iColSep & lRow & iColSep	

					.vspdData.Col = C_PlantCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ItemCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_AdjRate	
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep

					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag		

					strDel = strDel & "D" & iColSep & lRow & iColSep	

					.vspdData.Col = C_PlantCd		
					strDel = strDel & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ItemCd	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별가중치등록</font></td>
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
									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD6"><INPUT CLASS="clstxt" NAME="txtPlantCD" MAXLENGTH="4" SIZE=10  ALT ="공장" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCodeCond(1)">
														<INPUT NAME="txtPlantNM" MAXLENGTH="30" SIZE=25  ALT ="공장명" tag="14X"></TD>
										
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemAcct" MAXLENGTH="2" SIZE=10  ALT ="품목계정" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCodeCond(2)">
														<INPUT NAME="txtItemAcctNm" MAXLENGTH="30" SIZE=20  ALT ="품목계정명" tag="14X"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">조달구분</TD>
									<TD CLASS="TD6"><INPUT  NAME="txtProcurType" MAXLENGTH="2" SIZE=10  ALT ="조달구분코드" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcurType" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCodeCond(3)">
														<INPUT NAME="txtProcurTypeNm" MAXLENGTH="30" SIZE=25  ALT ="조달구분" tag="14X"></TD>
										
									<TD CLASS="TD5">품목</TD>
									<TD CLASS="TD6"><INPUT  NAME="txtItemCD" MAXLENGTH="18" SIZE=10  ALT ="품목" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenCodeCond(4)">
														<INPUT NAME="txtItemNM" MAXLENGTH="30" SIZE=25  ALT ="품목명" tag="14X"></TD>
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
								<script language =javascript src='./js/c1902ma1_vspdData_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemAcct" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hProcurType" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


