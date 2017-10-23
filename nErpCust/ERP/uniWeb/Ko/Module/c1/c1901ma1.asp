
<%@ LANGUAGE="VBSCRIPT" %>
<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 완성품환산율 등록 
'*  3. Program ID           : c1901ma1.asp
'*  4. Program Name         : 완성품환산율 등록 
'*  5. Program Desc         : 완성품환산율 등록 
'*  6. Modified date(First) : 2000/08/23
'*  7. Modified date(Last)  : 2002/06/20
'*  8. Modifier (First)     : Cho Ig Sung
'*  9. Modifier (Last)      : Lee Tae Soo / Park, Joon-Won
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================  -->

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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "c1901mb1.asp"                               'Biz Logic ASP


'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================
Dim C_RoutNo 
Dim C_OprNo 
Dim C_RoutOrder
Dim C_WcCd 
Dim C_WcNm 
Dim C_ProdRate 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

Dim lgStrRoutNoPrevKey
Dim lgStrOprNoPrevKey

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  

 C_RoutNo = 1	
 C_OprNo = 2															
 C_RoutOrder = 3
 C_WcCd = 4
 C_WcNm = 5
 C_ProdRate = 6
 
End Sub


'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE   
    lgBlnFlgChgValue = False  
    lgIntGrpCount = 0
    
    lgStrRoutNoPrevKey = ""    
    lgLngCurRows = 0  
    lgSortKey = 1
End Sub


'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub


'========================================================================================================
Sub InitSpreadSheet()

	call initSpreadPosVariables()

	With frm1.vspdData
	
    .MaxCols = C_ProdRate+1	
    .Col = .MaxCols
    .ColHidden = True
    
    ggoSpread.Source= frm1.vspdData
    ggoSpread.ClearSpreadData
    
    ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    

	.ReDraw = false

    Call GetSpreadColumnPos("A")
    
    ggoSpread.SSSetEdit C_RoutNo, "라우팅번호", 20
    ggoSpread.SSSetEdit C_OprNo, "공정", 10
    ggoSpread.SSSetEdit C_RoutOrder, "공정순서", 10
    ggoSpread.SSSetEdit C_WcCd, "작업장코드", 20
    ggoSpread.SSSetEdit C_WcNm, "작업장명", 27

    Call AppendNumberPlace("6","3","6")
    ggoSpread.SSSetFloat C_ProdRate,"완성품 환산율(%)",30,Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,,"0","100"


'	ggoSpread.SSSetSplit(C_OprNo)
	    
    
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_RoutNo		, -1, C_RoutNo
    ggoSpread.SpreadLock C_OprNo		, -1, C_OprNo
    ggoSpread.SpreadLock C_RoutOrder	, -1, C_RoutOrder
    ggoSpread.SpreadLock C_WcCd			, -1, C_WcCd
    ggoSpread.SpreadLock C_WcNm			, -1, C_WcNm
    ggoSpread.SSSetRequired	C_ProdRate	, -1, -1
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
  
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected	C_RoutNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_OprNo, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_RoutOrder, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_WcCd, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_WcNm, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired	C_ProdRate, pvStartRow, pvEndRow
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
			C_RoutNo			= iCurColumnPos(1)
			C_OprNo				= iCurColumnPos(2)
			C_RoutOrder			= iCurColumnPos(3)    
			C_WcCd				= iCurColumnPos(4)
			C_WcNm				= iCurColumnPos(5)
			C_ProdRate			= iCurColumnPos(6)
			
    End Select    
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
Function OpenPlantCd(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"
	arrParam(1) = "B_PLANT"	
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = ""
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"
    arrField(1) = "PLANT_NM"
    
    arrHeader(0) = "공장코드"	
    arrHeader(1) = "공장명"	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlantCd(arrRet)
	End If	

End Function

Function SetPlantCd(Byval arrRet)
	
	With frm1
		 frm1.txtPlantCd.focus
   		.txtPlantCd.value = arrRet(0)
   		.txtPlantNm.value = arrRet(1)
	
	End With
	
End Function


Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function

	If Trim(frm1.txtPlantCd.Value) = "" Then
		IntRetCD = DisplayMsgBox("125000","x","x","x") '공장을 먼저 입력하세요 
		frm1.txtPlantCd.focus
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "15!MP"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	

	arrField(0) = 1 								' Field명(0) :"ITEM_CD"
	arrField(1) = 2									' Field명(1) :"ITEM_NM"

	arrRet = window.showModalDialog("../../comasp/B1b11pa3.asp", Array(window.parent,arrParam, arrField), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
			
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If	

End Function

Function SetItemCd(Byval arrRet)
	
	With frm1
		 frm1.txtItemCd.focus
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
	End With
	
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029  
    
    Call ggoOper.LockField(Document, "N")              
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
                                                                                             
    Call InitSpreadSheet 
    Call InitVariables   
    
    Call SetDefaultVal
    Call SetToolbar("110010010000111")	
    frm1.txtPlantCd.focus
   	Set gActiveElement = document.activeElement			
    
End Sub


'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

Sub vspdData_Change(ByVal Col , ByVal Row )

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

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


'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")    
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    
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

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_OprNo Or NewCol <= C_OprNo Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)

	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub



Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
 	IF CheckRunningBizProcess = True Then
		Exit Sub
	END IF

    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	
    	If lgStrRoutNoPrevKey <> "" Then  
      	DbQuery
    	End If

    End if
    
End Sub


'========================================================================================================
Function FncQuery() 
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
    
    if frm1.txtItemCd.value = "" then
		frm1.txtItemNm.value = ""
    end if

    if frm1.txtPlantCd.value = "" then
		frm1.txtPlantNm.value = ""
    end if

    If Not chkField(Document, "1") Then	
       Exit Function
    End If
    
    If DbQuery = False then
		Exit function
	END IF
       
    FncQuery = True															
    
End Function


'========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear 
    
    ggoSpread.Source = frm1.vspdData
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
    
    If DbSave = False then
		Exit Function
	END IF	
    
    FncSave = True        
    
    Set gActiveElement = document.ActiveElement                                                     
    
End Function


'========================================================================================================
Function FncCopy() 
    On Error Resume Next
End Function

Function FncCancel() 
 
    if frm1.vspdData.maxrows < 1 then exit function 

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo           
    
    Set gActiveElement = document.ActiveElement                                          
End Function

Function FncInsertRow() 
    On Error Resume Next
End Function

Function FncDeleteRow() 
    On Error Resume Next
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   
End Function

Function FncPrev() 
    On Error Resume Next 
End Function

Function FncNext() 
    On Error Resume Next 
End Function

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

Function FncFind() 
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


'========================================================================================================
Function FncExit()
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

Function DbQuery() 

    DbQuery = False
    
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
    
    Err.Clear 

	Dim strVal
    
    With frm1

    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & .hPlantCd.value				
		strVal = strVal & "&txtItemCd=" & .hItemCd.value				
		strVal = strVal & "&lgStrRoutNoPrevKey=" & lgStrRoutNoPrevKey
		strVal = strVal & "&lgStrOprNoPrevKey=" & lgStrOprNoPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value
		strVal = strVal & "&txtItemCd=" & .txtItemCd.value				
		strVal = strVal & "&lgStrRoutNoPrevKey=" & lgStrRoutNoPrevKey
		strVal = strVal & "&lgStrOprNoPrevKey=" & lgStrOprNoPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If

	Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()

    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")	

	Call SetToolbar("110010010001111")
	Frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   
	
End Function

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
    Dim iColSep 
    Dim iRowSep   
	
    DbSave = False                                                          
    
	IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
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

            Case ggoSpread.UpdateFlag	
					
				strVal = strVal & "U" & iColSep & lRow & iColSep 

                .vspdData.Col = C_RoutNo	
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_OprNo
                strVal = strVal & Trim(.vspdData.Text) & iColSep

                .vspdData.Col = C_ProdRate
                strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
                
        End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
	End With
	
    DbSave = True                                                           
    
End Function

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>완성품환산율등록</font></td>
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
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlantCd frm1.txtPlantCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtItemCd" SIZE=13 MAXLENGTH=25 tag="12XXXU" ALT="품목코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 tag="14"></TD>
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
								<script language =javascript src='./js/c1901ma1_I529006242_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX= "-1"></IFRAME>		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

