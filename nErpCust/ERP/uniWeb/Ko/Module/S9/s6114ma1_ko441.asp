<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : S6114MA1_KO441															*
'*  4. Program Name         : 수출입경비조회															*
'*  5. Program Desc         : 																			*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2007/12/24																*
'*  8. Modified date(Last)  : 																			*
'*  9. Modifier (First)     : 																			*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :                                											*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<script language="VBScript"	  src="../../inc/incCliRdsQuery.vbs"></script>
<script language="VBScript"	  src="../../inc/incHRQuery.vbs"></script>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID 		= "s6114mb1_ko441.asp" 
Const C_MaxKey          = 20				                         

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                          
Dim IsOpenPop  
Dim lgSaveRow 
Dim lgStrColorFlag

<% 
   BaseDate     = GetSvrDate                                                        
%>  
Dim FirstDateOfDB 
Dim lgStartRow
Dim lgEndRow

Const C_PopBizArea		=	0	
Const C_PopCharge		=	1
Const C_HiddenCol       =   1 

Dim C_FLAG			'FLAG
Dim C_GUBUN			'구분
Dim C_JNL_NM		'비용
Dim C_BP_NM1		'협력업체
Dim C_BP_NM2		'고객사
Dim C_Total			'Total
Dim C_Jan			'1월
Dim C_Feb			'2월
Dim C_Mar			'3월
Dim C_Apr			'4월
Dim C_May			'5월
Dim C_Jun			'6월
Dim C_Jul			'7월
Dim C_Aug			'8월
Dim C_Sep			'9월
Dim C_Oct			'10월
Dim C_Nov			'11월
Dim C_Dec			'12월


FirstDateOfDB	= UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat) 

'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                     
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0    
End Sub

Sub initSpreadPosVariables()

	C_FLAG		= 1		'FLAG
	C_GUBUN		= 2		'구분
	C_JNL_NM	= 3		'비용
	C_BP_NM1	= 4		'협력업체
	C_BP_NM2	= 5		'고객사
	C_Total		= 6		'Total
	C_Jan		= 7		'1월
	C_Feb		= 8		'2월
	C_Mar		= 9		'3월
	C_Apr		= 10	'4월
	C_May		= 11	'5월
	C_Jun		= 12	'6월
	C_Jul		= 13	'7월
	C_Aug		= 14	'8월
	C_Sep		= 15	'9월
	C_Oct		= 16	'10월
	C_Nov		= 17	'11월
	C_Dec		= 18	'12월
			  
End Sub

'========================================================================================================
Sub SetDefaultVal()
	
	Frm1.txtYr.Text = cstr(FirstDateOfDB)
 	Call ggoOper.FormatDate(Frm1.txtYr, parent.gDateFormat, 3)	
	Set gActiveElement = document.ActiveElement
End Sub							
	
									
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub InitSpreadSheet()
	
    'Call SetZAdoSpreadSheet("Z_S6211MA1_KO441","S","A", "V20030523", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
	'Call ggoSpread.SSSetColHidden(C_HiddenCol,C_HiddenCol,True)
    'Call ggoSpread.SSSetSplit2(5)
    'Call SetSpreadLock()    
	
	Call initSpreadPosVariables()    

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

	With frm1.vspdData

        .ReDraw = False
        .MaxCols = C_Dec + 1
        .MaxRows = 0 

        Call GetSpreadColumnPos()

			  
		ggoSpread.SSSetEdit			C_FLAG	,			"FLAG"			, 5 
		ggoSpread.SSSetEdit			C_GUBUN	,			"구분"			, 8,2
		ggoSpread.SSSetEdit			C_JNL_NM	,		"비용"			, 20
		ggoSpread.SSSetEdit			C_BP_NM1	,		"협력업체"		, 20
		ggoSpread.SSSetEdit			C_BP_NM2	,		"고객사"		, 20
		ggoSpread.SSSetFloat		C_Total		,		"Total"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Jan		,		"1월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Feb		,		"2월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Mar		,		"3월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Apr		,		"4월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_May		,		"5월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Jun		,		"6월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Jul		,		"7월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Aug		,		"8월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Sep		,		"9월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Oct		,		"10월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		ggoSpread.SSSetFloat		C_Nov		,		"11월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"		
		ggoSpread.SSSetFloat		C_Dec		,		"12월"			, 15, Parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec, , ,"Z"
		
		Call ggoSpread.SSSetColHidden(C_HiddenCol,C_HiddenCol,True)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)		
		Call ggoSpread.SSSetSplit2(5)       

        Call SetSpreadLock
        
		.ReDraw = true

	End With
	
End Sub

'========================================================================================================

Sub GetSpreadColumnPos()
    Dim iCurColumnPos
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_FLAG		= iCurColumnPos(1) 		'FLAG
	C_GUBUN		= iCurColumnPos(2) 		'구분
	C_JNL_NM	= iCurColumnPos(3) 		'비용
	C_BP_NM1	= iCurColumnPos(4) 		'협력업체
	C_BP_NM2	= iCurColumnPos(5) 		'고객사
	C_Total		= iCurColumnPos(6) 		'Total
	C_Jan		= iCurColumnPos(7) 		'1월
	C_Feb		= iCurColumnPos(8) 		'2월
	C_Mar		= iCurColumnPos(9) 		'3월
	C_Apr		= iCurColumnPos(10) 	'4월
	C_May		= iCurColumnPos(11) 	'5월
	C_Jun		= iCurColumnPos(12) 	'6월
	C_Jul		= iCurColumnPos(13) 	'7월
	C_Aug		= iCurColumnPos(14) 	'8월
	C_Sep		= iCurColumnPos(15) 	'9월
	C_Oct		= iCurColumnPos(16) 	'10월
	C_Nov		= iCurColumnPos(17) 	'11월
	C_Dec		= iCurColumnPos(18) 	'12월
		 
	
End Sub
'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub


'========================================================================================================

Sub Form_Load()
    Call LoadInfTB19029				                                           

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
   
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")	  
	
   
    Frm1.txtYr.Focus
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Function FncQuery() 
    On Error Resume Next                                                      
    Err.Clear                                                                    

	'If ValidDateCheck(frm1.txtConFromDt, frm1.txtConToDt) = False Then Exit Function

    FncQuery = False                                                             
   
    Call ggoOper.ClearField(Document, "2")									     
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    
    Call InitVariables 														    
    
    If Not chkField(Document, "1") Then				
       Exit Function
    End If	    

	
    If DbQuery = False Then 
       Exit Function
    End If   

    If Err.number = 0 Then
       FncQuery = True                                                           
    End If   

    Set gActiveElement = document.ActiveElement      
End Function
	

'========================================================================================================
Function FncPrint()
    On Error Resume Next                                                         
    Err.Clear                                                                   

    FncPrint = False                                                             
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        

    If Err.number = 0 Then
       FncPrint = True                                                           
    End If   

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncExcel() 
    On Error Resume Next                                                        
    Err.Clear                                                                    

    FncExcel = False                                                              

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncExport(parent.C_MULTI)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncExcel = True                                                            
    End If   

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncFind() 
    On Error Resume Next                                                         
    Err.Clear                                                                    

    FncFind = False                                                             

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncFind(parent.C_MULTI, True)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncFind = True                                                            
    End If   

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncExit()
    On Error Resume Next                                                         
    Err.Clear                                                                    

    FncExit = False                                                              
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncExit = True                                                            
    End If   

    Set gActiveElement = document.ActiveElement  
End Function


'========================================================================================================
Function DbQuery() 
'    On Error Resume Next                                                        
'    Err.Clear                                                                    
'	
'    DbQuery = False
'    
'	Call LayerShowHide(1)
'
'    '--------- Developer Coding Part (Start) ----------------------------------------------------------
'    With frm1
'        If lgIntFlgMode  <> parent.OPMD_UMODE Then											
'																	
'			.txtYr.Text = .txtYr.Text
'			.txtHConBizArea.value	= .txtConBizArea.value
'
'			If .rdoConf.checked = True Then				 									
'				.txtHConRdoConfFlag.value = .rdoConf.value
'			ElseIf .rdoNonConf.checked = True Then
'				.txtHConRdoConfFlag.value = .rdoNonConf.value
'			Else
'				.txtHConRdoConfFlag.value = ""
'			End If
'			
'			.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
'			.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
'			.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))
'        End If    
'        
'        .txtHlgPageNo.value	= lgPageNo
'        
'        lgStartRow = .vspdData.MaxRows + 1										
'    End With
'    '--------- Developer Coding Part (End) ------------------------------------------------------------
'    
'	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
'
'    If Err.number = 0 Then
'       DbQuery = True																
'    End If   
'
'    Set gActiveElement = document.ActiveElement   


	Dim strVal
			
	DbQuery = False
	    
	If LayerShowHide(1) = False then
		Exit Function 
	End if
	    
	Err.Clear
	
	With frm1
		
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001  		
		strVal = strVal & "&txtHConBizArea=" & TRIM(.txtConBizArea.value)
		strVal = strVal & "&txtYr=" &.txtYr.text
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows		
		
		Call RunMyBizASP(MyBizASP, strVal) 

	End With
	
	DbQuery = True

	
End Function


'========================================================================================================
Function DbQueryOk()	
	
    On Error Resume Next                                                         
    Err.Clear                                                                     

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE										 
    lgSaveRow        = 1    	
	
	Call SetQuerySpreadColor
    
    'Call SetToolbar("11000000000111")
	

	'--------- Developer Coding Part (Start) ----------------------------------------------------------
	If frm1.vspdData.MaxRows > 0 Then
        frm1.vspdData.Focus		
	Else
		Call SetFocusToDocument("M")	
		frm1.txtConFromDt.txtYr
    End If 
   	'--------- Developer Coding Part (End) ----------------------------------------------------------
	
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================

'========================================================================================================
'	Name : SetQuerySpreadColor()
'	Description : 스프레트시트의 특정 컬럼의 배경색상을 변경 
'========================================================================================================
Sub SetQuerySpreadColor()

	
	Dim iLoopCnt
	Dim Spread
	
	Set Spread = frm1.vspdData
	

	'RGB(204,255,153) '연두  RGB(176,234,244) '하늘색 RGB(224,206,244) '연보라  RGB(251,226,153) '연주황  RGB(255,255,153) '연노랑 

	With Spread
		For  iLoopCnt = 1 to .MaxRows
			.Row = iLoopCnt
			.Col = 1
			If .Text = 12 or .Text = 22 Then			
				.Col = -1			
				.BackColor =  RGB(204,255,153) '연두
				.ForeColor = vbBlue
			End If
		Next
	End With
End Sub


Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
		'사업장 
			Case C_PopBizArea									
				iArrParam(1) = "B_BIZ_AREA"									
				iArrParam(2) = Trim(frm1.txtConBizArea.value)				
				iArrParam(3) = ""											
				iArrParam(4) = ""											
				iArrParam(5) = frm1.txtConBizArea.alt						
				
				iArrField(0) = "ED15" & Parent.gColSep & "BIZ_AREA_CD"		
				iArrField(1) = "ED30" & Parent.gColSep & "BIZ_AREA_NM"		
			
				iArrHeader(0) = frm1.txtConBizArea.alt					
				iArrHeader(1) = frm1.txtConBizAreaNm.alt					

				frm1.txtConBizArea.focus 
			
	End Select
	
	iArrParam(0) = iArrParam(5)										
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function


'========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
			   Case C_PopBizArea					
					.txtConBizArea.value = pvArrRet(0) 
					.txtConBizAreaNm.value = pvArrRet(1)  			
		End Select
	End With

	SetConPopup = True		
	
End Function


'==================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If

	Call OpenOrderBy("A")  
End Sub


'========================================================================================================
Sub OpenOrderBy(ByVal pvPsdNo)
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then												' Means that nothing is happened!!!
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call 3()       
   End If
End Sub


'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
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
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)	    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If	
End Sub


'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery
    	End If
    End If    
End Sub


'========================================================================================================
Sub txtYr_DblClick(Button)
	If Button = 1 Then
       Frm1.txtYr.Action = 7
       Call SetFocusToDocument("Y")	
       Frm1.txtYr.Focus
	End If
End Sub

'========================================================================================================
Sub txtYr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub
'


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>수출입경비조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="30" align=right></td>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%> >
								<TR>
									<TD CLASS="TD5" NOWRAP>년도</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/s6114ma1_ko441_fpDateTime1_txtYr.js'></script>
											</TD>
											
										</TR>
									</TABLE>							        							        
									</TD>
								    <TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP>	
									
									<INPUT TYPE=TEXT NAME="txtConBizArea" SIZE=10 MAXLENGTH=10 tag="12NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizArea) ">
									<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>									
								</TR>						
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							
							<TR>
								<TD HEIGHT="50%" COLSPAN=4>
									<script language =javascript src='./js/s6114ma1_ko441_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>    
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHConBizArea"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSalesGrp"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConRdoConfFlag"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgPageNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1"> 
<INPUT TYPE=HIDDEN name="txtMode"				tag="24" tabindex="-1">
<INPUT TYPE=HIDDEN name="txtMaxRows"			tag="24" tabindex="-1">	

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
