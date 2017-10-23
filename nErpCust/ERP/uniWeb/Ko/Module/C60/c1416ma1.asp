<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!-- '**********************************************************************************************
'*  1. Module명          : 원가 
'*  2. Function명        : C_cost_Element_by_Resource
'*  3. Program ID        : c1416ma
'*  4. Program 이름      : 가공비 원가요소 등록 
'*  5. Program 설명      : 가공비 원가요소 등록 수정 삭제 조회 
'*  6. Comproxy 리스트   : c11021, c11028 , ...
'*  7. 최초 작성년월일   : 2000/09/04
'*  8. 최종 수정년월일   : 2002/06/12
'*  9. 최초 작성자       : 송봉훈 
'* 10. 최종 작성자       : Cho Ig sung / Park, Joon-Won
'* 11. 전체 comment      :
'* 12. 공통 Coding Guide : this mark(☜) means that "Do not change"
'*                         this mark(⊙) Means that "may  change"
'*                         this mark(☆) Means that "must change"
'* 13. History           :
'*                         -2000/08/17 : ..........
'********************************************************************************************** -->


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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c1416mb1.asp"                             'Biz Logic ASP

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Dim C_RESOURCE_GRP_CD  
Dim C_RESOURCE_GRP_BTN  
Dim C_RESOURCE_GRP_NM  
Dim C_CE_CD  
Dim C_CE_NM  
Dim C_COMPOSITE_RATE	 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgStrPrevKey1

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	 C_RESOURCE_GRP_CD		= 1
	 C_RESOURCE_GRP_BTN		= 2
	 C_RESOURCE_GRP_NM		= 3
	 C_CE_CD				= 4
	 C_CE_NM				= 5
	 C_COMPOSITE_RATE		= 6

End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgSortKey = 1
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	
    .MaxCols = C_COMPOSITE_RATE + 1
    
    .Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false

	Call GetSpreadColumnPos("A")	
    

   	ggoSpread.SSSetEdit		C_RESOURCE_GRP_CD	,"자원코드"	,21,,,10,2
	ggoSpread.SSSetButton	C_RESOURCE_GRP_BTN
	ggoSpread.SSSetEdit		C_RESOURCE_GRP_NM	,"자원명"	,32
	ggoSpread.SSSetCombo		C_CE_CD	,"원가요소코드"	,20	,0 
	ggoSpread.SSSetCombo		C_CE_NM	,"원가요소명"	,32	,0 
	ggoSpread.SSSetFloat		C_COMPOSITE_RATE	,"구성비율(%)",30,Parent.ggExchRateNo,ggStrIntegeralPart,ggStrDeciPointPart	,Parent.gComNum1000	,Parent.gComNumDec,,,"Z",0,100


	call ggoSpread.MakePairsColumn(C_RESOURCE_GRP_CD,C_RESOURCE_GRP_BTN)

    Call ggoSpread.SSSetColHidden(C_CE_CD,C_CE_CD,True)

	.ReDraw = true
	
'    ggoSpread.SSSetSplit(C_RESOURCE_GRP_NM)	
    Call SetSpreadLock 
    Call initComboBox 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
sub SetSpreadLock()
    With frm1
    
	    .vspdData.ReDraw = False
	    ggoSpread.SSSetProtected C_RESOURCE_GRP_CD	,-1	,C_RESOURCE_GRP_CD
		ggoSpread.SSSetProtected C_RESOURCE_GRP_BTN	,-1	,C_RESOURCE_GRP_BTN
		ggoSpread.SSSetProtected C_RESOURCE_GRP_NM	,-1	,C_RESOURCE_GRP_NM
		ggoSpread.SSSetProtected C_CE_NM				,-1	,C_CE_NM
		ggoSpread.SSSetRequired	C_COMPOSITE_RATE	,-1
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		     
	    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	C_RESOURCE_GRP_CD	,pvStartRow	,pvEndRow
	ggoSpread.SSSetProtected C_RESOURCE_GRP_NM	,pvStartRow	,pvEndRow
	ggoSpread.SSSetRequired	C_CE_NM				,pvStartRow	,pvEndRow
    ggoSpread.SSSetRequired	C_COMPOSITE_RATE	,pvStartRow	,pvEndRow
    
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

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
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
			C_RESOURCE_GRP_CD			= iCurColumnPos(1)
			C_RESOURCE_GRP_BTN			= iCurColumnPos(2)
			C_RESOURCE_GRP_NM			= iCurColumnPos(3)    
			C_CE_CD				        = iCurColumnPos(4)
			C_CE_NM						= iCurColumnPos(5)
			C_COMPOSITE_RATE			= iCurColumnPos(6)
			
    End Select    
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 
'------------------------------------------  OpenCalType()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
'   ggoSpread.SetCombo "10" & vbtab & "20" & vbtab & "30" & vbtab & "50" , C_ItemAcct
'    ggoSpread.SetCombo "제품" & vbtab & "반제품" & vbtab & "원자재"& vbtab & "상품", C_ItemAcctNm
'    ggoSpread.SetCombo "M" & vbtab & "O" & vbtab & "P", C_ProcurType
'    ggoSpread.SetCombo "사내가공품" & vbtab & "외주가공품" & vbtab & "구매품", C_ProcurTypeNm
   
    Call CommonQueryRs(" COST_ELMT_CD,COST_ELMT_NM "," C_COST_ELMT ", " "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                           
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_CE_CD			'COLM_DATA_TYPE
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_CE_NM
     
End Sub


Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
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
		Call SetPlant(arrRet)
	End If	
	
End Function

Function OpenRscGrp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
  Dim intRetCD 
  
	If IsOpenPop = True Then Exit Function

	if Trim(frm1.txtPlantCD.value ) = "" then 
		intRetCD = DisplayMsgBox("125000","x","x","x")
		frm1.txtPlantCD.focus
	    Exit Function
	End If
    
	IsOpenPop = True

	arrParam(0) = "자원팝업"	
	arrParam(1) = "P_RESOURCE"				
	arrParam(2) = strCode
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =  " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " "
	arrParam(5) = "자원"			
	
    arrField(0) = "RESOURCE_CD"	
    arrField(1) = "DESCRIPTION"	
    
    arrHeader(0) = "자원코드"		
    arrHeader(1) = "자원명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtRscGRpCD.focus
		Exit Function
	Else
		Call SetRscGrp(arrRet,iWhere)
	End If	
	
End Function

Function SetRscGrp(Byval arrRet, Byval iWhere)
	
	With frm1
	
    	If iWhere = 0 Then
    		 frm1.txtRscGRpCD.focus
    		.txtRscGRpCd.value = arrRet(0)
    		.txtRscGRpNm.value = arrRet(1)
    	Else
    		.vspdData.Col = C_RESOURCE_GRP_CD
    		.vspdData.Text = arrRet(0)
    		.vspdData.Col = C_RESOURCE_GRP_NM
    		.vspdData.Text = arrRet(1)
    		    		
    		Call vspdData_Change(.vspdData.Col, .vspdData.Row)	
    	End If
	
	End With
	
End Function

Function SetPlant(byval arrRet)
	frm1.txtPlantCd.focus
	frm1.txtPlantCd.Value = arrRet(0)
	frm1.txtPlantNM.value = arrRet(1)
			
End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================

sub Form_Load()

	Call LoadInfTB19029  
	Call ggoOper.LockField(Document, "N")                        
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	                
	Call InitSpreadSheet
	Call InitVariables
	    
	Call SetDefaultVal
'	Call InitComboBox

	Call SetToolbar("110011010010111")
	    
	frm1.txtPlantCd.focus 
    
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
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
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
sub vspdData_Click(ByVal Col, ByVal Row)

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
	'	frm1.vspdData.Col = C_MajorCd
	
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
	Dim strTemp
	Dim intPos1

	With frm1 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_RESOURCE_GRP_BTN Then
        .vspdData.Col = C_RESOURCE_GRP_CD
        .vspdData.Row = Row
        
        Call OpenRscGrp(.vspdData.Text, 1)
        
    End If
     Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
End Sub

sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
		    
			Case  C_CE_NM
				.Col = Col
				intIndex = .Value
				.Col = C_CE_CD
				.Value = intIndex
			
			
		End Select
	End With
End Sub


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
'    If Col <= C_RESOURCE_GRP_NM Or NewCol <= C_RESOURCE_GRP_NM Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub



sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
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
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
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
    	
'   Call InitSpreadSheet
    Call InitComboBox
    		
    Call InitVariables
    
    if frm1.txtPlantCD.value = "" then
		frm1.txtPlantNM.value = ""
    end if
    
    if frm1.txtRscGRpCD.value = "" then
		frm1.txtRscGRpNM.value = ""
    end if
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
	If DbQuery = False Then
		Exit Function
	End If
    
    FncQuery = True															
    
End Function

function FncSave() 
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
     
    ggoSpread.Source = frm1.vspdData
    
    if PKeyCheck = false then
       Exit Function
    end if

	If DbSave = False Then
		Exit Function
	End If
    
    FncSave = True                                                          
    
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
function FncCopy() 
	dim iRow
	dim iNm
	dim iCd
	
	frm1.vspdData.ReDraw = False
	
    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
        
    frm1.vspdData.Col = C_resource_grp_cd
    frm1.vspdData.Text = ""

    frm1.vspdData.Col = C_resource_grp_nm
    frm1.vspdData.Text = ""
       
    with frm1.vspdData
    .row = .ActiveRow - 1
    .col = C_CE_CD
    iCd = .text
            
    .col = C_CE_NM
    iNm = .text
        
    end with
    
	With frm1.vspdData
		.col = C_CE_CD
		.row = .ActiveRow + 1
		.text = icd
				
		.col = C_CE_NM
		.row = .ActiveRow + 1
		.value = iNm
	
	End With

    frm1.vspdData.ReDraw = True
End Function
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

function FncCancel() 

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                 
    
    call InitData
    
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
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

function FncDeleteRow() 
    Dim lDeIRows

    if frm1.vspdData.maxrows < 1 then exit function 
	   

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDeIRows = ggoSpread.DeleteRow
    End With
    
End Function

function FncPrint()
    Call parent.FncPrint()    
End Function

function FncPrev() 
    On Error Resume Next
End Function

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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


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

function DbQuery() 
    
    
    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear

	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then

		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & .hPlantCd.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtRscGrpCd=" & .hRscGrpCd.value
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value				
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&txtRscGrpCd=" & .txtRscGRpCD.value
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    End If
    
	Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
function DbQueryOk()
	
   lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")

	Call SetToolbar("110011110011111")	
	Frm1.vspdData.Focus

	call InitData
    Set gActiveElement = document.ActiveElement   
	
End Function

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
					
		    .Col = C_CE_CD
			intIndex = .Value
			.Col = C_CE_NM
			.Value = intIndex   
			           					
		Next	
	End With
End Sub

function DbSave() 
    Dim pP21011
    Dim IRow        
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
    
      
    Call LayerShowHide(1)
    
  	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	
    
    
	With frm1
		.txtMode.value = Parent.UID_M0002
    
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
    For IRow = 1 To .vspdData.MaxRows
    
        .vspdData.Row = IRow
        .vspdData.Col = 0
        
        Select Case .vspdData.Text

            Case ggoSpread.InsertFlag	
				
				strVal = strVal & "C" & iColSep & IRow & iColSep
                                
                .vspdData.Col = C_RESOURCE_GRP_CD 
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_CE_CD
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_COMPOSITE_RATE
                strVal = strVal & UNIConvNum(.vspdData.Text,0) & iRowSep
                
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
					
				strVal = strVal & "U" & iColSep & IRow & iColSep
				
                .vspdData.Col = C_RESOURCE_GRP_CD
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_CE_CD 
                strVal = strVal & Trim(.vspdData.Text) & iColSep
                
                .vspdData.Col = C_COMPOSITE_RATE
                strVal = strVal & UNIConvNum(.vspdData.Text,0) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag	

				strDel = strDel & "D" & iColSep & IRow & iColSep
                .vspdData.Col = C_RESOURCE_GRP_CD 
                strDel = strDel & Trim(.vspdData.Text) & iColSep
                .vspdData.Col = C_CE_CD
                strDel = strDel & Trim(.vspdData.Text) & iRowSep
                
                lGrpCnt = lGrpCnt + 1
        End Select
                
    Next
	
	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel & strVal
	
	'
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()
	Call InitVariables
	frm1.vspdData.MaxRows = 0
	
    Call FncQuery()
End Function

function PKeyCheck()

dim i , j , k 
dim tempCd 
dim tempSum
Dim intRetCD
  
PKeyCheck = true
  
on error resume next
  
ggoSpread.source = frm1.vspddata
frm1.vspddata.row = frm1.vspddata.activerow
frm1.vspddata.col = C_RESOURCE_GRP_CD
frm1.vspddata.action = 0
ggoSpread.SSSort

with frm1.vspddata 
	for i =1 to .maxrows
		.row = i

		.col = 0
		if (.text <> ggoSpread.DeleteFlag ) then

			.col = C_COMPOSITE_RATE

			if .value < 0 then 
				intRetCD = DisplayMsgBox("231221","x","x","x") 
				PKeyCheck = false
				exit function
			end if

		'	.col = C_RESOURCE_GRP_CD

		'	tempCd = ucase(.value)
		'	tempSum = 0

		'	k = i

		'	for j = i to .maxrows       
		'		.row = j

		'		.col = C_RESOURCE_GRP_CD

		'		if tempCd  = ucase(.value) then

		'			.col = 0
		'			if (.text <> ggoSpread.DeleteFlag ) then
		'				.col = C_COMPOSITE_RATE
		'				tempSum = tempSum + .value
		'				k = j
		'			end if
		'		end if
		'	next
						
		'	i = k

		'	if tempSum <> 100 then 
		'		intRetCD =  DisplayMsgBox("231220","x","x","x")
		'		PKeyCheck = false
		'		exit function
		'	end if
		end if
	next 

end with

end function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>가공비원가요소등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
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
									<TD CLASS="TD6"><INPUT  ClASS="clstxt" NAME="txtPlantCD" MAXLENGTH="4" SIZE=10  ALT ="공장" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()">
													<INPUT NAME="txtPlantNM" MAXLENGTH="30" SIZE=25  ALT ="공장명" tag="14X"></TD>
									
									<TD CLASS="TD5">자원</TD>
									<TD CLASS="TD6"><INPUT NAME="txtRscGRpCD" MAXLENGTH="10" SIZE=10  ALT ="자원코드" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRscGrpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenRscGrp(frm1.txtRscGRpCD.value ,0)">
													<INPUT NAME="txtRscGRpNM" MAXLENGTH="30" SIZE=30  ALT ="자원명" tag="14X"></TD>
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
							<TD HEIGHT="100%">
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
						</TR></TABLE>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hReqStatus" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hRscGrpCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


