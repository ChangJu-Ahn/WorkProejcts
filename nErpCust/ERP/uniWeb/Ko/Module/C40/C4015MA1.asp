<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 원가요소별 표준Rate 등록 
'*  3. Program ID           : C4015ma1
'*  4. Program Name         : 원가요소별 표준Rate 등록 
'*  5. Program Desc         : 원가요소별 표준Rate 등록 
'*  6. Modified date(First) : 2004/03/22
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Tae Soo 
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

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit				

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "C4015MB1.asp"							'Biz Logic ASP 

Dim C_ItemAcctCd
Dim C_ItemAcctPb
Dim C_ItemAcctNm
Dim C_PlantCd 
Dim C_PlantPb 
Dim C_PlantNm 
Dim C_ItemGroupCd 
Dim C_ItemGroupPb 
Dim C_ItemGroupNm 
Dim C_ItemCd 
Dim C_ItemPb 
Dim C_ItemNm
Dim	C_CostElmtCd
Dim	C_CostElmtPb
Dim C_CostElmtNm 
Dim C_BasRate


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

Dim lgItemAcctPrevKey
Dim lgPlantPrevKey
Dim lgItemGroupPrevKey
Dim lgItemPrevKey
Dim lgCostElmtPrevKey

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================

'========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemAcctCd	= 1
	C_ItemAcctPb	= 2
	C_ItemAcctNm	= 3
	C_PlantCd		= 4 
	C_PlantPb		= 5 
	C_PlantNm		= 6 
	C_ItemGroupCd	= 7 
	C_ItemGroupPb	= 8 
	C_ItemGroupNm	= 9
	C_ItemCd		= 10
	C_ItemPb		= 11
	C_ItemNm		= 12
	C_CostElmtCd	= 13
	C_CostElmtPb	= 14
	C_CostElmtNm	= 15
	C_BasRate		= 16
End Sub


'========================================================================================================
sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE									'⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False									'⊙: Indicates that current mode is Create mode	
    lgIntGrpCount = 0 
    
    lgItemAcctPrevKey = ""	
    lgPlantPrevKey = ""											'⊙: initializes Previous Key	
    lgItemGroupPrevKey = ""	
    lgItemPrevKey = ""	
    lgCostElmtPrevKey = ""
    
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

	
    .MaxCols = C_BasRate+1	
	.Col = .MaxCols						
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false
	
    Call AppendNumberPlace("6","3","0")   
    Call GetSpreadColumnPos("A")
    

    

    ggoSpread.SSSetEdit C_ItemAcctCd,	"품목계정", 8,,,2,2
	ggoSpread.SSSetButton C_ItemAcctPb
    ggoSpread.SSSetEdit C_ItemAcctNm,	"품목계정명", 13
    ggoSpread.SSSetEdit C_PlantCd,		"공장", 10,,,4,2
	ggoSpread.SSSetButton C_PlantPb
    ggoSpread.SSSetEdit C_PlantNm,		"공장명", 15
    ggoSpread.SSSetEdit C_ItemGroupCd,	"품목그룹", 15,,,10,2
	ggoSpread.SSSetButton C_ItemGroupPb
    ggoSpread.SSSetEdit C_ItemGroupNm,	"품목그룹명", 15
    ggoSpread.SSSetEdit C_ItemCd,		"품목", 15,,,18,2
	ggoSpread.SSSetButton C_ItemPb
    ggoSpread.SSSetEdit C_ItemNm,		"품목명", 25
    ggoSpread.SSSetEdit C_CostElmtCd,	"원가요소", 8,,,20,2
	ggoSpread.SSSetButton C_CostElmtPb
    ggoSpread.SSSetEdit C_CostElmtNm,	"원가요소명", 10


    ggoSpread.SSSetFloat  C_BasRate		,"표준Rate",15,6,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"      


	call ggoSpread.MakePairsColumn(C_ItemAcctCd,C_ItemAcctPb)
	call ggoSpread.MakePairsColumn(C_PlantCd,C_PlantPb)
	call ggoSpread.MakePairsColumn(C_ItemGroupCd,C_ItemGroupPb)
	call ggoSpread.MakePairsColumn(C_ItemCd,C_ItemPb)
	call ggoSpread.MakePairsColumn(C_CostElmtCd,C_CostElmtPb)


	.ReDraw = true

'    ggoSpread.SSSetSplit(C_IndElmtNm)	
    Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
	ggoSpread.SpreadLock C_ItemAcctCd		, -1, C_ItemAcctCd
	ggoSpread.SpreadLock C_ItemAcctPb	, -1, C_ItemAcctPb
	ggoSpread.SpreadLock C_ItemAcctNm	, -1, C_ItemAcctNm
	ggoSpread.SpreadLock C_PlantCd		, -1, C_PlantCd
	ggoSpread.SpreadLock C_PlantPb	, -1, C_PlantPb
	ggoSpread.SpreadLock C_PlantNm	, -1, C_PlantNm
	ggoSpread.SpreadLock C_ItemGroupCd		, -1, C_ItemGroupCd
	ggoSpread.SpreadLock C_ItemGroupPb	, -1, C_ItemGroupPb
	ggoSpread.SpreadLock C_ItemGroupNm	, -1, C_ItemGroupNm
	ggoSpread.SpreadLock C_ItemCd		, -1, C_ItemCd
	ggoSpread.SpreadLock C_ItemPb	, -1, C_ItemPb
	ggoSpread.SpreadLock C_ItemNm	, -1, C_ItemNm
	ggoSpread.SpreadLock C_CostElmtCd		, -1, C_CostElmtCd
	ggoSpread.SpreadLock C_CostElmtPb	, -1, C_CostElmtPb
	ggoSpread.SpreadLock C_CostElmtNm	, -1, C_CostElmtNm
	
	
	ggoSpread.SSSetRequired C_BasRate		, -1, C_BasRate
	
	
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
                                         'Col               Row           Row2
    ggoSpread.SSSetRequired		C_ItemAcctCd	,pvStartRow		,pvEndRow
    ggoSpread.SSSetRequired		C_CostElmtCd	,pvStartRow		,pvEndRow
    ggoSpread.SSSetRequired		C_BasRate		,pvStartRow		,pvEndRow
   
    ggoSpread.SSSetProtected	C_ItemAcctNm		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_PlantCd		,pvStartRow		,pvEndRow
	ggoSpread.SSSetProtected	C_PlantPb		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_PlantNm		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ItemGroupCd		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ItemGroupPb		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ItemGroupNm		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ItemCd		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ItemPb		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_ItemNm		,pvStartRow		,pvEndRow    
	ggoSpread.SSSetProtected	C_CostElmtNm	,pvStartRow		,pvEndRow    
	
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
			        
			C_ItemAcctCd	= iCurColumnPos(1)
			C_ItemAcctPb	= iCurColumnPos(2)
			C_ItemAcctNm	= iCurColumnPos(3)
			C_PlantCd		= iCurColumnPos(4) 
			C_PlantPb		= iCurColumnPos(5) 
			C_PlantNm		= iCurColumnPos(6) 
			C_ItemGroupCd	= iCurColumnPos(7) 
			C_ItemGroupPb	= iCurColumnPos(8) 
			C_ItemGroupNm	= iCurColumnPos(9)
			C_ItemCd		= iCurColumnPos(10)
			C_ItemPb		= iCurColumnPos(11)
			C_ItemNm		= iCurColumnPos(12)
			C_CostElmtCd	= iCurColumnPos(13)
			C_CostElmtPb	= iCurColumnPos(14)
			C_CostElmtNm	= iCurColumnPos(15)
			C_BasRate		= iCurColumnPos(16)
			        
    End Select    
End Sub



Function OpenPopup(ByVal iWhere,Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere
		case 1,2
			arrParam(0) = "품목계정팝업"	
			arrParam(1) = "B_MINOR a,b_item_acct_inf b"
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""			
			arrParam(4) = "a.major_cd = 'P1001' and a.minor_cd = b.item_acct and b.item_acct_group <> '6MRO' "
			arrParam(5) = "품목계정"		
	
			arrField(0) = "minor_cd"		
			arrField(1) = "minor_nm"		
    
			arrHeader(0) = "품목계정"		
			arrHeader(1) = "품목계정명"		
		case 3,4
			arrParam(0) = "공장팝업"	
			arrParam(1) = "B_PLANT "
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""			
			arrParam(4) = ""			
			arrParam(5) = "공장"		
	
			arrField(0) = "PLANT_CD"		
			arrField(1) = "PLANT_NM"		
    
			arrHeader(0) = "공장코드"		
			arrHeader(1) = "공장명"		
		case 5,6
			arrParam(0) = "품목그룹팝업"	
			arrParam(1) = "B_ITEM_GROUP a,B_ITEM b,B_ITEM_BY_PLANT c"
			arrParam(2) = ""
			arrParam(3) = ""			
			arrParam(4) = "a.item_group_cd = b.item_group_cd and b.item_cd = c.item_cd "			
			
			IF iWhere = 5 Then	'Spread
				frm1.vspddata.Col = C_ItemAcctCd
				IF Trim(frm1.vspddata.text) <> "" Then
					arrParam(4) =  arrParam(4) & " and c.item_acct = " & FilterVar(frm1.vspddata.text, "''", "S")
				END IF
				
				frm1.vspddata.Col = C_PlantCd
				IF Trim(frm1.vspddata.text) <> "" Then
					arrParam(4) =  arrParam(4) & " and c.plant_cd= " & FilterVar(frm1.vspddata.text, "''", "S")
				END IF
			

				frm1.vspddata.Col = C_ItemGroupCd
				IF Trim(frm1.vspddata.text) <> "" Then
					arrParam(4) =  arrParam(4) & " and a.item_group_cd LIKE " & FilterVar(Trim(frm1.vspddata.text) & "%", "''", "S") 
				END IF
			ELSE
				IF Trim(frm1.txtItemAcctCd.Value) <> "" Then
					arrParam(4) =  arrParam(4) & " and c.item_acct = " & FilterVar(frm1.txtItemAcctCd.Value, "''", "S")
				END IF
				
				IF Trim(frm1.txtPlantCd.Value) <> "" Then
					arrParam(4) =  arrParam(4) & " and c.plant_cd= " & FilterVar(frm1.txtPlantCd.Value, "''", "S")
				END IF

				IF Trim(frm1.txtItemGroupCd.Value) <> "" Then
					arrParam(4) =  arrParam(4) & " and a.item_group_cd LIKE " & FilterVar(Trim(frm1.txtItemGroupCd.Value) & "%", "''", "S")
				END IF				
			END IF
			
			arrParam(5) = "품목그룹"		
	
			arrField(0) = "a.item_group_cd"		
			arrField(1) = "a.item_group_nm"		
    
			arrHeader(0) = "품목그룹코드"		
			arrHeader(1) = "품목그룹명"		
		case 7,8
		

			arrParam(0) = "품목팝업"	
			arrParam(1) = "B_ITEM a,B_ITEM_GROUP b,B_ITEM_BY_PLANT c"
			arrParam(2) = ""
			arrParam(3) = ""			
			arrParam(4) = "a.item_group_cd = b.item_group_cd and a.item_cd = c.item_cd"			
			arrParam(5) = "품목"	
			
			IF iWhere = 7 Then	'Spread
				frm1.vspddata.Col = C_ItemAcctCd
				IF Trim(frm1.vspddata.text) <> "" Then
					arrParam(4) =  arrParam(4) & " and c.item_acct = " & FilterVar(frm1.vspddata.text, "''", "S")
				END IF
				
				frm1.vspddata.Col = C_PlantCd
				IF Trim(frm1.vspddata.text) <> "" Then
					arrParam(4) =  arrParam(4) & " and c.plant_cd= " & FilterVar(frm1.vspddata.text, "''", "S")
				END IF

				frm1.vspddata.Col = C_ItemGroupCd
				IF Trim(frm1.vspddata.text) <> "" Then
					arrParam(4) =  arrParam(4) & " and a.item_group_cd= " & FilterVar(frm1.vspddata.text, "''", "S")
				END IF			

				frm1.vspddata.Col = C_ItemCd
				IF Trim(frm1.vspddata.text) <> "" Then
					arrParam(4) =  arrParam(4) & " and a.item_cd LIKE " & FilterVar(Trim(frm1.vspddata.text) & "%", "''", "S") 
				END IF
			ELSE	'조건부 
				IF Trim(frm1.txtItemAcctCd.Value) <> "" Then
					arrParam(4) =  arrParam(4) & " and c.item_acct = " & FilterVar(frm1.txtItemAcctCd.Value, "''", "S")
				END IF
				
				IF Trim(frm1.txtPlantCd.Value) <> "" Then
					arrParam(4) =  arrParam(4) & " and c.plant_cd= " & FilterVar(frm1.txtPlantCd.Value, "''", "S")
				END IF

				IF Trim(frm1.txtItemGroupCd.Value) <> "" Then
					arrParam(4) =  arrParam(4) & " and a.item_group_cd= " & FilterVar(frm1.txtItemGroupCd.Value, "''", "S")
				END IF			
				
				IF Trim(frm1.txtItemCd.Value) <> "" Then
					arrParam(4) =  arrParam(4) & " and a.item_cd LIKE  " & FilterVar(Trim(frm1.txtItemCd.Value) & "%" , "''", "S")
				END IF						
			END IF				

			arrField(0) = "a.item_cd"		
			arrField(1) = "a.item_nm"		
		
			arrHeader(0) = "품목코드"		
			arrHeader(1) = "품목명"			
		case 9
			arrParam(0) = "원가요소팝업"	
			arrParam(1) = "C_COST_ELMT_S "
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""			
			arrParam(4) = ""			
			arrParam(5) = "원가요소"		
	
			arrField(0) = "cost_elmt_cd"		
			arrField(1) = "cost_elmt_nm"		
    
			arrHeader(0) = "원가요소"		
			arrHeader(1) = "원가요소명"	
		
			
	End Select
				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopup(iWhere,arrRet)
	End If
		
End Function



Function SetPopup(byVal iWhere,byval arrRet)
	with frm1
	select case iWhere
		case 1
			.vspdData.Col = c_ItemAcctCd
			.vspdData.Text = arrRet(0)
			
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = c_ItemAcctCd
			.vspdData.Action = 0
						
			.vspdData.Col = C_ItemAcctNm
			.vspdData.Text = arrRet(1)
			
				
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = C_ItemAcctNm
			.vspdData.Action = 0	
		case 2
			.txtItemAcctCd.value=  arrRet(0)
			.txtItemAcctNm.value=  arrRet(1)

		case 3
			.vspdData.Col = c_PlantCd
			.vspdData.Text = arrRet(0)
			
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = c_PlantCd
			.vspdData.Action = 0
						
			.vspdData.Col = c_PlantNm
			.vspdData.Text = arrRet(1)
			
				
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = c_PlantNm
			.vspdData.Action = 0	
		case 4
			.txtPlantCd.value=  arrRet(0)
			.txtPlantNm.value=  arrRet(1)
			
		case 5
			.vspdData.Col = C_ItemGroupCd
			.vspdData.Text = arrRet(0)
			
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = C_ItemGroupCd
			.vspdData.Action = 0
						
			.vspdData.Col = C_ItemGroupNm
			.vspdData.Text = arrRet(1)
			
				
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = C_ItemGroupNm
			.vspdData.Action = 0	
		case 6
			.txtItemGroupCd.value=  arrRet(0)
			.txtItemGroupNm.value=  arrRet(1)
			
		case 7
			.vspdData.Col = C_ItemCd
			.vspdData.Text = arrRet(0)
			
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = C_ItemCd
			.vspdData.Action = 0
						
			.vspdData.Col = C_ItemNm
			.vspdData.Text = arrRet(1)
			
				
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = C_ItemNm
			.vspdData.Action = 0	
		case 8
			.txtItemCd.value=  arrRet(0)
			.txtItemNm.value=  arrRet(1)					

		case 9
			.vspdData.Col = C_CostElmtCd
			.vspdData.Text = arrRet(0)
			
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = C_CostElmtCd
			.vspdData.Action = 0
						
			.vspdData.Col = C_CostElmtNm
			.vspdData.Text = arrRet(1)
			
				
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			.vspdData.Col = C_CostElmtNm
			.vspdData.Action = 0	
	end select							
	end with		
		
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
    
    select case Col
		case C_ItemAcctCd
			ggoSpread.SpreadUnLock    C_PlantCd, Row, C_PlantCd,Row
			ggoSpread.SpreadUnLock    C_PlantPb, Row, C_PlantPb,Row
		case C_PlantCd
			ggoSpread.SpreadUnLock    C_ItemGroupCd, Row, C_ItemGroupCd,Row
			ggoSpread.SpreadUnLock    C_ItemGroupPb, Row, C_ItemGroupPb,Row
		case C_ItemGroupCd
			ggoSpread.SpreadUnLock    C_ItemCd, Row, C_ItemCd,Row
			ggoSpread.SpreadUnLock    C_ItemPb, Row, C_ItemPb,Row
	End Select
End Sub


sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_ItemAcctPb
				.vspdData.Col = Col
				.vspdData.Row = Row

				.vspdData.Col = C_ItemAcctCD
				Call OpenPopup(1,.vspdData.Text)

			Case C_PlantPb
				.vspdData.Col = Col
				.vspdData.Row = Row

				.vspdData.Col = C_PlantCD
				Call OpenPopup(3,.vspdData.Text)

			Case C_ItemGroupPb
				.vspdData.Col = Col
				.vspdData.Row = Row

				.vspdData.Col = C_ItemGroupCd
				Call OpenPopup(5,.vspdData.Text)

			Case C_ItemPb
				.vspdData.Col = Col
				.vspdData.Row = Row

				.vspdData.Col = C_ItemCd
				Call OpenPopup(7,.vspdData.Text)
			Case C_CostElmtPb
				.vspdData.Col = Col
				.vspdData.Row = Row

				.vspdData.Col = C_CostElmtCd
				Call OpenPopup(9,.vspdData.Text)									
												
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
    	If lgIndPrevKey <> "" Then        
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

    frm1.vspdData.Col = C_ItemAcctcd
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ItemAcctNm
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_PlantCd
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_PlantNm
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ItemGroupCd
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ItemGroupNm
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ItemCd
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_ItemNm
    frm1.vspdData.Text = ""            
    frm1.vspdData.Col = C_CostElmtCd
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_CostElmtNm
    frm1.vspdData.Text = ""
    frm1.vspdData.Col = C_BasRate
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
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
 			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgItemAcctPrevKey=" & lgItemAcctPrevKey
			strVal = strVal & "&lgPlantPrevKey=" & lgPlantPrevKey
			strVal = strVal & "&lgItemGroupPrevKey=" & lgItemGroupPrevKey
			strVal = strVal & "&lgItemPrevKey=" & lgItemPrevKey
			strVal = strVal & "&lgCostElmtPrevKey=" & lgCostElmtPrevKey
			strVal = strVal & "&txtItemAcctCd=" & Trim(frm1.txtItemAcctCD.value)
			strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCD.value)
			strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.txtItemGroupCD.value)
			strVal = strVal & "&txtItemCd=" & Trim(frm1.txtItemCD.value)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    	Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgItemAcctPrevKey=" & lgItemAcctPrevKey
			strVal = strVal & "&lgPlantPrevKey=" & lgPlantPrevKey
			strVal = strVal & "&lgItemGroupPrevKey=" & lgItemGroupPrevKey
			strVal = strVal & "&lgItemPrevKey=" & lgItemPrevKey
			strVal = strVal & "&lgCostElmtPrevKey=" & lgCostElmtPrevKey
			strVal = strVal & "&txtItemAcctCd=" & Trim(frm1.hItemAcctCD.value)
			strVal = strVal & "&txtPlantCd=" & Trim(frm1.hPlantCD.value)
			strVal = strVal & "&txtItemGroupCd=" & Trim(frm1.hItemGroupCD.value)
			strVal = strVal & "&txtItemCd=" & Trim(frm1.hItemCD.value)
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
	Dim strVal
	
    DbSave = False        
    
	If LayerShowHide(1) = False Then
		Exit Function
	End If	
	
	With frm1
		.txtMode.value = Parent.UID_M0002
		
		lGrpCnt = 1

		strVal = ""
    
		iColSep = Parent.gColSep
		iRowSep = Parent.gRowSep	

		For lRow = 1 To .vspdData.MaxRows

			.vspdData.Row = lRow
			.vspdData.Col = 0
        
			Select Case .vspdData.Text

	            Case ggoSpread.InsertFlag		

					strVal = strVal & "C" & iColSep & lRow & iColSep

					.vspdData.Col = C_ItemAcctCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_PlantCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ITemGroupCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ItemCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_CostElmtCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					
					.vspdData.Col = C_BasRate	
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep

					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.UpdateFlag		

					strVal = strVal & "U" & iColSep & lRow & iColSep	

					.vspdData.Col = C_ItemAcctCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_PlantCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ITemGroupCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ItemCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_CostElmtCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					
					.vspdData.Col = C_BasRate	
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep

					lGrpCnt = lGrpCnt + 1

	            Case ggoSpread.DeleteFlag		

					strVal = strVal & "D" & iColSep & lRow & iColSep	

					.vspdData.Col = C_ItemAcctCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_PlantCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ITemGroupCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_ItemCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep

					.vspdData.Col = C_CostElmtCd	
					strVal = strVal & Trim(.vspdData.Text) & iColSep
					
					.vspdData.Col = C_BasRate	
					strVal = strVal & UNIConvNum(Trim(.vspdData.Text),0) & iRowSep

					lGrpCnt = lGrpCnt + 1
                
	        End Select
                
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>원가요소별 표준Rate 등록</font></td>
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
									<TD CLASS="TD5">품목계정</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemAcctCd" MAXLENGTH="2" SIZE=15  ALT ="품목계정" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemAcctCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(2,frm1.txtItemAcctCd.Value)">
														<INPUT NAME="txtItemAcctNm" MAXLENGTH="30" SIZE=25  ALT ="품목계정명" tag="14X"></TD>

									<TD CLASS="TD5">공장</TD>
									<TD CLASS="TD6"><INPUT CLASS="clstxt" NAME="txtPlantCD" MAXLENGTH="4" SIZE=15  ALT ="공장" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(4,frm1.txtPlantCd.Value)">
														<INPUT NAME="txtPlantNM" MAXLENGTH="30" SIZE=25  ALT ="공장명" tag="14X"></TD>
										
								</TR>
								<TR>
									<TD CLASS="TD5">품목그룹</TD>
									<TD CLASS="TD6"><INPUT  NAME="txtItemGroupCd" MAXLENGTH="10" SIZE=15  ALT ="품목그룹" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProcurType" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(6,frm1.txtItemGroupCd.Value)">
														<INPUT NAME="txtItemGroupNm" MAXLENGTH="30" SIZE=35  ALT ="품목그룹명" tag="14X"></TD>
										
									<TD CLASS="TD5">품목</TD>
									<TD CLASS="TD6"><INPUT  NAME="txtItemCD" MAXLENGTH="18" SIZE=15  ALT ="품목" tag="1XXXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(8,frm1.txtItemCd.Value)">
														<INPUT NAME="txtItemNm" MAXLENGTH="30" SIZE=35  ALT ="품목명" tag="14X"></TD>
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
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="hItemAcctCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



