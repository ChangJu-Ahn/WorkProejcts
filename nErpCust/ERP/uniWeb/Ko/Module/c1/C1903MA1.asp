
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 원가수불유형등록 
'*  3. Program ID           : c1903ma1.asp
'*  4. Program Name         : 원가수불유형등록 
'*  5. Program Desc         : 원가수불유형등록 
'*  6. Modified date(First) : 2003/04/22
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

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "C1903MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim C_InOut                                  '☆: Spread Sheet의 Column별 상수 
Dim C_TrnsType 													
Dim C_TrnsTypeNm
Dim C_MovType
Dim C_MovTypeNm
Dim C_ItemAcct
Dim C_ItemAcctPb
Dim	C_ItemAcctNm
Dim C_PostingFlag
Dim C_PriceFlag
Dim C_CostFlag
Dim C_DiffFlag
Dim C_ToCdFlag	
Dim C_Remark
Dim C_APPD_FLAG
'Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 

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
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_InOut			= 1                                 
	C_TrnsType 		= 2											
	C_TrnsTypeNm	= 3
	C_MovType		= 4
	C_MovTypeNm		= 5
	C_ItemAcct		= 6
	C_ItemAcctPb	= 7
	C_ItemAcctNm	= 8
	C_PostingFlag	= 9
	C_PriceFlag		= 10
	C_CostFlag		= 11
	C_DiffFlag		= 12
	C_ToCdFlag		= 13
	C_Remark		= 14
	C_APPD_FLAG		= 15
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	    
	.MaxCols = C_APPD_FLAG + 1						
 	
    .Col = .MaxCols							
    .ColHidden = True
    
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    

	
	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	Call GetSpreadColumnPos("A")
	
	

	' ColumnPosition Header
    ggoSpread.SSSetEdit		C_InOut				,"입/출구분"	,10,,,,2
	ggoSpread.SSSetEdit	    C_TrnsType			,"수불구분"		,10,,,,2
	ggoSpread.SSSetEdit		C_TrnsTypeNm		,"수불구분명"	,15,,,,2
	ggoSpread.SSSetEdit	    C_MovType			,"수불유형"		,10,,,,2
	ggoSpread.SSSetEdit		C_MovTypeNm			,"수불유형명"	,30,,,,2
	ggoSpread.SSSetEdit	    C_ItemAcct			,"품목계정"		,10,,,,2
	ggoSpread.SSSetButton	C_ItemAcctPb
	ggoSpread.SSSetEdit		C_ItemAcctNm		,"품목계정명"	,10,,,,2	
	ggoSpread.SSSetEdit	    C_PostingFlag		,"POSTING여부"	,10,,,,2
	
	ggoSpread.SSSetCheck	C_PriceFlag			,"단가계산반영여부"	,10	, ,"",true 
	ggoSpread.SSSetCheck	C_CostFlag			,"재료비반영여부"	,10	, ,"",true 
	ggoSpread.SSSetCheck	C_DiffFlag			,"차이반영여부"		,10	, ,"",true 
	ggoSpread.SSSetCheck	C_ToCdFlag			,"수불처별조정여부"	,10	, ,"",true 
	
	ggoSpread.SSSetEdit	    C_Remark			,"Remark"		,30,,,30,2
	ggoSpread.SSSetEdit	    C_APPD_FLAG			,"APPD_FLAG"		,10,,,30,2
	
	Call ggoSpread.SSSetColHidden(C_APPD_FLAG,C_APPD_FLAG,True)
	Call ggoSpread.SSSetColHidden(C_ToCdFlag,C_ToCdFlag,True)
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_InOut, -1, C_InOut	
	ggoSpread.SpreadLock C_TrnsType, -1, C_TrnsType
	ggoSpread.SpreadLock C_TrnsTypeNm, -1, C_TrnsTypeNm
	ggoSpread.SpreadLock C_MovType, -1, C_MovType
	ggoSpread.SpreadLock C_MovTypeNm, -1, C_MovTypeNm
	ggoSpread.SpreadLock C_ItemAcct, -1, C_ItemAcct
	ggoSpread.SpreadLock C_ItemAcctPb, -1, C_ItemAcctPb
	ggoSpread.SpreadLock C_ItemAcctNm, -1, C_ItemAcctNm
	ggoSpread.SpreadLock C_PostingFlag, -1, C_PostingFlag
	
	'ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1

    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor2(ByVal iRow)

    With frm1
    
    .vspdData.ReDraw = False
    
    ggoSpread.Source = .vspdData
    .vspdData.Row = iRow
    .vspdData.Col = C_APPD_FLAG
	    
    IF Trim(.vspdData.Text) = "Y" Then
	    ggoSpread.SpreadLock C_InOut	,iRow,	C_APPD_FLAG	,iRow
	ELSE
		ggoSpread.SpreadLock C_InOut, iRow, C_MovTypeNm, iRow
		ggoSpread.SpreadLock C_ItemAcctNm, iRow, C_PostingFlag, iRow			
		
		.vspdData.Row = iRow
		.vspdData.Col = C_TrnsType		
		
		' 단가계산반영 Protected 처리 
		IF UCase(Trim(.vspdData.text)) <> "OR" and UCase(Trim(.vspdData.text)) <> "ST" Then
			ggoSpread.SSSetProtected C_PriceFlag,	iRow,	iRow
		END IF

		.vspdData.Row = iRow
		.vspdData.Col = C_TrnsType		
		
		' 재료비반영 Protected 처리 
		IF UCase(Trim(.vspdData.text)) <> "OI"  Then
			ggoSpread.SSSetProtected C_CostFlag,	iRow,	iRow
		END IF

		.vspdData.Row = iRow
		.vspdData.Col = C_TrnsType		

		' 차이반영 Protected 처리 
		IF UCase(Trim(.vspdData.text)) <> "OI" and UCase(Trim(.vspdData.text)) <> "ST"  Then
			ggoSpread.SSSetProtected C_DiffFlag,	iRow,	iRow	
		END IF


		.vspdData.Row = iRow
		.vspdData.Col = C_MovType

		' I차이반영 Protected 처리 
		IF UCase(Trim(.vspdData.text)) = "I99" or UCase(Trim(.vspdData.text)) = "I9Z" or UCase(Trim(.vspdData.text)) = "IZZ" Then
			ggoSpread.SSSetProtected C_CostFlag,	iRow,	iRow	
			ggoSpread.SSSetProtected C_DiffFlag,	iRow,	iRow	
		END IF


		'수불처별 조정 Protected 처리 
		.vspdData.Row = iRow
		.vspdData.Col = C_TrnsType		
		IF UCase(Trim(.vspdData.text)) =  "PR" or UCase(Trim(.vspdData.text)) = "OR" or UCase(Trim(.vspdData.text)) = "MR" Then
			ggoSpread.SSSetProtected C_ToCdFlag,	iRow,	iRow	
		END IF
	
		
	END IF

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
			C_InOut				= iCurColumnPos(1)
			C_TrnsType			= iCurColumnPos(2)
			C_TrnsTypeNm		= iCurColumnPos(3)    
			C_MovType			= iCurColumnPos(4)
			C_MovTypeNm			= iCurColumnPos(5)
			C_ItemAcct			= iCurColumnPos(6)
			C_ItemAcctPb		= iCurColumnPos(7)
			C_ItemAcctNm		= iCurColumnPos(8)
			C_PostingFlag		= iCurColumnPos(9)
			C_PriceFlag			= iCurColumnPos(10)
			C_CostFlag			= iCurColumnPos(11)
			C_DiffFlag			= iCurColumnPos(12)
			C_ToCdFlag			= iCurColumnPos(13)
			C_Remark			= iCurColumnPos(14)
			C_APPD_FLAG			= iCurColumnPos(15)
    End Select    
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
      
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    'Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	
    Call SetToolbar("110110010011111")	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call CookiePage (0)                                                              '☜: Check Cookie
			
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
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If DbQuery() = False Then                                                      '☜: Query db data
       Exit Function
    End If
	
   If Err.number = 0 Then	
       FncQuery = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function
	
'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
   If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If

	'-----------------------
    'Delete function call area
    '-----------------------%>
    IntRetCd = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")		           
	If IntRetCd = vbNo Then
		Exit Function
	End If

    If DbDelete = False Then
		Exit Function
	End If	
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDelete = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                      '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

   If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			
			SetSpreadColor2 .ActiveRow
			ggoSpread.SpreadUnLock C_ItemAcct,.ActiveRow,C_ItemAcct,.ActiveRow
			ggoSpread.SSSetRequired C_ItemAcct,	.ActiveRow,	.ActiveRow
			ggoSpread.SpreadUnLock C_ItemAcctPb,.ActiveRow,C_ItemAcctPb,.ActiveRow

			.Col  = C_ItemAcct
			.Row  = .ActiveRow
			.Text = ""
        
			.Col  = C_ItemAcctNm
			.Text = ""
        
			.Col  = C_TrnsType
			
			IF	.Text = "OR" Then
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadUnLock C_PriceFlag,.ActiveRow,C_PriceFlag,.ActiveRow

			ELSEIF .Text = "OI" Then 
				ggoSpread.Source = frm1.vspdData
				ggoSpread.SpreadUnLock C_CostFlag,.ActiveRow,C_CostFlag,.ActiveRow
				ggoSpread.SpreadUnLock C_DiffFlag,.ActiveRow,C_DiffFlag,.ActiveRow
			ELSEIF .Text = "ST" Then
				ggoSpread.SpreadUnLock C_PriceFlag,.ActiveRow,C_PriceFlag,.ActiveRow
				ggoSpread.SpreadUnLock C_DiffFlag,.ActiveRow,C_DiffFlag,.ActiveRow
			END IF   

				
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field

	
	'---------------------------------------------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    On Error Resume Next
    
    Dim iDx
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
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
        'SetSpreadColor2 .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
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
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

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
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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
    Dim iRow
    
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
    
   
    IF frm1.vspdData.MaxRows > 0 Then
		For iRow = 1 To frm1.vspdData.MaxRows
			Call SetSpreadColor2(iRow)
		Next
	END IF
End Sub


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery()

	Dim strVal
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    
    DbQuery = False                                                              '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
  
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtVersion="		 & Trim(frm1.txtVersion.value)               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
    End With
		
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave()
		
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
    Dim iColSep 
    Dim iRowSep  

    On Error Resume Next
    DbSave = False                                                               '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call DisableToolBar(Parent.TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
		
    Frm1.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
    ggoSpread.Source = frm1.vspdData

    strVal = ""
    lGrpCnt = 1
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	  


	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 

               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                     strVal = strVal & "U"                       & iColSep
                                                     
													 strVal = strVal & lRow                      & iColSep
                    .vspdData.Col = C_TrnsType     : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_MovType      : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    
                    .vspdData.Col = C_PriceFlag   : 
         				  IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF

                    .vspdData.Col = C_CostFlag   : 
         				  IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF
                    .vspdData.Col = C_DiffFlag   : 
         				  IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF
                    .vspdData.Col = C_ToCdFlag   : 
         				  IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF
                    .vspdData.Col = C_Remark      : strVal = strVal & Trim(.vspdData.Text)      & iColSep 
                    .vspdData.Col = C_ItemAcct     : strVal = strVal & Trim(.vspdData.Text)      & iRowSep
                   lGrpCnt = lGrpCnt + 1
				Case ggoSpread.InsertFlag                     

                                                     strVal = strVal & "C"                       & iColSep
                                                     
													 strVal = strVal & lRow                      & iColSep
                    .vspdData.Col = C_TrnsType     : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_MovType      : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_PriceFlag   : 
         				  IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF

                    .vspdData.Col = C_CostFlag   : 
         				  IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF
                    .vspdData.Col = C_DiffFlag   : 
         				  IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF
                    .vspdData.Col = C_ToCdFlag   : 
         				  IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF
                    .vspdData.Col = C_Remark      : strVal = strVal & Trim(.vspdData.Text)      & iColSep 
                    .vspdData.Col = C_ItemAcct     : strVal = strVal & Trim(.vspdData.Text)      & iRowSep
                    
                   lGrpCnt = lGrpCnt + 1
           End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strVal
		
	End With
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

   If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
    LayerShowHide(1)						
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003					
    strVal = strVal & "&txtVersion=" & Trim(frm1.hVersion.value)		
	
	Call RunMyBizASP(MyBizASP, strVal)									
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
     If Err.number = 0 Then	 
       DbDelete = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
	Dim iRow
	
    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("110110010011111")	
	Frm1.vspdData.Focus
	
	IF frm1.vspdData.MaxRows > 0 Then
		For iRow = 1 To frm1.vspdData.MaxRows
			Call SetSpreadColor2(iRow)
		Next
	END IF
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery() = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "1")										     '⊙: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	Call DisplayMsgBox("800154", "x","x","x")					 '☜: 작업이 완료되었습니다 
	frm1.txtVersion.focus()
	
	Set gActiveElement = document.ActiveElement   
 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Name : OpenPlant
' Desc : 공장 팝업 
'========================================================================================================
Function OpenVersion()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "VERSION팝업"	
	arrParam(1) = "C_MOVETYPE_CONFIGURATION"
	arrParam(2) = Trim(frm1.txtVersion.Value)
	arrParam(3) = ""											
	arrParam(4) = ""												
	arrParam(5) = "Version"							
	
    arrField(0) = "VER_CD"						
    
    
    arrHeader(0) = "Version"				
    	
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtVersion.focus
		Exit Function
	Else
		Call SetVersion(arrRet)
	End If
		
End Function


Function SetVersion(byval arrRet)
	frm1.txtVersion.focus
	frm1.txtVersion.Value    = arrRet(0)		
End Function

Function OpenItemAcct(ByVal strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True


	arrParam(0) = "품목계정팝업"  	
	arrParam(1) = "B_Minor a,b_item_acct_inf b"					
	arrParam(2) = strCode
	arrParam(3) = ""								
	arrParam(4) = "a.Major_Cd='P1001' and a.minor_cd = b.item_acct and b.item_acct_group <> '6MRO'"		
	arrParam(5) = "품목계정"    				

	arrField(0) = "a.MINOR_CD"				
	arrField(1) = "a.MINOR_NM"					
    
	arrHeader(0) = "품목계정"	  
	arrHeader(1) = "품목계정명"		
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemAcct(arrRet)
	End If
		
End Function


Function SetItemAcct(byval arrRet)

	With frm1
	        .vspdData.ReDraw = False
			.vspdData.Col = C_ItemAcct
			.vspdData.Text = arrRet(0)
			Call vspddata_Change(.vspddata.col, .vspddata.row)
			
			.vspdData.Col = C_ItemAcctNm
			.vspdData.Text = arrRet(1)
			Call vspddata_Change(.vspddata.col, .vspddata.row)
				
			.vspdData.Col = C_ItemAcct
			'.vspdData.Action = 0			

            .vspdData.ReDraw = True
		    .vspdData.Focus
	
	End With

End Function
'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim StrCode1
	Dim intRetCD

	With frm1
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_ItemAcctPb
				.vspdData.Col = C_ItemAcct
				.vspdData.Row = Row
				
				
				Call OpenItemAcct(.vspdData.Text)


		End Select
           	Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(Col, Row)


    Call SetPopupMenuItemInf("0001111111")
    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

	
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort	Col			'Sort in ascending
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col	,lgSortKey	'Sort in descending
            lgSortKey = 1
        End If
    Else
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
'   Event Name : vspdData_MouseDown
'   Event Desc :
'========================================================================================================
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
'    If Col <= C_CE_NM Or NewCol <= C_CE_NM Then
'        Cancel = True
'        Exit Sub
'    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
    

End Sub


'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery() = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" bgColor=White text=Black>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>원가수불유형등록</font></td>
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
									<TD CLASS="TD5">Version</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtVersion" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="Version"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVersion" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenVersion()" ID="btnVersion"></TD>
									<TD CLASS="TD5"></TD>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c1903ma1_OBJECT1_vspdData.js'></script>
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
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hVersion" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

