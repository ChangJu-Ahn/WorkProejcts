<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         : Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
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
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: Turn on the Option Explicit option.
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "s1113mb2_ko441.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables For spreadsheet
'========================================================================================================

Dim C_Bp_Cd
Dim C_Bp_Cd_Popup
Dim C_Bp_Nm
Dim C_Item_Cd
Dim C_Item_Cd_Popup
Dim C_Item_Nm
Dim C_Item_Cd_Spec
Dim C_Deal_type
Dim C_Deal_type_Popup
Dim C_Deal_type_nm
Dim C_Pay_meth
Dim C_Pay_meth_Popup
Dim C_Pay_meth_nm
Dim C_Valid_from_dt
Dim C_Unit
Dim C_Unit_Popup
Dim C_Cur
Dim C_Cur_Popup
Dim C_Item_Price
Dim C_Price_Flag
Dim C_Price_Flag_Nm
Dim C_Remark
Dim C_ChgFlg
	
Dim gblnWinEvent

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          
<% 
   BaseDate     = GetSvrDate                                                                  'Get DB Server Date
%>   
dim lastdate
dim firstdate
dim ExampleDate
LastDate    = UNIGetLastDay ("<%=BaseDate%>",parent.gServerDateFormat)                                  'Last  day of this month
FirstDate   = UNIGetFirstDay("<%=BaseDate%>",parent.gServerDateFormat)                                  'First day of this month
ExampleDate = UniDateAdd("m", -2, "<%=BaseDate%>",parent.gServerDateFormat)
ExampleDate = UNIConvDateAToB("<%=BaseDate%>" ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Group-1
'========================================================================================================
'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
	C_Bp_Cd               = 1
	C_Bp_Cd_Popup         = 2
	C_Bp_Nm               = 3
	C_Item_Cd             = 4 
	C_Item_Cd_Popup       = 5
	C_Item_Nm             = 6
	C_Item_Cd_Spec        = 7 
	C_Deal_type           = 8
	C_Deal_type_Popup     = 9
	C_Deal_type_nm        = 10
	C_Pay_meth            = 11
	C_Pay_meth_Popup      = 12
	C_Pay_meth_nm         = 13
	C_Valid_from_dt       = 14
	C_Unit                = 15
	C_Unit_Popup          = 16
	C_Cur                 = 17
	C_Cur_Popup           = 18
	C_Item_Price          = 19
	C_Price_Flag          = 20
	C_Price_Flag_Nm       = 21
	C_Remark			  = 22
	C_ChgFlg			  = 23
End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
   		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()

	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Ucase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtPlantCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
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
	<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(ByVal Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(ByVal pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   Select Case pOpt
       Case "MQ"
                  lgKeyStream = frm1.txtPlantCd.Value  & Parent.gColSep       'You Must append one character(Parent.gColSep)
       Case "MN"
                  lgKeyStream = Frm1.htxtPlantCd.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
   End Select                 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

	
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBoxf
'========================================================================================================
Sub InitComboBox()    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Function Name : InitSpreadComboBox
' Function Desc :
'========================================================================================================
Sub InitSpreadComboBox()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Dim strCD
    Dim strVal		    
	
	'****************************
	'List Minor code(Price Flag Code)
	'****************************	
	strCD = "T" & vbTab & "F" 
	strVal = "진단가" & vbTab & "가단가"
	ggoSpread.Source = frm1.vspdData
   
    ggoSpread.SetCombo Replace(strCD ,Chr(11),vbTab), C_Price_Flag
    ggoSpread.SetCombo Replace(strVal,Chr(11),vbTab), C_Price_Flag_Nm
    
   
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 

	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
'			.Col = C_StudyOnOffCd  :  intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
'			.Col = C_StudyOnOffNm  :  .Value = intindex					
		Next	
	End With
End Sub

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()

	With frm1.vspdData
       ggoSpread.Source = frm1.vspdData
       ggoSpread.Spreadinit "V20021105",, parent.gAllowDragDropSpread
       .ReDraw = false
       .MaxCols   = C_ChgFlg + 1                                                  ' ☜:☜: Add 1 to Maxcols
       Call ggoSpread.ClearSpreadData()
       Call AppendNumberPlace("6","4","2")
	   .Col       = .MaxCols        : .ColHidden = True
       .MaxRows = 0                                                                  ' ☜: Clear spreadsheet data 
       Call GetSpreadColumnPos("A")
				
 			      
		ggoSpread.SSSetEdit     C_Bp_Cd,                "고객" ,15, 0,,10,2
		ggoSpread.SSSetButton   C_Bp_Cd_Popup    
		ggoSpread.SSSetEdit     C_Bp_Nm,                "고객명", 25, 0 
		ggoSpread.SSSetEdit     C_Item_Cd,              "품목" ,15, 0,,18,2
		ggoSpread.SSSetButton   C_Item_Cd_Popup    
		ggoSpread.SSSetEdit     C_Item_Nm,              "품목명", 25, 0 
		ggoSpread.SSSetEdit     C_Item_Cd_Spec,         "규격", 20, 0         
		ggoSpread.SSSetEdit     C_Deal_type,            "판매유형", 15, 0,,15,2        
		ggoSpread.SSSetButton   C_Deal_type_Popup
		ggoSpread.SSSetEdit     C_Deal_type_Nm,         "판매유형명", 15, 0
		ggoSpread.SSSetEdit     C_Pay_meth,            "결제방법", 10, 0,,5,2
		ggoSpread.SSSetButton   C_Pay_meth_Popup
		ggoSpread.SSSetEdit     C_Pay_meth_Nm,         "결제방법명", 15, 0
		ggoSpread.SSSetDate     C_Valid_from_dt,        "적용일", 10, 2, Parent.gDateFormat   
		ggoSpread.SSSetEdit     C_unit,                 "단위", 10, 0,,3,2
		ggoSpread.SSSetButton   C_unit_Popup                 
		ggoSpread.SSSetEdit     C_Cur,                  "화폐", 10, 0,,3,2
		ggoSpread.SSSetButton   C_Cur_Popup 
		ggoSpread.SSSetFloat    C_Item_Price,           "단가",15, "C" , ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"	
		
		ggoSpread.SSSetCombo    C_Price_Flag,           "단가구분",10, 0		
		ggoSpread.SSSetEdit    C_Price_Flag_Nm,           "단가구분명",15, 0
		ggoSpread.SSSetEdit     C_Remark,				"비고", 30, 0,,240
			     
		ggoSpread.SSSetEdit     C_ChgFlg, "Chgfg", 1, 2  

		call ggoSpread.MakePairsColumn(C_Bp_Cd,C_Bp_Cd_Popup)
		call ggoSpread.MakePairsColumn(C_Item_Cd,C_Item_Cd_Popup)
		call ggoSpread.MakePairsColumn(C_Deal_type,C_Deal_type_Popup)
		call ggoSpread.MakePairsColumn(C_Pay_meth,C_Pay_meth_Popup)
		call ggoSpread.MakePairsColumn(C_unit,C_unit_Popup)
		call ggoSpread.MakePairsColumn(C_Cur,C_Cur_Popup)

		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)

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
                                 'Col-1             Row-1       Col-2           Row-2   
    ggoSpread.spreadlock    C_Bp_Nm, -1
    ggoSpread.spreadUnlock  C_Item_Cd , -1
    ggoSpread.spreadlock    C_Item_Nm, -1
    ggoSpread.spreadlock    C_Item_Cd_Spec, -1
    ggoSpread.spreadUnlock  C_Deal_type , -1
    ggoSpread.spreadlock    C_Deal_type_nm, -1
    ggoSpread.spreadUnlock  C_Pay_meth  , -1
    ggoSpread.spreadlock    C_Pay_meth_nm , -1
    ggoSpread.spreadUnlock  C_Valid_from_dt, -1    
    ggoSpread.spreadUnlock	C_Remark, -1 
    
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1    
    
    .vspdData.ReDraw = False
    
    ggoSpread.SSSetRequired    C_Bp_Cd,              pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Bp_Nm  ,            pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Item_Cd,            pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Item_Nm  ,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Item_Cd_Spec  ,     pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Deal_type,          pvStartRow, pvEndRow
    ggoSpread.SSSetProtected   C_Deal_type_Nm,       pvStartRow, pvEndRow
    ggoSpread.SSSetRequired    C_Pay_meth,           pvStartRow, pvEndRow
	ggoSpread.SSSetProtected   C_Pay_meth_nm,        pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Valid_from_dt,      pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Unit,               pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Cur,                pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Item_Price,         pvStartRow, pvEndRow
	ggoSpread.SSSetRequired    C_Price_Flag,         pvStartRow, pvEndRow
	'ggoSpread.SSSetRequired    C_Price_Flag_Nm,         pvStartRow, pvEndRow
	ggoSpread.SSSetProtected   C_Price_Flag_Nm  ,     pvStartRow, pvEndRow
	    
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
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Bp_Cd				= iCurColumnPos(1)
			C_Bp_Cd_Popup       = iCurColumnPos(2)
			C_Bp_Nm				= iCurColumnPos(3)    
			C_Item_Cd			= iCurColumnPos(4)
			C_Item_Cd_Popup     = iCurColumnPos(5)
			C_Item_Nm			= iCurColumnPos(6)
			C_Item_Cd_Spec		= iCurColumnPos(7)
			C_Deal_type			= iCurColumnPos(8)
			C_Deal_type_Popup	= iCurColumnPos(9)
			C_Deal_type_nm		= iCurColumnPos(10)
			C_Pay_meth			= iCurColumnPos(11)
			C_Pay_meth_Popup    = iCurColumnPos(12)
			C_Pay_meth_nm		= iCurColumnPos(13)
			C_Valid_from_dt		= iCurColumnPos(14)
			C_Unit				= iCurColumnPos(15)
			C_Unit_Popup		= iCurColumnPos(16)
			C_Cur				= iCurColumnPos(17)
			C_Cur_Popup			= iCurColumnPos(18)
			C_Item_Price		= iCurColumnPos(19)
			C_Price_Flag		= iCurColumnPos(20)
			C_Price_Flag_Nm		= iCurColumnPos(21)
			C_Remark			= iCurColumnPos(22)					
			C_ChgFlg			= iCurColumnPos(23)
    End Select    
End Sub

'========================================================================================================
'========================================================================================================
'                        5.2 Common Group-2
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
      
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    Call InitSpreadSheet                                                             'Setup the Spread sheet
	Call InitVariables
    Call SetDefaultVal

	frm1.txtPlantCd.focus
'	Call SetToolBar("11000000000011")                                              '☆: Developer must customize
    Call SetToolBar("1110110100101111")          '⊙: 버튼 툴바 제어 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitSpreadComboBox()
	Call CookiePage(0)       
	                                                       
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
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False															  '☜: Processing is NG
	
    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					  '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										  '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
        															
    If Not chkField(Document, "1") Then									          '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                            '⊙: Initializes local global variables

	If DbQuery("MQ") = False Then                                                 '☜: Query db data
       Exit Function
    End If
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
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
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
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
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False                                                               '☜: Processing is NG
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                       '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                            '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                          '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                        '☜: Query db data
       Exit Function
    End If

    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

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

	With Frm1
        .vspdData.Col  = C_SchoolCD
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG

    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
     Frm1.vspdData.Row = frm1.vspdData.ActiveRow
'   	 Frm1.vspdData.Col = C_StudyOnOffCD :     iDx = Frm1.vspdData.value
'     Frm1.vspdData.Col = C_StudyOnOffNM :     Frm1.vspdData.value = iDx
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


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
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

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
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then	 
       FncPrev = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then	 
       FncNext = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_MULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

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
    Call ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub


'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitSpreadComboBox()
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		              '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Group-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery(pDirect)
msgbox " DbQuery "
	Dim strVal
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    DbQuery = False                                                               '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                                '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                         '☜: Show Processing Message
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call MakeKeyStream(pDirect)
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="       & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="  & lgKeyStream         '☜: Query Key
        strVal = strVal     & "&txtMaxRows="    & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey        '☜: Next key tag
    End With
    '--------- Developer Coding Part (End) ------------------------------------------------------------

msgbox " DbQuery 100 strVal : " & strVal
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    If Err.number = 0 Then	 
       DbQuery = True                                                             '☜: Processing is OK
    End If

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

    DbSave = False

    If LayerShowHide(1) = False Then
         Exit Function
    End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

    With frm1


    strVal = ""
    strDel = ""
    lGrpCnt = 1

       For lRow = 1 To .vspdData.MaxRows

           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text

                
               Case ggoSpread.InsertFlag                                      '☜: Update
			strVal = ""
			
                                                   strVal = strVal & "C" & parent.gColSep
                                                   strVal = strVal & lRow & parent.gRowSep
'                    .vspdData.Col = C_bp_cd		    :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                    .vspdData.Col = C_Deal_type	    :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
'                    .vspdData.Col = C_Item_Cd	    :	strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    
                    'msgbox strVal

                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
			strVal = ""
			
                                                   strVal = strVal & "U" & parent.gColSep
                                                   strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_REMARK1			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ISSUE_REQ_NO	    :	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                                            			:	strVal = strVal & Trim(.txtPlantCd.value) & parent.gColSep
                    .vspdData.Col = C_ITEM_SEQ			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REQ_QTY			:	strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ISSUE_QTY			:	strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    
                    'msgbox strVal

                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

msgbox " 400"
       .txtMode.value        = parent.UID_M0002
msgbox " 401"
       .txtUpdtUserId.value  = "unierp" 'parent.gUsrID
msgbox " 402"
       .txtInsrtUserId.value = parent.gUsrID
msgbox " 403"
       .txtMaxRows.value     = lGrpCnt - 1
msgbox " 404"
       .txtSpread.value      = strDel & strVal

msgbox " 405"

    'Call RunMyBizASP(MyBizASP, strVal)   
    Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    End With

    DbSave = True

End Function

'========================================================================================================
' Name : DbDelete
' Desc : This sub is called by FncDelete
'========================================================================================================
Function DbDelete()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbDelete = False                                                              '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	 
       DbDelete = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Sub DbQueryOk()
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'Call SetToolbar("11000000001111")                                              '☆: Developer must customize
	Call SetToolBar("1110111100111111") 
	Frm1.vspdData.Focus
    Call InitData()
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ggoOper.LockField(Document, "Q")
	
    Set gActiveElement = document.ActiveElement   
    ggospread.source = frm1.vspdData
End Sub
	
'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Sub DbSaveOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    
	                                              '☆: Developer must customize
    If DbQuery("MQ") = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement   
End Sub
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Sub DbDeleteOk()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : OpenReference
' Desc : developer describe this line 
'========================================================================================================
Function OpenReference()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 		
	MsgBox "You need to code this part.",,Parent.gLogoName 
	'------ Developer Coding part (End)    -------------------------------------------------------------- 
End Function

'========================================================================================================
' Name : OpenZipCode()
' Desc : developer describe this line 
'========================================================================================================
Function OpenZipCode(ZipCode,Row)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "우편번호 팝업"                             ' Popup Name
	arrParam(1) = "ADDRESS"                                       ' Table Name
	arrParam(2) = ZipCode                                         ' Code Condition
	arrParam(3) = ""                                              ' Name Cindition
	arrParam(4) = ""                                              ' Where Condition
	arrParam(5) = "우편코드"
	
    arrField(0) = "ZipCd"                                         ' Field명(0)
    arrField(1) = "Address"                                       ' Field명(1)
    
    arrHeader(0) = "우편번호"	                              ' Header명(0)
    arrHeader(1) = "주소"                                     ' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SubSetZipCode(arrRet,Row)
	End If	
	
End Function

'========================================================================================================
'
'
'========================================================================================================
Sub SubSetZipCode(arrRet,Row)

	With frm1.vspdData 
          .Row  = Row
          .Col  = C_ZipCode
          .Text = arrRet(0)
          .Col  = C_AddressNm
          .Text = arrRet(1)
          ggoSpread.Source = frm1.vspdData
          ggoSpread.UpdateRow frm1.vspdData.Row
          
	End With

End Sub

'========================================================================================================
' Name : OpenSchoolCd()
' Desc : developer describe this line 
'========================================================================================================
Function OpenSchoolCd(pOpt)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "학교코드 팝업"                             ' Popup Name
	arrParam(1) = "SCHOOL"                                        ' Table Name
	arrParam(2) = frm1.txtSchoolCdC.value                         ' Code Condition
	arrParam(3) = ""                                              ' Name Cindition
	arrParam(4) = ""                                              ' Where Condition
	arrParam(5) = "학교코드"
	
    arrField(0) = "SchoolCD"                                      ' Field명(0)
    arrField(1) = "SchoolNM"                                      ' Field명(1)
    arrField(2) = "F2" & Parent.gColSep & "DonatedMoney"                 ' Field명(2)
    arrField(3) = "DD" & Parent.gColSep & "FoundedDT"                    ' Field명(3)
    
    arrHeader(0) = "학교코드"	                              ' Header명(0)
    arrHeader(1) = "학교코드명"                               ' Header명(1)
    arrHeader(2) = "기부금"                                   ' Header명(1)
    arrHeader(3) = "설립일"                                   ' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SubSetSchoolInf(arrRet,pOpt)
	End If	
	
    Call SetFocusToDocument("M")	                              ' This move focus to Document . You must not delete this line

    Select Case pOpt
         Case "C" : Frm1.txtSchoolCdC.focus
         Case "D" : Frm1.txtSchoolCdD.focus
   End Select          
    
End Function

'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetSchoolInf(arrRet,pOpt)
    Select Case pOpt
         Case "C"
            With Frm1
              .txtSchoolCdC.value = arrRet(0)
              .txtSchoolNmC.value = arrRet(1)		
            End With
         Case "D"
            With Frm1
              .txtSchoolCdD.value = arrRet(0)
              .txtSchoolNmD.value = arrRet(1)		
            End With
   End Select          
End Sub


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)

	With frm1.vspdData 
 
		ggoSpread.Source = frm1.vspdData

		If Row > 0 Then
			Select Case Col
			Case C_Bp_Cd_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenBp_Cd (.text)
			Case C_Item_Cd_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenItem_Cd (.text) 
			Case C_Deal_type_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenDeal_type (.Text)
			Case C_Pay_meth_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenPay_meth (.Text)
			Case C_Unit_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenUnit (.Text)
			Case C_Cur_Popup
				.Col = Col - 1
				.Row = Row
				Call OpenCur (.Text)
			End Select
		 
			Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")  
		End If

	End With
End Sub

'===========================================================================
 Function  OpenCur(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "화폐"       <%' 팝업 명칭 %>
  arrParam(1) = "B_CURRENCY"      <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "화폐"       <%' TextBox 명칭 %>

  arrField(0) = "CURRENCY"        <%' Field명(0)%>
  arrField(1) = "CURRENCY_DESC"        <%' Field명(1)%>

  arrHeader(0) = "화폐"       <%' Header명(0)%>
  arrHeader(1) = "화폐명"      <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetCur(arrRet)
  End If
 End Function


'===========================================================================
Function SetCur(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Cur
  .vspdData.Text = arrRet(0) 
 End With
 Call vspdData_Change(C_Cur,frm1.vspdData.ActiveRow)
End Function


'===========================================================================
 Function  OpenUnit(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "단위"       <%' 팝업 명칭 %>
  arrParam(1) = "B_UNIT_OF_MEASURE"      <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "단위"       <%' TextBox 명칭 %>

  arrField(0) = "UNIT"        <%' Field명(0)%>
  arrField(1) = "UNIT_NM"        <%' Field명(1)%>

  arrHeader(0) = "단위"       <%' Header명(0)%>
  arrHeader(1) = "단위명"      <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetUnit(arrRet)
  End If
 End Function


'===========================================================================
Function SetUnit(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Unit
  .vspdData.Text = arrRet(0)
 End With
End Function


'===========================================================================
 Function  OpenPay_meth(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "결제방법"       <%' 팝업 명칭 %>
  arrParam(1) = "B_MINOR"                  <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)           <%' Code Condition%>
  arrParam(3) = ""             <%' Name Cindition%>
  arrParam(4) = "MAJOR_CD=" & FilterVar("B9004", "''", "S") & ""         <%' Where Condition%>
  arrParam(5) = "결제방법"       <%' TextBox 명칭 %>

  arrField(0) = "MINOR_CD"        <%' Field명(0)%>
  arrField(1) = "MINOR_NM"        <%' Field명(1)%>

  arrHeader(0) = "결제방법"       <%' Header명(0)%>
  arrHeader(1) = "결제방법명"          <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetPay_meth(arrRet)
  End If
 End Function


Function SetPay_meth(Byval arrRet)  
 With frm1
  .vspdData.Col =C_Pay_meth
  .vspdData.Text = arrRet(0)
  .vspdData.Col =C_Pay_meth_NM
  .vspdData.Text = arrRet(1)
 End With
End Function

'===========================================================================
 Function  OpenDeal_type(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "판매유형"       <%' 팝업 명칭 %>
  arrParam(1) = "B_minor"                  <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)           <%' Code Condition%>
  arrParam(3) = ""             <%' Name Cindition%>
  arrParam(4) = "MAJOR_CD=" & FilterVar("S0001", "''", "S") & ""         <%' Where Condition%>
  arrParam(5) = "판매유형"       <%' TextBox 명칭 %>

  arrField(0) = "minor_cd"        <%' Field명(0)%>
  arrField(1) = "minor_nm"        <%' Field명(1)%>

  arrHeader(0) = "판매유형"       <%' Header명(0)%>
  arrHeader(1) = "판매유형명"          <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetDeal_type(arrRet)
  End If
 End Function

'===========================================================================
Function SetDeal_type(Byval arrRet)  
 With frm1
  .vspdData.Col =  C_Deal_type
  .vspdData.Text = arrRet(0)
  .vspdData.Col =  C_Deal_type_nm
  .vspdData.Text = arrRet(1)
 End With
End Function


 Function  OpenItem_cd(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "품목"       <%' 팝업 명칭 %>
  arrParam(1) = "B_ITEM"           <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = ""         <%' Where Condition%>
  arrParam(5) = "품목"       <%' TextBox 명칭 %>

  arrField(0) = "Item_cd"        <%' Field명(0)%>
  arrField(1) = "Item_nm"        <%' Field명(1)%>  
	arrField(2) = "Spec"	        <%' Field명(2)%>
	arrField(3) = "HH" & parent.gColSep & "Basic_Unit"	        <%' Field명(3)%>

  arrHeader(0) = "품목"       <%' Header명(0)%>
  arrHeader(1) = "품목명"       <%' Header명(1)%>
	arrHeader(2) = "규격"       <%' Header명(2)%>
	arrHeader(3) = "단위"       <%' Header명(3)%>
	   
  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetItem_cd(arrRet)
  End If
 End Function

'===========================================================================
Function SetItem_cd(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Item_cd
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_Item_nm
  .vspdData.Text = arrRet(1)
  .vspdData.Col = C_Item_Cd_Spec
  .vspdData.Text = arrRet(2)
  .vspdData.Col = C_Unit
  .vspdData.Text = arrRet(3)  
 End With
End Function


Function  OpenBp_cd(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "고객"       <%' 팝업 명칭 %>
  arrParam(1) = "B_Biz_Partner"           <%' TABLE 명칭 %>
  arrParam(2) = Trim(strCode)       <%' Code Condition%>
  arrParam(3) = ""         <%' Name Cindition%>
  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"         <%' Where Condition%>
  arrParam(5) = "고객"       <%' TextBox 명칭 %>

  arrField(0) = "Bp_cd"        <%' Field명(0)%>
  arrField(1) = "Bp_nm"        <%' Field명(1)%>

  arrHeader(0) = "고객"       <%' Header명(0)%>
  arrHeader(1) = "고객명"       <%' Header명(1)%>

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetBp_cd(arrRet)
  End If
 End Function

Function SetBp_cd(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Bp_cd
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_Bp_nm
  .vspdData.Text = arrRet(1)
 End With
End Function

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
       
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Select Case Col
'         Case  C_StudyOnOffNm
'                iDx = Frm1.vspdData.value
'   	            Frm1.vspdData.Col = C_StudyOnOffCd
'                Frm1.vspdData.value = iDx
         Case Else
    End Select    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
    Call CheckMinNumSpread(frm1.vspdData, Col, Row)		
    
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)

    Call SetPopupMenuItemInf("1111111111")    
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
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub


'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Occurs when the user clicks the left mouse button while the pointer is in a cell
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    


'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
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
	
	arrHeader(0) = "공장"		
	arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
		frm1.txtPlantCd.focus
	End If	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>작업장별품목별수량조회</font></td>
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
									<TD CLASS="TD656" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>									
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
							<TABLE <%=LR_SPACE_TYPE_30%>>
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"      TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMode"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"         TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"    TAG="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtPlantCd"     TAG="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

